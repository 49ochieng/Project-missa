/**
 * ACS Call Automation Transcription Module
 * Provides real-time transcription via Azure Communication Services as a fallback
 * when Teams native transcription isn't enabled by the meeting organizer.
 *
 * Prerequisites:
 * - ACS resource provisioned in Azure
 * - Azure Cognitive Services (Speech) endpoint
 * - ACS_CONNECTION_STRING and COGNITIVE_SERVICES_ENDPOINT env vars
 *
 * Flow:
 * 1. When a Graph call is established and transcript polling finds no transcript,
 *    this module starts an ACS transcription session
 * 2. A WebSocket server receives real-time transcription events from ACS
 * 3. Parsed transcript chunks are forwarded to Project Missa
 */

import { CallAutomationClient } from "@azure/communication-call-automation";
import { WebSocketServer, WebSocket } from "ws";
import { getConfig } from "../config";

interface AcsTranscriptionSession {
  callConnectionId: string;
  callId: string; // Graph call ID for correlation
  active: boolean;
}

// Active ACS transcription sessions: Graph callId → session
const activeSessions = new Map<string, AcsTranscriptionSession>();

let acsClient: CallAutomationClient | null = null;
let wss: WebSocketServer | null = null;
let wsPort: number = 0;

/**
 * Check if ACS transcription is configured and available
 */
export function isAcsConfigured(): boolean {
  const config = getConfig();
  return !!(config.acsConnectionString && config.cognitiveServicesEndpoint);
}

/**
 * Initialize the ACS client and WebSocket server
 * Call this once at startup if ACS is configured
 */
export function initializeAcs(port: number = 8081): void {
  const config = getConfig();

  if (!config.acsConnectionString) {
    console.log("[ACS] Not configured — ACS_CONNECTION_STRING not set");
    return;
  }

  if (!config.cognitiveServicesEndpoint) {
    console.log("[ACS] Not configured — COGNITIVE_SERVICES_ENDPOINT not set");
    return;
  }

  try {
    acsClient = new CallAutomationClient(config.acsConnectionString);
    console.log("[ACS] CallAutomationClient initialized");

    // Start WebSocket server for receiving transcription events
    wss = new WebSocketServer({ port });
    wsPort = port;

    wss.on("connection", (ws: WebSocket) => {
      console.log("[ACS-WS] Transcription client connected");

      ws.on("message", (data: Buffer) => {
        handleTranscriptionMessage(data);
      });

      ws.on("close", () => {
        console.log("[ACS-WS] Transcription client disconnected");
      });

      ws.on("error", (err) => {
        console.error("[ACS-WS] WebSocket error:", err.message);
      });
    });

    wss.on("error", (err) => {
      console.error("[ACS-WS] Server error:", err.message);
    });

    console.log(`[ACS] WebSocket transcription server started on port ${port}`);
  } catch (error) {
    console.error("[ACS] Failed to initialize:", error instanceof Error ? error.message : error);
  }
}

/**
 * Parse and handle incoming transcription messages from ACS WebSocket
 */
function handleTranscriptionMessage(data: Buffer): void {
  try {
    const decoder = new TextDecoder();
    const json = JSON.parse(decoder.decode(data));

    if (json.kind === "TranscriptionMetadata") {
      const metadata = json.transcriptionMetadata;
      console.log(`[ACS] Transcription metadata — locale: ${metadata?.locale}, connectionId: ${metadata?.callConnectionId}`);
    } else if (json.kind === "TranscriptionData") {
      const td = json.transcriptionData;

      if (!td || !td.text) return;

      // Only forward final results
      if (td.resultStatus !== "Final") return;

      const speaker = td.participantRawID || "Unknown";
      const text = td.text;
      const offsetMs = td.offset ? Math.floor(td.offset / 10000) : 0; // Convert from ticks to ms
      const durationMs = td.duration ? Math.floor(td.duration / 10000) : 0;

      console.log(`[ACS] Transcript: [${speaker}] ${text}`);

      // Find which call this belongs to by checking active sessions
      for (const [callId, session] of activeSessions) {
        if (session.active) {
          sendChunkToProjectMissa(callId, speaker, text, offsetMs, durationMs);
          break; // In practice, match by callConnectionId if multiple sessions
        }
      }
    }
  } catch (error) {
    console.error("[ACS] Error parsing transcription message:", error instanceof Error ? error.message : error);
  }
}

/**
 * Send a transcript chunk to Project Missa
 */
async function sendChunkToProjectMissa(
  callId: string,
  speaker: string,
  text: string,
  offsetMs: number,
  durationMs: number
): Promise<void> {
  const config = getConfig();

  try {
    const response = await fetch(`${config.projectMissaUrl}/api/meeting-transcripts/chunk`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shared-Secret": config.sharedSecret,
      },
      body: JSON.stringify({
        callId,
        text: `[${speaker}]: ${text}`,
        speakerId: speaker,
        timestamp: new Date().toISOString(),
        offsetMs,
        durationMs,
        isFinal: true,
        source: "acs_transcription",
      }),
    });

    if (!response.ok) {
      console.error(`[ACS] Failed to send chunk to Project Missa: ${response.status}`);
    }
  } catch (error) {
    console.error("[ACS] Error sending chunk:", error instanceof Error ? error.message : error);
  }
}

/**
 * Start ACS transcription for a call
 * Uses connectCall with ServerCallLocator to connect to an existing call,
 * then starts transcription via the WebSocket transport.
 *
 * @param callId - Graph Communications call ID (for correlation)
 * @param serverCallId - ACS server call ID (from call events)
 * @returns success/failure
 */
export async function startAcsTranscription(
  callId: string,
  serverCallId: string
): Promise<{ success: boolean; message: string }> {
  if (!acsClient) {
    return { success: false, message: "ACS not initialized" };
  }

  const config = getConfig();

  try {
    const callbackUri = `${config.botEndpoint}/api/acs/callbacks`;
    const transportUrl = `wss://${new URL(config.botEndpoint).hostname}:${wsPort}/ws/transcription`;

    console.log(`[ACS] Starting transcription for call ${callId} (serverCallId: ${serverCallId})`);

    const callLocator = { kind: "serverCallLocator" as const, id: serverCallId };

    const connectResult = await acsClient.connectCall(callLocator, callbackUri, {
      callIntelligenceOptions: {
        cognitiveServicesEndpoint: config.cognitiveServicesEndpoint!,
      },
      transcriptionOptions: {
        transportUrl,
        transportType: "websocket",
        locale: "en-US",
        startTranscription: true,
      },
    });

    const callConnectionId = connectResult.callConnectionProperties.callConnectionId;

    if (callConnectionId) {
      activeSessions.set(callId, {
        callConnectionId,
        callId,
        active: true,
      });

      console.log(`[ACS] Connected and transcribing — callConnectionId: ${callConnectionId}`);
      return { success: true, message: "ACS transcription started" };
    }

    return { success: false, message: "No callConnectionId returned from connectCall" };
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    console.error(`[ACS] Failed to start transcription: ${msg}`);
    return { success: false, message: msg };
  }
}

/**
 * Stop ACS transcription for a call
 */
export async function stopAcsTranscription(callId: string): Promise<void> {
  const session = activeSessions.get(callId);
  if (!session) return;

  session.active = false;
  activeSessions.delete(callId);

  if (acsClient && session.callConnectionId) {
    try {
      const connection = acsClient.getCallConnection(session.callConnectionId);
      const callMedia = connection.getCallMedia();
      await callMedia.stopTranscription();
      console.log(`[ACS] Stopped transcription for call ${callId}`);
    } catch (error) {
      console.error(`[ACS] Error stopping transcription: ${error instanceof Error ? error.message : error}`);
    }
  }
}

/**
 * Check if ACS transcription is active for a call
 */
export function isAcsTranscriptionActive(callId: string): boolean {
  return activeSessions.has(callId) && (activeSessions.get(callId)?.active ?? false);
}

/**
 * Shutdown ACS resources
 */
export function shutdownAcs(): void {
  if (wss) {
    wss.close();
    console.log("[ACS] WebSocket server closed");
  }
  activeSessions.clear();
}
