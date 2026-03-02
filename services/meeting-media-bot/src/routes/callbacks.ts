/**
 * Graph Communications API callback handler
 * Handles notification webhooks from Microsoft Graph for call events
 */

import { Router, Request, Response } from "express";
import { updateCallState, getCall, leaveCall, CallResource } from "../graph/callManager";
import { createTranscriber, removeTranscriber, TranscriptionChunk } from "../speech/transcriber";
import { getConfig } from "../config";

const router = Router();

/**
 * Notification payload from Graph
 */
interface GraphNotification {
  value: Array<{
    changeType: string;
    resource: string;
    resourceData: CallResource;
    clientState?: string;
  }>;
}

/**
 * Send transcription chunk to Project Missa
 */
async function sendChunkToProjectMissa(
  callId: string,
  chunk: TranscriptionChunk
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
        text: chunk.text,
        speakerId: chunk.speakerId,
        timestamp: chunk.timestamp.toISOString(),
        offsetMs: chunk.offsetMs,
        durationMs: chunk.durationMs,
        isFinal: chunk.isFinal,
        source: "azure_speech",
      }),
    });

    if (!response.ok) {
      console.error(
        `[Callbacks] Failed to send chunk to Project Missa: ${response.status}`
      );
    }
  } catch (error) {
    console.error(`[Callbacks] Error sending chunk to Project Missa:`, error);
  }
}

/**
 * Send meeting status update to Project Missa
 */
async function sendStatusToProjectMissa(
  callId: string,
  status: string,
  error?: string
): Promise<void> {
  const config = getConfig();

  try {
    const response = await fetch(`${config.projectMissaUrl}/api/meeting-capture/status`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shared-Secret": config.sharedSecret,
      },
      body: JSON.stringify({
        callId,
        status,
        error,
        timestamp: new Date().toISOString(),
      }),
    });

    if (!response.ok) {
      console.error(
        `[Callbacks] Failed to send status to Project Missa: ${response.status}`
      );
    }
  } catch (error) {
    console.error(`[Callbacks] Error sending status to Project Missa:`, error);
  }
}

/**
 * Handle call established - start transcription
 */
async function handleCallEstablished(callId: string): Promise<void> {
  console.log(`[Callbacks] Call established: ${callId}`);

  // Create transcriber with event handlers
  const transcriber = createTranscriber(callId, {
    onTranscriptionChunk: (chunk) => {
      sendChunkToProjectMissa(callId, chunk);
    },
    onError: (error) => {
      console.error(`[Callbacks] Transcription error for call ${callId}:`, error);
      sendStatusToProjectMissa(callId, "transcription_error", error.message);
    },
    onSessionStarted: () => {
      console.log(`[Callbacks] Transcription session started for call ${callId}`);
      sendStatusToProjectMissa(callId, "transcribing");
    },
    onSessionStopped: () => {
      console.log(`[Callbacks] Transcription session stopped for call ${callId}`);
    },
  });

  try {
    // Initialize with push stream (audio will be fed from media processing)
    // Note: For now, this creates the stream. Media platform integration will feed audio.
    transcriber.initializePushStream();
    await transcriber.start();
    
    await sendStatusToProjectMissa(callId, "joined");
  } catch (error) {
    console.error(`[Callbacks] Failed to start transcription:`, error);
    await sendStatusToProjectMissa(
      callId,
      "transcription_error",
      error instanceof Error ? error.message : "Unknown error"
    );
  }
}

/**
 * Handle call terminated - clean up
 */
async function handleCallTerminated(callId: string): Promise<void> {
  console.log(`[Callbacks] Call terminated: ${callId}`);

  // Clean up transcriber
  removeTranscriber(callId);

  // Notify Project Missa
  await sendStatusToProjectMissa(callId, "ended");
}

/**
 * POST /api/calls/callback
 * Webhook endpoint for Graph Communications notifications
 */
router.post("/callback", async (req: Request, res: Response) => {
  console.log(`[Callbacks] Received notification`);

  // Handle validation request from Graph
  const validationToken = req.query.validationToken as string;
  if (validationToken) {
    console.log(`[Callbacks] Validation request, returning token`);
    res.type("text/plain").status(200).send(validationToken);
    return;
  }

  // Parse notification
  const notification = req.body as GraphNotification;

  if (!notification.value || notification.value.length === 0) {
    console.warn(`[Callbacks] Empty notification received`);
    res.status(200).send();
    return;
  }

  // Process each notification
  for (const item of notification.value) {
    const callData = item.resourceData;

    if (!callData || !callData.id) {
      console.warn(`[Callbacks] Notification missing call data`);
      continue;
    }

    const callId = callData.id;
    const state = callData.state;

    console.log(`[Callbacks] Call ${callId} state: ${state}`);

    // Update local call state
    updateCallState(callId, callData);

    // Handle state transitions
    switch (state) {
      case "established":
        await handleCallEstablished(callId);
        break;

      case "terminated":
        await handleCallTerminated(callId);
        break;

      case "establishing":
        console.log(`[Callbacks] Call ${callId} is establishing...`);
        await sendStatusToProjectMissa(callId, "joining");
        break;

      case "terminating":
        console.log(`[Callbacks] Call ${callId} is terminating...`);
        break;

      default:
        console.log(`[Callbacks] Call ${callId} state: ${state}`);
    }
  }

  // Always return 200 to acknowledge receipt
  res.status(200).send();
});

/**
 * GET /api/calls/callback (health check)
 */
router.get("/callback", (_req: Request, res: Response) => {
  res.json({ status: "ok", endpoint: "calls-callback" });
});

export default router;
