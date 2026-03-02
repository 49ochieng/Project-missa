/**
 * Meeting API Routes for Project Missa
 * Receives transcript chunks and status updates from meeting-media-bot
 */

import { Router, Request, Response, NextFunction } from "express";
import { IDatabase } from "../storage/database";
import { TranscriptSource } from "../storage/meetingTypes";
import { AppConfig } from "../utils/config";
import type { ILogger } from "@microsoft/teams.common";

let dbInstance: IDatabase | null = null;
let configInstance: AppConfig | null = null;
let loggerInstance: ILogger | null = null;

/**
 * Initialize the meeting routes with dependencies
 */
export function initializeMeetingRoutes(
  db: IDatabase,
  config: AppConfig,
  logger: ILogger
): void {
  dbInstance = db;
  configInstance = config;
  loggerInstance = logger;
}

const router = Router();

/**
 * Middleware to verify shared secret from meeting-media-bot
 */
function verifySharedSecret(req: Request, res: Response, next: NextFunction): void {
  if (!configInstance) {
    res.status(500).json({ error: "Server not initialized" });
    return;
  }

  const providedSecret = req.headers["x-shared-secret"] as string;
  const expectedSecret = configInstance.meetingMediaBotSharedSecret;

  if (!providedSecret || !expectedSecret || providedSecret !== expectedSecret) {
    loggerInstance?.warn(`Unauthorized request to ${req.path}`);
    res.status(401).json({ error: "Unauthorized" });
    return;
  }

  next();
}

/**
 * POST /api/meeting-transcripts/chunk
 * Receive transcript chunk from meeting-media-bot
 * 
 * Body: {
 *   callId: string,
 *   text: string,
 *   speakerId?: string,
 *   timestamp: string (ISO),
 *   offsetMs: number,
 *   durationMs: number,
 *   isFinal: boolean,
 *   source: "azure_speech" | "graph_transcript"
 * }
 */
router.post("/meeting-transcripts/chunk", verifySharedSecret, async (req: Request, res: Response) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }

  const {
    callId,
    text,
    speakerId,
    timestamp,
    source,
  } = req.body;

  loggerInstance?.debug(`Received transcript chunk for call ${callId}: "${text?.substring(0, 50)}..."`);

  if (!callId || !text) {
    res.status(400).json({ error: "callId and text are required" });
    return;
  }

  try {
    // Map incoming source to database TranscriptSource type
    let dbSource: TranscriptSource = "speech"; // default
    if (source === "azure_speech" || source === "speech") {
      dbSource = "speech";
    } else if (source === "graph_transcript" || source === "graphTranscript") {
      dbSource = "graphTranscript";
    }
    
    // Ensure meeting exists before inserting transcript chunks
    const existingMeeting = await dbInstance.getMeeting(callId);
    if (!existingMeeting) {
      // Create meeting record if it doesn't exist
      await dbInstance.upsertMeeting({
        meetingId: callId,
        conversationId: "api-upload", // Placeholder for API-uploaded transcripts
        joinUrl: `api://meeting/${callId}`, // Placeholder URL for API-uploaded meetings
        title: `Meeting ${callId}`,
        organizerAadId: speakerId || "unknown",
      });
      loggerInstance?.debug(`Created meeting record for ${callId}`);
    }
    
    // Look up meeting by call ID (stored in meetingJoinUrl or a new field)
    // For now, we'll use callId as meetingId directly
    const chunkId = await dbInstance.appendTranscriptChunk({
      meetingId: callId,
      speaker: speakerId || "Unknown",
      text,
      timestampUtc: timestamp || new Date().toISOString(),
      confidence: 1.0, // Azure Speech provides high confidence final results
      source: dbSource,
    });

    loggerInstance?.debug(`Stored transcript chunk ${chunkId} for meeting ${callId}`);

    res.json({
      success: true,
      chunkId,
      meetingId: callId,
    });
  } catch (error) {
    loggerInstance?.error(`Error storing transcript chunk: ${error}`);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

/**
 * POST /api/meeting-capture/status
 * Receive status update from meeting-media-bot
 * 
 * Body: {
 *   callId: string,
 *   status: "joining" | "joined" | "transcribing" | "transcription_error" | "ended",
 *   error?: string,
 *   timestamp: string (ISO)
 * }
 */
router.post("/meeting-capture/status", verifySharedSecret, async (req: Request, res: Response) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }

  const { callId, status, error } = req.body;

  loggerInstance?.info(`Meeting ${callId} status update: ${status}${error ? ` (error: ${error})` : ""}`);

  if (!callId || !status) {
    res.status(400).json({ error: "callId and status are required" });
    return;
  }

  try {
    // Map meeting-media-bot status to our MeetingStatus enum
    let meetingStatus: "joining" | "recording" | "ended" | "failed" | "cancelled";
    switch (status) {
      case "joining":
        meetingStatus = "joining";
        break;
      case "joined":
      case "transcribing":
        meetingStatus = "recording";
        break;
      case "transcription_error":
        meetingStatus = "failed";
        break;
      case "ended":
        meetingStatus = "ended";
        break;
      default:
        meetingStatus = "recording";
    }

    await dbInstance.updateMeetingStatus({
      meetingId: callId,
      status: meetingStatus,
      endedAt: meetingStatus === "ended" ? new Date().toISOString() : undefined,
    });

    loggerInstance?.debug(`Updated meeting ${callId} status to ${meetingStatus}`);

    res.json({
      success: true,
      meetingId: callId,
      mappedStatus: meetingStatus,
    });
  } catch (error) {
    loggerInstance?.error(`Error updating meeting status: ${error}`);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

/**
 * GET /api/meeting-transcripts/:meetingId
 * Get full transcript for a meeting
 */
router.get("/meeting-transcripts/:meetingId", async (req: Request, res: Response) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }

  const meetingId = req.params.meetingId as string;

  loggerInstance?.debug(`Retrieving transcript for meeting ${meetingId}`);

  try {
    const result = await dbInstance.getTranscriptByMeetingId(meetingId);

    if (!result) {
      res.status(404).json({ error: "Meeting not found" });
      return;
    }

    // Build full text from chunks
    const fullText = result.chunks
      .map((chunk) => `[${chunk.speaker}]: ${chunk.text}`)
      .join("\n");

    res.json({
      meetingId: result.meetingId,
      title: result.title,
      startedAt: result.startedAt,
      endedAt: result.endedAt,
      chunkCount: result.totalChunks,
      fullText,
      chunks: result.chunks.map((chunk) => ({
        speaker: chunk.speaker,
        speakerAadId: chunk.speakerAadId,
        text: chunk.text,
        timestamp: chunk.timestampUtc,
        confidence: chunk.confidence,
        source: chunk.source,
      })),
      participants: result.participants,
    });
  } catch (error) {
    loggerInstance?.error(`Error retrieving transcript: ${error}`);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

/**
 * GET /api/meetings/:meetingId
 * Get meeting details
 */
router.get("/meetings/:meetingId", async (req: Request, res: Response) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }

  const meetingId = req.params.meetingId as string;

  try {
    const meeting = await dbInstance.getMeeting(meetingId);

    if (!meeting) {
      res.status(404).json({ error: "Meeting not found" });
      return;
    }

    res.json(meeting);
  } catch (error) {
    loggerInstance?.error(`Error retrieving meeting: ${error}`);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

/**
 * Health check endpoint (no auth required)
 */
router.get("/health", (_req: Request, res: Response) => {
  res.json({
    status: "ok",
    service: "project-missa",
    meetingApiEnabled: true,
    timestamp: new Date().toISOString(),
  });
});

export default router;
