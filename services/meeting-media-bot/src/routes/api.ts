/**
 * API routes for meeting-media-bot
 * Called by Project Missa to start/stop meeting capture
 */

import { Router, Request, Response, NextFunction } from "express";
import { joinMeeting, leaveCall, getCall, getActiveCalls, isCallActive } from "../graph/callManager";
import { stopTranscriptPolling } from "../graph/transcriptPoller";
import { getConfig } from "../config";

const router = Router();

/**
 * Middleware to verify shared secret from Project Missa
 */
function verifySharedSecret(req: Request, res: Response, next: NextFunction): void {
  const config = getConfig();
  const providedSecret = req.headers["x-shared-secret"] as string;

  if (!providedSecret || providedSecret !== config.sharedSecret) {
    console.warn(`[API] Unauthorized request to ${req.path}`);
    res.status(401).json({ error: "Unauthorized" });
    return;
  }

  next();
}

/**
 * POST /api/meetings/join
 * Join a Teams meeting and start capturing
 * 
 * Body: { joinUrl: string, meetingId?: string, displayName?: string }
 */
router.post("/meetings/join", verifySharedSecret, async (req: Request, res: Response) => {
  const { joinUrl, meetingId, displayName } = req.body;

  console.log(`[API] Join meeting request: ${joinUrl?.substring(0, 50)}...`);

  if (!joinUrl) {
    res.status(400).json({ error: "joinUrl is required" });
    return;
  }

  try {
    const result = await joinMeeting(joinUrl, displayName);

    if (!result.success) {
      console.error(`[API] Failed to join meeting: ${result.error}`);
      res.status(500).json({
        success: false,
        error: result.error,
      });
      return;
    }

    console.log(`[API] Successfully joined meeting, callId: ${result.callId}`);

    res.json({
      success: true,
      callId: result.callId,
      meetingId,
      status: "joining",
    });
  } catch (error) {
    console.error(`[API] Error joining meeting:`, error);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

/**
 * POST /api/meetings/leave
 * Leave a meeting and stop capturing
 * 
 * Body: { callId: string }
 */
router.post("/meetings/leave", verifySharedSecret, async (req: Request, res: Response) => {
  const { callId } = req.body;

  console.log(`[API] Leave meeting request: ${callId}`);

  if (!callId) {
    res.status(400).json({ error: "callId is required" });
    return;
  }

  try {
    // Stop transcript polling first
    stopTranscriptPolling(callId);

    // Leave the call
    const result = await leaveCall(callId);

    if (!result.success) {
      console.error(`[API] Failed to leave meeting: ${result.error}`);
      res.status(500).json({
        success: false,
        error: result.error,
      });
      return;
    }

    console.log(`[API] Successfully left meeting: ${callId}`);

    res.json({
      success: true,
      callId,
      status: "left",
    });
  } catch (error) {
    console.error(`[API] Error leaving meeting:`, error);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

/**
 * GET /api/meetings/:callId/status
 * Get current status of a meeting capture
 */
router.get("/meetings/:callId/status", verifySharedSecret, (req: Request, res: Response) => {
  const { callId } = req.params;

  console.log(`[API] Status request for call: ${callId}`);

  const call = getCall(callId);

  if (!call) {
    res.status(404).json({ error: "Call not found" });
    return;
  }

  res.json({
    callId,
    state: call.state,
    isActive: isCallActive(callId),
    myParticipantId: call.myParticipantId,
  });
});

/**
 * GET /api/meetings/active
 * List all active meeting captures
 */
router.get("/meetings/active", verifySharedSecret, (_req: Request, res: Response) => {
  console.log(`[API] List active meetings request`);

  const activeCalls = getActiveCalls();
  const meetings = [];

  for (const [callId, call] of activeCalls) {
    meetings.push({
      callId,
      state: call.state,
      isActive: isCallActive(callId),
    });
  }

  res.json({
    count: meetings.length,
    meetings,
  });
});

/**
 * POST /api/meetings/:callId/stop-transcription
 * Stop transcript polling without leaving the meeting
 */
router.post("/meetings/:callId/stop-transcription", verifySharedSecret, (req: Request, res: Response) => {
  const { callId } = req.params;

  console.log(`[API] Stop transcription request for call: ${callId}`);

  stopTranscriptPolling(callId);

  res.json({
    success: true,
    callId,
    status: "transcription_stopped",
  });
});

/**
 * Health check endpoint (no auth required)
 */
router.get("/health", (_req: Request, res: Response) => {
  res.json({
    status: "ok",
    service: "meeting-media-bot",
    timestamp: new Date().toISOString(),
  });
});

export default router;
