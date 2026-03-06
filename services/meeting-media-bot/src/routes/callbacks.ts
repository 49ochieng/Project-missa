/**
 * Graph Communications API callback handler
 * Handles notification webhooks from Microsoft Graph for call events
 */

import { Router, Request, Response } from "express";
import { updateCallState, CallResource, getJoinUrl } from "../graph/callManager";
import { startTranscriptPolling, stopTranscriptPolling, isPollingActive } from "../graph/transcriptPoller";
import { isAcsConfigured, startAcsTranscription } from "../acs/acsTranscriber";
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
 * Handle call established - start transcript polling, with ACS fallback
 */
async function handleCallEstablished(callId: string, serverCallId?: string): Promise<void> {
  console.log(`[Callbacks] Call established: ${callId}`);

  await sendStatusToProjectMissa(callId, "joined");

  // Get the join URL stored when we initiated the call
  const joinUrl = getJoinUrl(callId);
  if (!joinUrl) {
    console.warn(`[Callbacks] No join URL found for call ${callId}, skipping transcript polling`);
    return;
  }

  // Start live transcript polling via Teams Graph API
  const result = await startTranscriptPolling(callId, joinUrl);
  if (result.success) {
    console.log(`[Callbacks] Transcript polling started: ${result.message}`);
    await sendStatusToProjectMissa(callId, "transcribing");
  } else {
    console.warn(`[Callbacks] Transcript polling unavailable: ${result.message}`);
    await sendStatusToProjectMissa(callId, "joined", result.message);
  }

  // ACS fallback: if ACS is configured and transcript polling didn't find a transcript,
  // try ACS transcription after 30 seconds
  if (isAcsConfigured() && serverCallId) {
    setTimeout(async () => {
      // Check if transcript polling has found anything
      if (!isPollingActive(callId)) {
        console.log(`[Callbacks] Transcript polling inactive for ${callId} — skipping ACS fallback`);
        return;
      }

      console.log(`[Callbacks] Checking if ACS fallback is needed for ${callId}...`);
      const acsResult = await startAcsTranscription(callId, serverCallId);
      if (acsResult.success) {
        console.log(`[Callbacks] ACS transcription fallback started for ${callId}`);
        await sendStatusToProjectMissa(callId, "transcribing_acs");
      } else {
        console.warn(`[Callbacks] ACS fallback failed: ${acsResult.message}`);
      }
    }, 30000);
  }
}

/**
 * Handle call terminated - clean up
 */
async function handleCallTerminated(callId: string): Promise<void> {
  console.log(`[Callbacks] Call terminated: ${callId}`);

  // Stop transcript polling
  stopTranscriptPolling(callId);

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

    // Extract serverCallId if available (needed for ACS fallback)
    const serverCallId = (callData as unknown as Record<string, unknown>).serverCallId as string | undefined;

    // Handle state transitions
    switch (state) {
      case "established":
        await handleCallEstablished(callId, serverCallId);
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
