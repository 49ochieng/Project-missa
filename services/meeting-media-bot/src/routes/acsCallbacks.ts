/**
 * ACS Call Automation callback routes
 * Handles event notifications from Azure Communication Services
 */

import { Router, Request, Response } from "express";

const router = Router();

/**
 * POST /api/acs/callbacks
 * Webhook for ACS Call Automation events (CallConnected, TranscriptionStarted, etc.)
 */
router.post("/callbacks", (req: Request, res: Response) => {
  const events = req.body;

  if (!Array.isArray(events)) {
    console.log("[ACS-Callback] Received non-array event:", typeof events);
    res.status(200).send();
    return;
  }

  for (const event of events) {
    const eventType = event.type || event.eventType;
    console.log(`[ACS-Callback] Event: ${eventType}`);

    switch (eventType) {
      case "Microsoft.Communication.CallConnected":
        console.log("[ACS-Callback] ACS call connected");
        break;

      case "Microsoft.Communication.CallDisconnected":
        console.log("[ACS-Callback] ACS call disconnected");
        break;

      case "Microsoft.Communication.TranscriptionStarted":
        console.log("[ACS-Callback] Transcription started");
        break;

      case "Microsoft.Communication.TranscriptionStopped":
        console.log("[ACS-Callback] Transcription stopped");
        break;

      case "Microsoft.Communication.TranscriptionFailed":
        console.error("[ACS-Callback] Transcription failed:", event.data?.resultInformation);
        break;

      default:
        console.log(`[ACS-Callback] Unhandled event: ${eventType}`);
    }
  }

  res.status(200).send();
});

export default router;
