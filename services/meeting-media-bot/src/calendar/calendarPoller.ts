/**
 * Calendar Poller — Tier 2 Auto-Join
 *
 * Polls the configured user's Microsoft 365 calendar every 60 seconds.
 * When a Teams meeting is starting (within the next 2 minutes) or has just
 * started (within the last 5 minutes), automatically joins it via Graph API.
 * Auto-leaves when the meeting end time has passed.
 *
 * Requires:
 *   - AUTO_JOIN_USER_EMAIL or AUTO_JOIN_USER_OBJECT_ID in .env
 *   - Calendars.Read.All Application permission on the Azure AD app
 */

import { getGraphClient } from "../graph/graphClient";
import { joinMeeting, leaveCall, isCallActive } from "../graph/callManager";
import { getConfig } from "../config";

interface CalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  onlineMeeting: { joinUrl: string } | null;
}

interface TrackedMeeting {
  eventId: string;
  callId: string;
  subject: string;
  joinUrl: string;
  startTime: Date;
  endTime: Date;
}

// Active meetings we auto-joined
const trackedMeetings = new Map<string, TrackedMeeting>();

// Events we've already tried joining (to avoid repeated attempts on failure)
const attemptedEventIds = new Set<string>();

let pollerInterval: NodeJS.Timeout | null = null;
let resolvedUserId: string | null = null;

/**
 * Resolve user object ID from email if needed
 */
async function resolveUserId(): Promise<string | null> {
  if (resolvedUserId) return resolvedUserId;

  const config = getConfig();

  // Use object ID directly if configured (preferred — no extra API call)
  if (config.autoJoinUserObjectId) {
    resolvedUserId = config.autoJoinUserObjectId;
    console.log(`[CalendarPoller] Using configured user ID: ${resolvedUserId}`);
    return resolvedUserId;
  }

  // Resolve from email
  if (config.autoJoinUserEmail) {
    const client = getGraphClient();
    const response = await client.request<{ id: string }>(
      "GET",
      `/v1.0/users/${encodeURIComponent(config.autoJoinUserEmail)}?$select=id`
    );

    if (response.success && response.data?.id) {
      resolvedUserId = response.data.id;
      console.log(`[CalendarPoller] Resolved ${config.autoJoinUserEmail} → ${resolvedUserId}`);
      return resolvedUserId;
    }

    console.error(`[CalendarPoller] Could not resolve user ID for ${config.autoJoinUserEmail}`);
  }

  return null;
}

/**
 * Fetch upcoming online meetings from user's calendar (next 90 minutes)
 */
async function fetchUpcomingMeetings(userId: string): Promise<CalendarEvent[]> {
  const client = getGraphClient();
  const now = new Date();
  const windowEnd = new Date(now.getTime() + 90 * 60 * 1000); // +90 minutes

  const startDateTime = now.toISOString();
  const endDateTime = windowEnd.toISOString();

  const query = [
    `startDateTime=${encodeURIComponent(startDateTime)}`,
    `endDateTime=${encodeURIComponent(endDateTime)}`,
    `$select=id,subject,start,end,onlineMeeting`,
    `$filter=isOnlineMeeting eq true`,
    `$orderby=start/dateTime`,
    `$top=10`,
  ].join("&");

  const response = await client.request<{ value: CalendarEvent[] }>(
    "GET",
    `/v1.0/users/${userId}/calendarView?${query}`
  );

  if (!response.success || !response.data?.value) {
    return [];
  }

  return response.data.value.filter((e) => e.onlineMeeting?.joinUrl);
}

/**
 * Notify Project Missa about an auto-join (best-effort)
 */
async function notifyProjectMissa(callId: string, status: string, subject: string): Promise<void> {
  const config = getConfig();
  try {
    await fetch(`${config.projectMissaUrl}/api/meeting-capture/status`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Shared-Secret": config.sharedSecret,
      },
      body: JSON.stringify({
        callId,
        status,
        error: status === "joined" ? `Auto-joined calendar meeting: "${subject}"` : undefined,
        timestamp: new Date().toISOString(),
      }),
    });
  } catch {
    // Non-critical — don't crash the poller
  }
}

/**
 * Main polling loop
 */
async function pollCalendar(): Promise<void> {
  const userId = await resolveUserId();
  if (!userId) return;

  const now = new Date();

  try {
    const events = await fetchUpcomingMeetings(userId);

    for (const event of events) {
      const startTime = new Date(event.start.dateTime + (event.start.dateTime.endsWith("Z") ? "" : "Z"));
      const endTime = new Date(event.end.dateTime + (event.end.dateTime.endsWith("Z") ? "" : "Z"));
      const joinUrl = event.onlineMeeting!.joinUrl;

      // ── Auto-join: meeting starts within 2 min ahead OR started up to 5 min ago ──
      const msUntilStart = startTime.getTime() - now.getTime();
      const mssinceStart = now.getTime() - startTime.getTime();
      const shouldJoin =
        msUntilStart <= 2 * 60 * 1000 &&   // Starting within 2 min
        mssinceStart <= 5 * 60 * 1000 &&   // Not more than 5 min past start
        !trackedMeetings.has(event.id) &&   // Not already joined
        !attemptedEventIds.has(event.id);  // Not previously attempted (failed)

      if (shouldJoin) {
        console.log(`[CalendarPoller] Auto-joining: "${event.subject}" (starts ${startTime.toISOString()})`);
        attemptedEventIds.add(event.id);

        const result = await joinMeeting(joinUrl);

        if (result.success && result.callId) {
          trackedMeetings.set(event.id, {
            eventId: event.id,
            callId: result.callId,
            subject: event.subject,
            joinUrl,
            startTime,
            endTime,
          });
          console.log(`[CalendarPoller] Joined "${event.subject}", callId: ${result.callId}`);
          await notifyProjectMissa(result.callId, "joined", event.subject);
        } else {
          console.error(`[CalendarPoller] Failed to join "${event.subject}": ${result.error}`);
          // Remove from attempted set after 5 min so it retries
          setTimeout(() => attemptedEventIds.delete(event.id), 5 * 60 * 1000);
        }
      }
    }

    // ── Auto-leave: meeting end time has passed by 5+ minutes ──
    for (const [eventId, tracked] of trackedMeetings.entries()) {
      const msAfterEnd = now.getTime() - tracked.endTime.getTime();
      if (msAfterEnd >= 5 * 60 * 1000) {
        console.log(`[CalendarPoller] Meeting "${tracked.subject}" ended — leaving call ${tracked.callId}`);
        try {
          await leaveCall(tracked.callId);
          await notifyProjectMissa(tracked.callId, "ended", tracked.subject);
        } catch (err) {
          console.error(`[CalendarPoller] Error leaving call:`, err);
        }
        trackedMeetings.delete(eventId);
      } else if (!isCallActive(tracked.callId)) {
        // Call was terminated externally (meeting ended early)
        console.log(`[CalendarPoller] Call ${tracked.callId} terminated externally, removing from tracking`);
        await notifyProjectMissa(tracked.callId, "ended", tracked.subject);
        trackedMeetings.delete(eventId);
      }
    }
  } catch (error) {
    console.error("[CalendarPoller] Poll error:", error);
  }
}

/**
 * Start the calendar auto-join poller
 */
export function startCalendarPoller(): void {
  const config = getConfig();
  const intervalMs = config.calendarPollIntervalMs;

  console.log(`[CalendarPoller] Starting — polling every ${intervalMs / 1000}s`);

  // Resolve user ID on first poll (async, non-blocking)
  pollCalendar().catch((err) => console.error("[CalendarPoller] Initial poll error:", err));

  pollerInterval = setInterval(() => {
    pollCalendar().catch((err) => console.error("[CalendarPoller] Poll error:", err));
  }, intervalMs);
}

/**
 * Stop the calendar poller
 */
export function stopCalendarPoller(): void {
  if (pollerInterval) {
    clearInterval(pollerInterval);
    pollerInterval = null;
    console.log("[CalendarPoller] Stopped");
  }
}

/**
 * Get currently auto-joined meetings (for status/debugging)
 */
export function getTrackedMeetings(): TrackedMeeting[] {
  return Array.from(trackedMeetings.values());
}
