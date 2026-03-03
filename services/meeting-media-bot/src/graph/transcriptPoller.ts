/**
 * Live transcript polling using Microsoft Teams Graph Transcript API
 * Polls for Teams built-in transcript segments and forwards them to Project Missa
 */

import { getGraphClient } from "./graphClient";
import { getConfig } from "../config";

interface OnlineMeeting {
  id: string;
  joinWebUrl: string;
  subject?: string;
  organizer?: { user?: { id: string } };
}

interface Transcript {
  id: string;
  createdDateTime?: string;
  endDateTime?: string;
}

interface ActivePoller {
  callId: string;
  joinUrl: string;
  organizerUserId: string;
  meetingId: string;
  transcriptId: string | null;
  lastContentLength: number;
  intervalHandle: NodeJS.Timeout;
  active: boolean;
}

const activePollers = new Map<string, ActivePoller>();

/**
 * Extract organizer user ID from a Teams meeting join URL
 * URL contains context param with "oid" (organizer object ID)
 */
function extractOrganizerFromUrl(joinUrl: string): string | null {
  try {
    const url = new URL(joinUrl);
    const contextParam = url.searchParams.get("context");
    if (contextParam) {
      const context = JSON.parse(decodeURIComponent(contextParam));
      return context.oid || null;
    }
  } catch {
    // Try regex fallback
    const match = joinUrl.match(/"oid"\s*:\s*"([^"]+)"/);
    if (match) return match[1];
  }
  return null;
}

/**
 * Find an online meeting by its join URL for a given user
 */
async function findMeetingByJoinUrl(
  userId: string,
  joinUrl: string
): Promise<OnlineMeeting | null> {
  const client = getGraphClient();

  // Encode the joinUrl for use in a filter
  const encodedUrl = encodeURIComponent(joinUrl);
  const response = await client.request<{ value: OnlineMeeting[] }>(
    "GET",
    `/v1.0/users/${userId}/onlineMeetings?$filter=joinWebUrl eq '${joinUrl}'`
  );

  if (!response.success || !response.data?.value?.length) {
    // Try beta endpoint as fallback
    const betaResponse = await client.request<{ value: OnlineMeeting[] }>(
      "GET",
      `/beta/users/${userId}/onlineMeetings?$filter=joinWebUrl eq '${joinUrl}'`
    );
    if (betaResponse.success && betaResponse.data?.value?.length) {
      return betaResponse.data.value[0];
    }
    return null;
  }

  return response.data.value[0];
}

/**
 * Get available transcripts for a meeting
 */
async function getMeetingTranscripts(
  userId: string,
  meetingId: string
): Promise<Transcript[]> {
  const client = getGraphClient();
  const response = await client.request<{ value: Transcript[] }>(
    "GET",
    `/v1.0/users/${userId}/onlineMeetings/${meetingId}/transcripts`
  );

  if (!response.success || !response.data?.value) return [];
  return response.data.value;
}

/**
 * Get transcript content in VTT format (returns raw text)
 */
async function getTranscriptContent(
  userId: string,
  meetingId: string,
  transcriptId: string
): Promise<string | null> {
  const client = getGraphClient();
  const response = await client.request<string>(
    "GET",
    `/v1.0/users/${userId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content?$format=text/vtt`
  );

  if (!response.success || !response.data) return null;
  return response.data as string;
}

/**
 * Parse VTT content into transcript segments
 */
function parseVttSegments(vtt: string): Array<{ speaker: string; text: string; startMs: number }> {
  const segments: Array<{ speaker: string; text: string; startMs: number }> = [];
  const lines = vtt.split("\n");

  let i = 0;
  while (i < lines.length) {
    // Look for timestamp lines: 00:00:01.000 --> 00:00:05.000
    const timestampMatch = lines[i]?.match(/^(\d{2}:\d{2}:\d{2}\.\d{3})\s*-->/);
    if (timestampMatch) {
      const startTime = timestampMatch[1];
      const [h, m, s] = startTime.split(":").map(Number);
      const startMs = (h * 3600 + m * 60 + s) * 1000;

      // Check for speaker in the next lines
      i++;
      let speaker = "Unknown";
      let textLines: string[] = [];

      while (i < lines.length && lines[i].trim() !== "") {
        const line = lines[i];
        // VTT speaker format: <v Speaker Name>text</v>
        const speakerMatch = line.match(/<v ([^>]+)>(.+?)(?:<\/v>)?$/);
        if (speakerMatch) {
          speaker = speakerMatch[1];
          textLines.push(speakerMatch[2].replace(/<[^>]+>/g, "").trim());
        } else if (line.trim()) {
          textLines.push(line.replace(/<[^>]+>/g, "").trim());
        }
        i++;
      }

      const text = textLines.join(" ").trim();
      if (text) {
        segments.push({ speaker, text, startMs });
      }
    } else {
      i++;
    }
  }

  return segments;
}

/**
 * Send a transcript chunk to Project Missa
 */
async function sendChunkToProjectMissa(
  callId: string,
  speaker: string,
  text: string,
  startMs: number
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
        offsetMs: startMs,
        durationMs: 0,
        isFinal: true,
        source: "teams_transcript",
      }),
    });

    if (!response.ok) {
      console.error(`[TranscriptPoller] Failed to send chunk: ${response.status}`);
    }
  } catch (error) {
    console.error(`[TranscriptPoller] Error sending chunk:`, error);
  }
}

/**
 * Main polling function - checks for new transcript content
 */
async function pollTranscript(poller: ActivePoller): Promise<void> {
  if (!poller.active) return;

  const client = getGraphClient();

  try {
    // If we don't have a transcript ID yet, look for one
    if (!poller.transcriptId) {
      const transcripts = await getMeetingTranscripts(poller.organizerUserId, poller.meetingId);
      if (transcripts.length > 0) {
        // Use the most recent transcript (last one)
        poller.transcriptId = transcripts[transcripts.length - 1].id;
        console.log(`[TranscriptPoller] Found transcript ${poller.transcriptId} for call ${poller.callId}`);
      } else {
        console.log(`[TranscriptPoller] No transcript yet for call ${poller.callId} - waiting for Teams transcription to be enabled...`);
        return;
      }
    }

    // Get current transcript content
    const content = await getTranscriptContent(
      poller.organizerUserId,
      poller.meetingId,
      poller.transcriptId
    );

    if (!content) return;

    // Only process if there's new content
    const contentLength = content.length;
    if (contentLength <= poller.lastContentLength) return;

    console.log(`[TranscriptPoller] New transcript content detected (${poller.lastContentLength} -> ${contentLength} chars)`);

    // Parse VTT and get new segments
    const allSegments = parseVttSegments(content);

    // We need to figure out which segments are new
    // Parse the old content too to find the cut-off point
    const oldContent = content.substring(0, poller.lastContentLength);
    const oldSegments = parseVttSegments(oldContent);
    const newSegments = allSegments.slice(oldSegments.length);

    poller.lastContentLength = contentLength;

    // Send new segments to Project Missa
    for (const segment of newSegments) {
      console.log(`[TranscriptPoller] New segment - ${segment.speaker}: ${segment.text}`);
      await sendChunkToProjectMissa(poller.callId, segment.speaker, segment.text, segment.startMs);
    }

  } catch (error) {
    console.error(`[TranscriptPoller] Polling error for call ${poller.callId}:`, error);
  }
}

/**
 * Start transcript polling for a meeting
 * Returns true if polling started successfully, false otherwise
 */
export async function startTranscriptPolling(
  callId: string,
  joinUrl: string
): Promise<{ success: boolean; message: string }> {
  console.log(`[TranscriptPoller] Starting transcript polling for call ${callId}`);

  // Extract organizer user ID from the join URL
  const organizerUserId = extractOrganizerFromUrl(joinUrl);
  if (!organizerUserId) {
    console.error(`[TranscriptPoller] Could not extract organizer ID from URL: ${joinUrl.substring(0, 60)}...`);
    return { success: false, message: "Could not determine meeting organizer" };
  }

  console.log(`[TranscriptPoller] Organizer user ID: ${organizerUserId}`);

  // Find the meeting
  console.log(`[TranscriptPoller] Looking up meeting...`);
  const meeting = await findMeetingByJoinUrl(organizerUserId, joinUrl);

  if (!meeting) {
    console.error(`[TranscriptPoller] Could not find meeting for URL`);
    return { success: false, message: "Meeting not found in Graph API" };
  }

  console.log(`[TranscriptPoller] Found meeting: ${meeting.id} - ${meeting.subject || "No title"}`);

  // Set up the poller
  const poller: ActivePoller = {
    callId,
    joinUrl,
    organizerUserId,
    meetingId: meeting.id,
    transcriptId: null,
    lastContentLength: 0,
    intervalHandle: null as unknown as NodeJS.Timeout,
    active: true,
  };

  // Start polling every 10 seconds
  poller.intervalHandle = setInterval(() => pollTranscript(poller), 10000);
  activePollers.set(callId, poller);

  // Do an immediate first poll
  pollTranscript(poller);

  return {
    success: true,
    message: `Monitoring meeting transcript. Ask the meeting organizer to enable transcription (Meeting controls → ... → Start transcription) if not already enabled.`,
  };
}

/**
 * Stop transcript polling for a call
 */
export function stopTranscriptPolling(callId: string): void {
  const poller = activePollers.get(callId);
  if (poller) {
    poller.active = false;
    clearInterval(poller.intervalHandle);
    activePollers.delete(callId);
    console.log(`[TranscriptPoller] Stopped polling for call ${callId}`);
  }
}

/**
 * Check if transcript polling is active for a call
 */
export function isPollingActive(callId: string): boolean {
  return activePollers.has(callId) && (activePollers.get(callId)?.active ?? false);
}
