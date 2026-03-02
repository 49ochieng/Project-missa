/**
 * Call Manager - Handles joining and managing Teams meetings via Graph Cloud Communications API
 */

import { getGraphClient, GraphClient } from "./graphClient";
import { getConfig } from "../config";

/**
 * Call resource from Graph API
 */
export interface CallResource {
  "@odata.type": string;
  id: string;
  state: "incoming" | "establishing" | "established" | "hold" | "transferring" | "transferAccepted" | "redirecting" | "terminating" | "terminated";
  resultInfo?: {
    code: number;
    subcode: number;
    message: string;
  };
  direction: "incoming" | "outgoing";
  callbackUri: string;
  source?: ParticipantInfo;
  targets?: ParticipantInfo[];
  chatInfo?: ChatInfo;
  meetingInfo?: MeetingInfo;
  tenantId: string;
  myParticipantId?: string;
}

export interface ParticipantInfo {
  "@odata.type"?: string;
  identity?: {
    application?: { id: string; displayName: string };
    user?: { id: string; displayName: string };
  };
}

export interface ChatInfo {
  "@odata.type"?: string;
  threadId: string;
  messageId?: string;
  replyChainMessageId?: string;
}

export interface MeetingInfo {
  "@odata.type": string;
  joinUrl?: string;
  organizerId?: string;
  allowConversationWithoutHost?: boolean;
}

export interface JoinMeetingResult {
  success: boolean;
  callId?: string;
  error?: string;
}

/**
 * Media configuration for the call
 */
interface MediaConfig {
  "@odata.type": string;
  preUploadedMedia?: unknown[];
}

/**
 * Join meeting request payload
 */
interface JoinMeetingPayload {
  "@odata.type": string;
  callbackUri: string;
  requestedModalities: string[];
  mediaConfig: MediaConfig;
  chatInfo?: ChatInfo;
  meetingInfo: MeetingInfo;
  tenantId?: string;
}

/**
 * Active calls tracking
 */
const activeCalls = new Map<string, CallResource>();

/**
 * Join a Teams meeting using the join URL
 * 
 * @param joinUrl - The Teams meeting join URL
 * @param displayName - Display name for the bot in the meeting (optional)
 * @returns Result with call ID or error
 */
export async function joinMeeting(
  joinUrl: string,
  displayName?: string
): Promise<JoinMeetingResult> {
  const config = getConfig();
  const client = getGraphClient();

  console.log(`[CallManager] Joining meeting: ${joinUrl.substring(0, 50)}...`);

  // Construct callback URI for Graph to notify us of call events
  const callbackUri = `${config.botEndpoint}/api/calls/callback`;

  // Build join request
  const payload: JoinMeetingPayload = {
    "@odata.type": "#microsoft.graph.call",
    callbackUri,
    requestedModalities: ["audio"],
    mediaConfig: {
      "@odata.type": "#microsoft.graph.serviceHostedMediaConfig",
    },
    meetingInfo: {
      "@odata.type": "#microsoft.graph.organizerMeetingInfo",
      joinUrl,
      allowConversationWithoutHost: true,
    },
  };

  // If tenant ID is available, include it
  if (config.azureTenantId) {
    payload.tenantId = config.azureTenantId;
  }

  const response = await client.request<CallResource>(
    "POST",
    "/v1.0/communications/calls",
    payload
  );

  if (!response.success || !response.data) {
    console.error(`[CallManager] Failed to join meeting: ${response.error}`);
    return {
      success: false,
      error: response.error || "Failed to join meeting",
    };
  }

  const call = response.data;
  console.log(`[CallManager] Call created: ${call.id}, state: ${call.state}`);

  // Track active call
  activeCalls.set(call.id, call);

  return {
    success: true,
    callId: call.id,
  };
}

/**
 * Leave/terminate a call
 * 
 * @param callId - The call ID to terminate
 */
export async function leaveCall(callId: string): Promise<{ success: boolean; error?: string }> {
  const client = getGraphClient();

  console.log(`[CallManager] Leaving call: ${callId}`);

  const response = await client.request(
    "DELETE",
    `/v1.0/communications/calls/${callId}`
  );

  if (!response.success) {
    console.error(`[CallManager] Failed to leave call: ${response.error}`);
    return {
      success: false,
      error: response.error,
    };
  }

  // Remove from tracking
  activeCalls.delete(callId);
  console.log(`[CallManager] Successfully left call: ${callId}`);

  return { success: true };
}

/**
 * Update call state from callback notification
 */
export function updateCallState(callId: string, call: Partial<CallResource>): void {
  const existing = activeCalls.get(callId);
  if (existing) {
    activeCalls.set(callId, { ...existing, ...call });
    console.log(`[CallManager] Updated call ${callId} state: ${call.state}`);
  } else if (call.id) {
    activeCalls.set(callId, call as CallResource);
    console.log(`[CallManager] Tracking new call ${callId} state: ${call.state}`);
  }
}

/**
 * Get call by ID
 */
export function getCall(callId: string): CallResource | undefined {
  return activeCalls.get(callId);
}

/**
 * Get all active calls
 */
export function getActiveCalls(): Map<string, CallResource> {
  return activeCalls;
}

/**
 * Check if call is still active
 */
export function isCallActive(callId: string): boolean {
  const call = activeCalls.get(callId);
  if (!call) return false;
  
  return !["terminated", "terminating"].includes(call.state);
}

/**
 * Get participants in a call
 */
export async function getCallParticipants(
  callId: string
): Promise<{ success: boolean; participants?: ParticipantInfo[]; error?: string }> {
  const client = getGraphClient();

  const response = await client.request<{ value: ParticipantInfo[] }>(
    "GET",
    `/v1.0/communications/calls/${callId}/participants`
  );

  if (!response.success || !response.data) {
    return {
      success: false,
      error: response.error,
    };
  }

  return {
    success: true,
    participants: response.data.value,
  };
}

/**
 * Subscribe to tone notifications (DTMF)
 */
export async function subscribeToTone(
  callId: string
): Promise<{ success: boolean; error?: string }> {
  const config = getConfig();
  const client = getGraphClient();

  const response = await client.request(
    "POST",
    `/v1.0/communications/calls/${callId}/subscribeToTone`,
    {
      clientContext: "meeting-media-bot",
    }
  );

  return {
    success: response.success,
    error: response.error,
  };
}
