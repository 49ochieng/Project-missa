/**
 * Meeting Media Bot Client
 * Handles communication with the meeting-media-bot service for real-time capture
 */

import { getAppConfig } from "../utils/config";
import type { ILogger } from "@microsoft/teams.common";

export interface JoinMeetingResult {
  success: boolean;
  callId?: string;
  meetingId?: string;
  error?: string;
}

export interface LeaveMeetingResult {
  success: boolean;
  error?: string;
}

export interface MeetingStatus {
  callId: string;
  state: string;
  isActive: boolean;
  isTranscribing: boolean;
}

/**
 * Client for communicating with meeting-media-bot service
 */
export class MeetingMediaBotClient {
  private logger: ILogger;
  private baseUrl: string;
  private sharedSecret: string;

  constructor(logger: ILogger) {
    this.logger = logger;
    const config = getAppConfig();
    this.baseUrl = config.meetingMediaBotUrl;
    this.sharedSecret = config.meetingMediaBotSharedSecret;
  }

  /**
   * Request the meeting-media-bot to join a Teams meeting
   * 
   * @param joinUrl - Teams meeting join URL
   * @param meetingId - Optional meeting ID for tracking
   * @returns Result with call ID if successful
   */
  async startMeetingCapture(
    joinUrl: string,
    meetingId?: string
  ): Promise<JoinMeetingResult> {
    this.logger.debug(`Requesting meeting capture for: ${joinUrl.substring(0, 50)}...`);

    try {
      const response = await fetch(`${this.baseUrl}/api/meetings/join`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shared-Secret": this.sharedSecret,
        },
        body: JSON.stringify({
          joinUrl,
          meetingId,
        }),
      });

      const data = await response.json() as JoinMeetingResult;

      if (!response.ok) {
        this.logger.error(`Failed to start meeting capture: ${data.error}`);
        return {
          success: false,
          error: data.error || `HTTP ${response.status}`,
        };
      }

      this.logger.info(`Meeting capture started, callId: ${data.callId}`);
      return data;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`Error starting meeting capture: ${errorMessage}`);
      return {
        success: false,
        error: errorMessage,
      };
    }
  }

  /**
   * Request the meeting-media-bot to leave a meeting
   * 
   * @param callId - The call ID returned from startMeetingCapture
   * @returns Result indicating success or failure
   */
  async stopMeetingCapture(callId: string): Promise<LeaveMeetingResult> {
    this.logger.debug(`Requesting to stop meeting capture for callId: ${callId}`);

    try {
      const response = await fetch(`${this.baseUrl}/api/meetings/leave`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shared-Secret": this.sharedSecret,
        },
        body: JSON.stringify({ callId }),
      });

      const data = await response.json() as LeaveMeetingResult;

      if (!response.ok) {
        this.logger.error(`Failed to stop meeting capture: ${data.error}`);
        return {
          success: false,
          error: data.error || `HTTP ${response.status}`,
        };
      }

      this.logger.info(`Meeting capture stopped for callId: ${callId}`);
      return data;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`Error stopping meeting capture: ${errorMessage}`);
      return {
        success: false,
        error: errorMessage,
      };
    }
  }

  /**
   * Get the status of an active meeting capture
   * 
   * @param callId - The call ID to check
   * @returns Meeting status or error
   */
  async getMeetingCaptureStatus(callId: string): Promise<{ success: boolean; status?: MeetingStatus; error?: string }> {
    this.logger.debug(`Getting status for callId: ${callId}`);

    try {
      const response = await fetch(`${this.baseUrl}/api/meetings/${callId}/status`, {
        method: "GET",
        headers: {
          "X-Shared-Secret": this.sharedSecret,
        },
      });

      if (!response.ok) {
        const data = await response.json() as { error?: string };
        return {
          success: false,
          error: data.error || `HTTP ${response.status}`,
        };
      }

      const status = await response.json() as MeetingStatus;
      return {
        success: true,
        status,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`Error getting meeting status: ${errorMessage}`);
      return {
        success: false,
        error: errorMessage,
      };
    }
  }

  /**
   * Check if the meeting-media-bot service is available
   * Retries once after a delay to handle Azure cold starts
   */
  async checkHealth(): Promise<boolean> {
    const url = `${this.baseUrl}/api/health`;
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        this.logger.info(`[HealthCheck] Attempt ${attempt}: ${url}`);
        const response = await fetch(url, {
          method: "GET",
          signal: AbortSignal.timeout(15000),
        });
        if (response.ok) {
          this.logger.info(`[HealthCheck] Meeting media bot is reachable`);
          return true;
        }
        this.logger.warn(`[HealthCheck] HTTP ${response.status} from ${url}`);
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        this.logger.warn(`[HealthCheck] Attempt ${attempt} failed: ${msg} (url: ${url})`);
      }
      if (attempt < 2) {
        this.logger.info(`[HealthCheck] Retrying in 3s (cold start recovery)...`);
        await new Promise(r => setTimeout(r, 3000));
      }
    }
    this.logger.error(`[HealthCheck] Meeting media bot unreachable after 2 attempts: ${url}`);
    return false;
  }
}

// Singleton instance
let clientInstance: MeetingMediaBotClient | null = null;

/**
 * Get or create the MeetingMediaBotClient instance
 */
export function getMeetingMediaBotClient(logger: ILogger): MeetingMediaBotClient {
  if (!clientInstance) {
    clientInstance = new MeetingMediaBotClient(logger);
  }
  return clientInstance;
}
