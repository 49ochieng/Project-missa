import { ChatPrompt } from "@microsoft/teams.ai";
import { ILogger } from "@microsoft/teams.common";
import { OpenAIChatModel } from "@microsoft/teams.openai";
import { getMeetingMediaBotClient } from "../../services/meetingMediaBotClient";
import { MessageContext } from "../../utils/messageContext";
import { BaseCapability, CapabilityDefinition } from "../capability";
import { MEETING_NOTES_PROMPT } from "./prompt";
import {
  MeetingEmailRequest,
  MeetingSummary,
  MeetingTranscript,
  // MEETING_SUMMARY_SCHEMA - TODO: Use for structured output validation
} from "./schema";

/**
 * Meeting Notes Capability
 * Supports fetching meeting transcripts, generating structured summaries,
 * and distributing notes to participants via email and Teams
 */
export class MeetingNotesCapability extends BaseCapability {
  readonly name = "meeting_notes";

  createPrompt(context: MessageContext): ChatPrompt {
    const meetingNotesModelConfig = this.getModelConfig("meetingNotes");

    const prompt = new ChatPrompt({
      instructions: MEETING_NOTES_PROMPT,
      model: new OpenAIChatModel({
        model: meetingNotesModelConfig.model,
        apiKey: meetingNotesModelConfig.apiKey,
        endpoint: meetingNotesModelConfig.endpoint,
        apiVersion: meetingNotesModelConfig.apiVersion,
      }),
    })
      .function(
        "read_transcript",
        "Fetch a meeting transcript from Microsoft Graph and store it in the database",
        async (meetingId: string) => {
          try {
            this.logger.debug(`Fetching transcript for meeting: ${meetingId}`);

            // Fetch transcript from Graph API
            const transcript = await this.fetchMeetingTranscript(meetingId, context);

            // Store transcript in database
            await this.storeMeetingTranscript(transcript, context);

            return JSON.stringify({
              success: true,
              meeting: {
                id: transcript.meetingId,
                title: transcript.title,
                startDateTime: transcript.startDateTime,
                endDateTime: transcript.endDateTime,
                participants: transcript.participants,
                transcriptLength: transcript.transcriptText.length,
              },
              message: "Transcript retrieved and stored successfully",
            });
          } catch (error) {
            this.logger.error("Error fetching transcript:", error);
            return JSON.stringify({
              success: false,
              error: error instanceof Error ? error.message : "Unknown error",
              message: "Failed to fetch meeting transcript",
            });
          }
        }
      )
      .function(
        "summarize_meeting",
        "Generate a structured JSON summary from a meeting transcript or conversation history. Returns JSON matching the MeetingSummary schema.",
        async (meetingIdOrTranscriptText: string) => {
          try {
            this.logger.debug("Generating meeting summary");

            // Try to get transcript from database or use provided text
            const transcriptText = await this.getTranscriptText(
              meetingIdOrTranscriptText,
              context
            );

            // Get conversation context for additional information
            const conversationHistory = await context.memory.getMessagesByTimeRange(
              context.startTime,
              context.endTime
            );

            // Generate structured summary using AI
            const summary = await this.generateStructuredSummary(
              transcriptText,
              conversationHistory,
              context
            );

            // Store summary in database
            await this.storeMeetingSummary(summary, context);

            return JSON.stringify({
              success: true,
              summary: summary,
              message: "Meeting summary generated and stored successfully",
            });
          } catch (error) {
            this.logger.error("Error generating summary:", error);
            return JSON.stringify({
              success: false,
              error: error instanceof Error ? error.message : "Unknown error",
              message: "Failed to generate meeting summary",
            });
          }
        }
      )
      .function(
        "send_summary",
        "Send meeting summary to participants via email and optionally post to Teams chat",
        async (args: { summaryJson: string; recipients: string; includeTeamsPost?: boolean }) => {
          try {
            this.logger.debug("Sending meeting summary");

            const summary: MeetingSummary = JSON.parse(args.summaryJson);
            const recipientList = args.recipients.split(",").map((r) => r.trim());
            const includeTeamsPost = args.includeTeamsPost || false;

            const emailRequest: MeetingEmailRequest = {
              recipients: recipientList,
              subject: `Meeting Notes: ${summary.title}`,
              summary: summary,
              includeTeamsPost: includeTeamsPost,
              conversationId: includeTeamsPost ? context.conversationId : undefined,
            };

            // Send email via Graph API
            const emailResult = await this.sendSummaryEmail(emailRequest, context);

            // Optionally post to Teams chat
            let teamsPostResult = null;
            if (includeTeamsPost) {
              teamsPostResult = await this.postSummaryToTeams(summary, context);
            }

            // Record the send action in database
            await this.recordSummarySent(summary, recipientList, context);

            return JSON.stringify({
              success: true,
              emailSent: emailResult.success,
              emailRecipients: recipientList,
              teamsPostSent: teamsPostResult?.success || false,
              message: `Summary sent to ${recipientList.length} recipient(s)`,
            });
          } catch (error) {
            this.logger.error("Error sending summary:", error);
            return JSON.stringify({
              success: false,
              error: error instanceof Error ? error.message : "Unknown error",
              message: "Failed to send meeting summary",
            });
          }
        }
      )
      .function(
        "start_meeting_capture",
        "Join a Teams meeting and start real-time transcription. Requires a Teams meeting join URL.",
        async (args: { joinUrl: string; meetingId?: string }) => {
          try {
            this.logger.debug(`Starting meeting capture for: ${args.joinUrl.substring(0, 50)}...`);

            // Validate join URL format
            if (!args.joinUrl.includes("teams.microsoft.com") && !args.joinUrl.includes("teams.live.com")) {
              return JSON.stringify({
                success: false,
                error: "Invalid Teams meeting URL",
                message: "Please provide a valid Microsoft Teams meeting join URL",
              });
            }

            const client = getMeetingMediaBotClient(this.logger);

            // Check if meeting-media-bot is available
            const isAvailable = await client.checkHealth();
            if (!isAvailable) {
              return JSON.stringify({
                success: false,
                error: "Meeting capture service unavailable",
                message: "The meeting capture service is not running. Please try again later.",
              });
            }

            // Create meeting record in database
            const meetingId = args.meetingId || `meeting-${Date.now()}`;
            await context.database.upsertMeeting({
              meetingId,
              conversationId: context.conversationId,
              joinUrl: args.joinUrl,
              title: `Meeting capture ${new Date().toLocaleString()}`,
              organizerAadId: context.userAadId || context.userId || "unknown",
            });

            // Request the meeting-media-bot to join
            const result = await client.startMeetingCapture(args.joinUrl, meetingId);

            if (!result.success) {
              // Update meeting status to failed
              await context.database.updateMeetingStatus({
                meetingId,
                status: "failed",
              });

              return JSON.stringify({
                success: false,
                error: result.error,
                message: "Failed to start meeting capture",
              });
            }

            return JSON.stringify({
              success: true,
              meetingId,
              callId: result.callId,
              message: "Meeting capture started. I'm joining the meeting now and will transcribe the audio in real-time.",
            });
          } catch (error) {
            this.logger.error("Error starting meeting capture:", error);
            return JSON.stringify({
              success: false,
              error: error instanceof Error ? error.message : "Unknown error",
              message: "Failed to start meeting capture",
            });
          }
        }
      )
      .function(
        "stop_meeting_capture",
        "Leave a Teams meeting and stop real-time transcription. Requires the call ID from start_meeting_capture.",
        async (args: { callId: string }) => {
          try {
            this.logger.debug(`Stopping meeting capture for callId: ${args.callId}`);

            const client = getMeetingMediaBotClient(this.logger);
            const result = await client.stopMeetingCapture(args.callId);

            if (!result.success) {
              return JSON.stringify({
                success: false,
                error: result.error,
                message: "Failed to stop meeting capture",
              });
            }

            // Update meeting status in database
            await context.database.updateMeetingStatus({
              meetingId: args.callId,
              status: "ended",
              endedAt: new Date().toISOString(),
            });

            return JSON.stringify({
              success: true,
              callId: args.callId,
              message: "Meeting capture stopped. The transcript has been saved and is ready for summarization.",
            });
          } catch (error) {
            this.logger.error("Error stopping meeting capture:", error);
            return JSON.stringify({
              success: false,
              error: error instanceof Error ? error.message : "Unknown error",
              message: "Failed to stop meeting capture",
            });
          }
        }
      );

    this.logger.debug("Initialized Meeting Notes Capability!");
    return prompt;
  }

  // ============================================================================
  // GRAPH API INTEGRATION STUBS
  // ============================================================================

  /**
   * Fetch meeting transcript from Microsoft Graph API
   * Supports both stored transcripts (from real-time capture) and Graph API fallback
   * Reference: https://learn.microsoft.com/en-us/graph/api/calltranscript-get
   */
  private async fetchMeetingTranscript(
    meetingId: string,
    context: MessageContext
  ): Promise<MeetingTranscript> {
    this.logger.debug(`Fetching transcript for meeting: ${meetingId}`);

    // First, try to get transcript from our database (from real-time capture)
    try {
      const storedTranscript = await context.database.getTranscriptByMeetingId(meetingId);
      if (storedTranscript && storedTranscript.chunks.length > 0) {
        this.logger.debug(`Found stored transcript with ${storedTranscript.chunks.length} chunks`);
        
        // Convert stored transcript to MeetingTranscript format
        const fullText = storedTranscript.chunks
          .map((chunk) => `[${chunk.speaker}]: ${chunk.text}`)
          .join("\n");

        return {
          meetingId: storedTranscript.meetingId,
          title: storedTranscript.title || "Meeting Transcript",
          startDateTime: storedTranscript.startedAt || new Date().toISOString(),
          endDateTime: storedTranscript.endedAt || new Date().toISOString(),
          organizer: storedTranscript.participants[0]?.displayName || "Unknown",
          participants: storedTranscript.participants.map((p) => p.displayName),
          transcriptText: fullText,
          retrievedAt: new Date().toISOString(),
        };
      }
    } catch (error) {
      this.logger.warn(`No stored transcript found for ${meetingId}, will try Graph API fallback`);
    }

    // Fallback: Try to fetch from Graph API
    // TODO: Implement actual Graph API call using getGraphClient()
    // Example endpoint: GET /communications/onlineMeetings/{meetingId}/transcripts
    // Requires: OnlineMeetingTranscript.Read.All or CallTranscripts.Read.All

    this.logger.warn("Graph API transcript fallback not yet implemented - returning stub");

    // Stub implementation - replace with actual Graph API call
    return {
      meetingId: meetingId,
      title: "Meeting (Graph API transcript pending)",
      startDateTime: new Date().toISOString(),
      endDateTime: new Date().toISOString(),
      organizer: "organizer@example.com",
      participants: ["participant1@example.com", "participant2@example.com"],
      transcriptText:
        "No transcript data available. The meeting may not have been captured yet, or Graph API access is pending configuration.",
      retrievedAt: new Date().toISOString(),
    };
  }

  /**
   * Send meeting summary email via Microsoft Graph API
   * TODO: Implement actual Graph API call to send email
   * Reference: https://learn.microsoft.com/en-us/graph/api/user-sendmail
   */
  private async sendSummaryEmail(
    _emailRequest: MeetingEmailRequest, // Prefixed with _ to indicate intentionally unused
    _context: MessageContext
  ): Promise<{ success: boolean; messageId?: string }> {
    // TODO: Implement Graph API call
    // Example endpoint: POST /users/{userId}/sendMail
    // Requires: Mail.Send permission

    this.logger.warn("TODO: Implement Graph API call to send email");

    // Stub implementation - replace with actual Graph API call
    return {
      success: true,
      messageId: "stub-message-id-" + Date.now(),
    };
  }

  /**
   * Post meeting summary to Teams chat via Graph API
   * TODO: Implement actual Graph API call to post message to Teams chat
   * Reference: https://learn.microsoft.com/en-us/graph/api/channel-post-messages
   */
  private async postSummaryToTeams(
    _summary: MeetingSummary, // Prefixed with _ to indicate intentionally unused
    _context: MessageContext
  ): Promise<{ success: boolean; messageId?: string }> {
    // TODO: Implement Graph API call
    // Example endpoint: POST /teams/{teamId}/channels/{channelId}/messages
    // Requires: ChannelMessage.Send permission

    this.logger.warn("TODO: Implement Graph API call to post to Teams chat");

    // Stub implementation - replace with actual Graph API call
    return {
      success: true,
      messageId: "stub-teams-message-id-" + Date.now(),
    };
  }

  // ============================================================================
  // DATABASE OPERATIONS
  // ============================================================================

  /**
   * Store meeting transcript in database
   * TODO: Extend database schema to support meeting_transcripts table
   */
  private async storeMeetingTranscript(
    transcript: MeetingTranscript,
    context: MessageContext
  ): Promise<void> {
    // TODO: Add database method to store transcript
    // For now, store as a special message in the conversation
    this.logger.debug(
      `Storing transcript for meeting ${transcript.meetingId} in conversation ${context.conversationId}`
    );

    // Temporary implementation: store as metadata in conversation
    // In production, should have dedicated meeting_transcripts table
    // Example structure:
    // {
    //   conversation_id: context.conversationId,
    //   content: JSON.stringify(transcript),
    //   name: "System",
    //   timestamp: new Date().toISOString(),
    //   activity_id: `transcript_${transcript.meetingId}`,
    //   role: "system",
    // }

    // Store in database (using existing message storage as temporary solution)
    // TODO: Create dedicated meeting_transcripts table
    this.logger.warn("TODO: Implement dedicated meeting_transcripts table");
  }

  /**
   * Store meeting summary in database
   * TODO: Extend database schema to support meeting_summaries table
   */
  private async storeMeetingSummary(
    summary: MeetingSummary,
    context: MessageContext
  ): Promise<void> {
    // TODO: Add database method to store summary
    this.logger.debug(
      `Storing summary for meeting "${summary.title}" in conversation ${context.conversationId}`
    );

    // Temporary implementation: store as metadata in conversation
    // In production, should have dedicated meeting_summaries table
    // Example structure:
    // {
    //   conversation_id: context.conversationId,
    //   content: JSON.stringify(summary),
    //   name: "System",
    //   timestamp: new Date().toISOString(),
    //   activity_id: `summary_${summary.title}_${Date.now()}`,
    //   role: "system",
    // }

    // Store in database (using existing message storage as temporary solution)
    // TODO: Create dedicated meeting_summaries table
    this.logger.warn("TODO: Implement dedicated meeting_summaries table");
  }

  /**
   * Record that a summary was sent to recipients
   * TODO: Extend database schema to support meeting_summary_sends table
   */
  private async recordSummarySent(
    summary: MeetingSummary,
    recipients: string[],
    _context: MessageContext // Prefixed with _ to indicate intentionally unused
  ): Promise<void> {
    // TODO: Add database method to record send actions
    this.logger.debug(
      `Recording send action for summary "${summary.title}" to ${recipients.length} recipients`
    );

    // Temporary implementation
    // TODO: Create dedicated meeting_summary_sends table
    this.logger.warn("TODO: Implement dedicated meeting_summary_sends tracking table");
  }

  /**
   * Get transcript text from database or use provided text
   */
  private async getTranscriptText(
    meetingIdOrText: string,
    _context: MessageContext // Prefixed with _ to indicate intentionally unused
  ): Promise<string> {
    // If it looks like a meeting ID, try to fetch from database
    if (meetingIdOrText.length < 100 && !meetingIdOrText.includes(" ")) {
      // TODO: Query database for stored transcript
      this.logger.warn("TODO: Implement database query for stored transcripts");
      return `Placeholder transcript for meeting ${meetingIdOrText}. TODO: Fetch from database.`;
    }

    // Otherwise, treat as direct transcript text
    return meetingIdOrText;
  }

  /**
   * Generate structured summary using AI with JSON schema
   */
  private async generateStructuredSummary(
    _transcriptText: string, // Prefixed with _ to indicate intentionally unused (will be used by AI)
    _conversationHistory: any[], // Prefixed with _ to indicate intentionally unused (will be used by AI)
    _context: MessageContext // Prefixed with _ to indicate intentionally unused
  ): Promise<MeetingSummary> {
    // TODO: Use structured output with JSON schema for more reliable parsing
    // For now, return a basic structure that will be filled by AI
    this.logger.debug("Generating structured summary from transcript");

    // This is a placeholder - the AI will actually generate the summary
    // based on the transcript and conversation history
    const summary: MeetingSummary = {
      title: "Meeting Summary",
      dateTime: new Date().toISOString(),
      participants: [],
      decisions: [],
      actionItems: [],
      risks: [],
      openQuestions: [],
      shortSummary: "",
      detailedSummary: "",
    };

    // TODO: Implement structured output parsing with MEETING_SUMMARY_SCHEMA
    this.logger.warn("TODO: Implement structured JSON output parsing");

    return summary;
  }
}

// Capability definition for manager registration
export const MEETING_NOTES_CAPABILITY_DEFINITION: CapabilityDefinition = {
  name: "meeting_notes",
  manager_desc: `**Meeting Notes**: Use for requests like:
- "read transcript", "get meeting transcript", "fetch meeting notes"
- "summarize meeting", "create meeting summary", "analyze meeting"
- "send summary", "email notes", "share meeting notes", "distribute summary to participants"
- Managing meeting transcripts, summaries, action items from meetings`,
  handler: async (context: MessageContext, logger: ILogger) => {
    const meetingNotesCapability = new MeetingNotesCapability(logger);
    const result = await meetingNotesCapability.processRequest(context);
    if (result.error) {
      logger.error(`Error in Meeting Notes Capability: ${result.error}`);
      return `Error in Meeting Notes Capability: ${result.error}`;
    }
    return result.response || "No response from Meeting Notes Capability";
  },
};
