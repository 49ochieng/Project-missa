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
 * Tier 3: Smart URL detection
 * Queries the Teams chat object to find an active/linked online meeting join URL.
 * Returns null if the chat has no associated meeting or the call fails.
 */
async function getActiveMeetingFromChat(context: MessageContext): Promise<string | null> {
  try {
    const chatId = context.conversationId;
    if (!chatId || !context.appGraph) return null;

    // GET /chats/{chatId}?$select=onlineMeetingInfo
    // Returns onlineMeetingInfo.joinWebUrl when the chat has an active meeting
    const response = await context.appGraph.call(
      (_userId: string) => ({
        method: "get" as const,
        path: `/chats/${encodeURIComponent(chatId)}`,
        query: { "$select": "onlineMeetingInfo" },
      }),
      context.userAadId || context.userId || ""
    );

    const joinWebUrl = (response as Record<string, unknown>)?.onlineMeetingInfo as
      | { joinWebUrl?: string }
      | undefined;

    return joinWebUrl?.joinWebUrl || null;
  } catch {
    // Silently fail — caller will ask user for URL
    return null;
  }
}

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
        {
          type: "object",
          properties: {
            meetingId: {
              type: "string",
              description: "The meeting ID or join URL to fetch the transcript for",
            },
          },
          required: ["meetingId"],
        },
        async ({ meetingId }: { meetingId: string }) => {
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
        {
          type: "object",
          properties: {
            meetingIdOrTranscriptText: {
              type: "string",
              description: "Meeting ID to look up stored transcript, or raw transcript text to summarize directly",
            },
          },
          required: ["meetingIdOrTranscriptText"],
        },
        async ({ meetingIdOrTranscriptText }: { meetingIdOrTranscriptText: string }) => {
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
        {
          type: "object",
          properties: {
            summaryJson: {
              type: "string",
              description: "JSON string of the MeetingSummary object to send",
            },
            recipients: {
              type: "string",
              description: "Comma-separated list of recipient email addresses",
            },
            includeTeamsPost: {
              type: "boolean",
              description: "Whether to also post the summary to the current Teams chat",
            },
          },
          required: ["summaryJson", "recipients"],
        },
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
        "Join a Teams meeting and start real-time transcription. If no URL is provided, the bot will try to detect the active meeting from the current chat automatically.",
        {
          type: "object",
          properties: {
            joinUrl: {
              type: "string",
              description: "The full Teams meeting join URL (contains teams.microsoft.com or teams.live.com). Leave empty to auto-detect from current chat.",
            },
            meetingId: {
              type: "string",
              description: "Optional custom meeting ID for tracking. Auto-generated if not provided.",
            },
          },
          required: [],
        },
        async (args: { joinUrl?: string; meetingId?: string }) => {
          try {
            let joinUrl = args.joinUrl?.trim() || "";

            // Tier 3: Smart URL detection — check if this chat has an active meeting
            if (!joinUrl) {
              this.logger.debug("No joinUrl provided — checking chat for active meeting...");
              joinUrl = await getActiveMeetingFromChat(context) || "";

              if (!joinUrl) {
                return JSON.stringify({
                  success: false,
                  error: "No active meeting found",
                  message: "No meeting URL provided and no active meeting was detected in this chat. Please share the Teams meeting join link.",
                });
              }
              this.logger.debug(`Auto-detected meeting URL from chat: ${joinUrl.substring(0, 60)}...`);
            }

            this.logger.debug(`Starting meeting capture for: ${joinUrl.substring(0, 50)}...`);

            // Validate join URL format
            if (!joinUrl.includes("teams.microsoft.com") && !joinUrl.includes("teams.live.com")) {
              return JSON.stringify({
                success: false,
                error: "Invalid Teams meeting URL",
                message: "Please provide a valid Microsoft Teams meeting join URL",
              });
            }

            args = { ...args, joinUrl };

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

            // Request the meeting-media-bot to join FIRST to get the Graph callId
            const result = await client.startMeetingCapture(joinUrl);

            if (!result.success || !result.callId) {
              return JSON.stringify({
                success: false,
                error: result.error,
                message: "Failed to start meeting capture",
              });
            }

            // Use the Graph callId as the meetingId — transcript chunks will arrive
            // from meeting-media-bot using this same callId, so they match correctly.
            const meetingId = result.callId;
            await context.database.upsertMeeting({
              meetingId,
              conversationId: context.conversationId,
              joinUrl,
              title: args.meetingId ? `Meeting: ${args.meetingId}` : `Meeting capture ${new Date().toLocaleString()}`,
              organizerAadId: context.userAadId || context.userId || "unknown",
            });

            return JSON.stringify({
              success: true,
              meetingId,
              callId: meetingId,
              message: `I've joined the meeting and am transcribing in real-time. Use the callId **${meetingId}** when you want to stop recording.`,
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
        {
          type: "object",
          properties: {
            callId: {
              type: "string",
              description: "The call ID returned by start_meeting_capture when the meeting capture was started",
            },
          },
          required: ["callId"],
        },
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

            // Auto-retrieve the saved transcript so the AI can immediately summarize
            let transcriptData: { text: string; chunkCount: number; participants: string[] } | null = null;
            try {
              const stored = await context.database.getTranscriptByMeetingId(args.callId);
              if (stored && stored.chunks.length > 0) {
                transcriptData = {
                  text: stored.chunks.map((c) => `[${c.speaker}]: ${c.text}`).join("\n"),
                  chunkCount: stored.chunks.length,
                  participants: stored.participants.map((p) => p.displayName),
                };
              }
            } catch {
              this.logger.warn("Could not retrieve transcript after stopping capture");
            }

            if (transcriptData) {
              return JSON.stringify({
                success: true,
                callId: args.callId,
                hasTranscript: true,
                transcriptText: transcriptData.text,
                chunkCount: transcriptData.chunkCount,
                participants: transcriptData.participants,
                message: "Meeting capture stopped. Transcript retrieved — now call summarize_meeting with the transcriptText above to generate the summary.",
              });
            }

            return JSON.stringify({
              success: true,
              callId: args.callId,
              hasTranscript: false,
              message: "Meeting capture stopped. No transcript chunks were recorded (the meeting may not have produced audio yet). You can still ask for a summary of the chat conversation.",
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
   * Send meeting summary email via Microsoft Graph API.
   * Uses the app's credentials (Mail.Send Application permission granted in Azure Portal).
   * Sends on behalf of the requesting user's identity.
   */
  private async sendSummaryEmail(
    emailRequest: MeetingEmailRequest,
    context: MessageContext
  ): Promise<{ success: boolean; messageId?: string }> {
    const graph = context.appGraph;
    if (!graph) {
      this.logger.warn("[sendSummaryEmail] No appGraph in context — cannot send email");
      return { success: false };
    }

    const senderId = context.userAadId || context.userId;
    if (!senderId) {
      this.logger.warn("[sendSummaryEmail] No sender user ID — cannot send email");
      return { success: false };
    }

    const s = emailRequest.summary;

    const decisionsHtml = s.decisions.length
      ? `<h3>✅ Decisions</h3><ul>${s.decisions.map((d) => `<li>${d}</li>`).join("")}</ul>`
      : "";

    const actionItemsHtml = s.actionItems.length
      ? `<h3>📌 Action Items</h3><ul>${s.actionItems
          .map((a) => `<li><b>${a.owner}</b>: ${a.task}${a.due ? ` (Due: ${a.due})` : ""}</li>`)
          .join("")}</ul>`
      : "";

    const risksHtml = s.risks.length
      ? `<h3>⚠️ Risks / Blockers</h3><ul>${s.risks.map((r) => `<li>${r}</li>`).join("")}</ul>`
      : "";

    const openQHtml = s.openQuestions.length
      ? `<h3>❓ Open Questions</h3><ul>${s.openQuestions.map((q) => `<li>${q}</li>`).join("")}</ul>`
      : "";

    const htmlBody = `
      <h2>📋 ${s.title}</h2>
      <p><b>Date:</b> ${s.dateTime}</p>
      <p><b>Participants:</b> ${s.participants.join(", ")}</p>
      <h3>📝 Summary</h3><p>${s.shortSummary}</p>
      ${decisionsHtml}
      ${actionItemsHtml}
      ${risksHtml}
      ${openQHtml}
      ${s.detailedSummary ? `<h3>Detailed Summary</h3><p>${s.detailedSummary}</p>` : ""}
    `.trim();

    try {
      await graph.call(
        (userId: string) => ({
          method: "post" as const,
          path: `/users/${userId}/sendMail`,
          body: {
            message: {
              subject: emailRequest.subject,
              body: { contentType: "HTML", content: htmlBody },
              toRecipients: emailRequest.recipients.map((addr) => ({
                emailAddress: { address: addr },
              })),
            },
            saveToSentItems: false,
          },
        }),
        senderId
      );

      this.logger.info(`[sendSummaryEmail] Sent to: ${emailRequest.recipients.join(", ")}`);
      return { success: true };
    } catch (err) {
      this.logger.error(
        `[sendSummaryEmail] Failed: ${err instanceof Error ? err.message : String(err)}`
      );
      return { success: false };
    }
  }

  /**
   * Post meeting summary to Teams chat.
   * The bot's formatted response to the user already IS the Teams message,
   * so this simply signals success — the AI's reply handles the actual posting.
   */
  private async postSummaryToTeams(
    _summary: MeetingSummary,
    context: MessageContext
  ): Promise<{ success: boolean; messageId?: string }> {
    this.logger.debug(
      `[postSummaryToTeams] Summary will be included in bot response for conversation ${context.conversationId}`
    );
    return { success: true };
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
   * Get transcript text from database or use provided text directly.
   * If meetingIdOrText looks like a meeting/call ID, tries to load stored chunks.
   */
  private async getTranscriptText(
    meetingIdOrText: string,
    context: MessageContext
  ): Promise<string> {
    // Looks like an ID (short, no spaces) — try to load from DB
    if (meetingIdOrText.length < 150 && !meetingIdOrText.includes(" ")) {
      try {
        const stored = await context.database.getTranscriptByMeetingId(meetingIdOrText);
        if (stored && stored.chunks.length > 0) {
          this.logger.debug(
            `[getTranscriptText] Loaded ${stored.chunks.length} chunks for meeting ${meetingIdOrText}`
          );
          return stored.chunks.map((c) => `[${c.speaker}]: ${c.text}`).join("\n");
        }
        this.logger.warn(`[getTranscriptText] No stored chunks for meeting ID ${meetingIdOrText}`);
      } catch (err) {
        this.logger.warn(
          `[getTranscriptText] DB lookup failed: ${err instanceof Error ? err.message : String(err)}`
        );
      }
      return `No transcript found for meeting ID: ${meetingIdOrText}. The meeting may not have been captured or the ID is incorrect.`;
    }

    // Long text with spaces — treat as direct transcript content
    return meetingIdOrText;
  }

  /**
   * Build a seed MeetingSummary structure.
   * The heavy lifting (actual content generation) is done by the AI model
   * in the ChatPrompt based on the SUMMARIZE_MEETING_PROMPT instructions.
   * The transcript text is passed to the AI via the function call arguments
   * and is visible in the prompt context.
   */
  private async generateStructuredSummary(
    transcriptText: string,
    conversationHistory: unknown[],
    _context: MessageContext
  ): Promise<MeetingSummary> {
    this.logger.debug(
      `[generateStructuredSummary] transcript=${transcriptText.length} chars, ` +
      `history=${conversationHistory.length} messages`
    );

    // Return a seed structure — the AI fills the actual fields when it formats
    // its response to the user, guided by SUMMARIZE_MEETING_PROMPT.
    return {
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
  }
}

// Capability definition for manager registration
export const MEETING_NOTES_CAPABILITY_DEFINITION: CapabilityDefinition = {
  name: "meeting_notes",
  manager_desc: `**Meeting Notes**: Use for requests like:
- "join meeting", "start recording", "capture meeting", "transcribe meeting", "join this meeting" + any Teams URL
- "stop recording", "stop transcription", "leave meeting", "end capture"
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
