'use strict';

var identity = require('@azure/identity');
var teams_apps = require('@microsoft/teams.apps');
var teams_common = require('@microsoft/teams.common');
var teams_dev = require('@microsoft/teams.dev');
var express = require('express');
var teams_ai = require('@microsoft/teams.ai');
var teams_openai = require('@microsoft/teams.openai');
var teams_api = require('@microsoft/teams.api');
var chrono = require('chrono-node');
var mssql = require('mssql');
var Database = require('better-sqlite3');
var path = require('path');

function _interopDefault (e) { return e && e.__esModule ? e : { default: e }; }

function _interopNamespace(e) {
  if (e && e.__esModule) return e;
  var n = Object.create(null);
  if (e) {
    Object.keys(e).forEach(function (k) {
      if (k !== 'default') {
        var d = Object.getOwnPropertyDescriptor(e, k);
        Object.defineProperty(n, k, d.get ? d : {
          enumerable: true,
          get: function () { return e[k]; }
        });
      }
    });
  }
  n.default = e;
  return Object.freeze(n);
}

var express__default = /*#__PURE__*/_interopDefault(express);
var chrono__namespace = /*#__PURE__*/_interopNamespace(chrono);
var mssql__namespace = /*#__PURE__*/_interopNamespace(mssql);
var Database__default = /*#__PURE__*/_interopDefault(Database);
var path__default = /*#__PURE__*/_interopDefault(path);

var __defProp = Object.defineProperty;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __esm = (fn, res) => function __init() {
  return fn && (res = (0, fn[__getOwnPropNames(fn)[0]])(fn = 0)), res;
};
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};

// src/utils/config.ts
var config_exports = {};
__export(config_exports, {
  AI_MODELS: () => AI_MODELS,
  DATABASE_CONFIG: () => DATABASE_CONFIG,
  getAppConfig: () => getAppConfig,
  getModelConfig: () => getModelConfig,
  loadConfig: () => loadConfig,
  logModelConfigs: () => logModelConfigs,
  validateEnvironment: () => validateEnvironment
});
function getModelConfig(capabilityType) {
  switch (capabilityType.toLowerCase()) {
    case "manager":
      return AI_MODELS.MANAGER;
    case "summarizer":
      return AI_MODELS.SUMMARIZER;
    case "actionitems":
      return AI_MODELS.ACTION_ITEMS;
    case "search":
      return AI_MODELS.SEARCH;
    case "meetingnotes":
      return AI_MODELS.MEETING_NOTES;
    default:
      return AI_MODELS.DEFAULT;
  }
}
function validateEnvironment(logger2) {
  const hasAoaiKey = process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY;
  const hasAoaiEndpoint = process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT;
  if (!hasAoaiKey || !hasAoaiEndpoint) {
    throw new Error(`Missing required environment variables: ${[
      !hasAoaiKey && "AOAI_API_KEY / SECRET_AZURE_OPENAI_API_KEY",
      !hasAoaiEndpoint && "AOAI_ENDPOINT / AZURE_OPENAI_ENDPOINT"
    ].filter(Boolean).join(", ")}`);
  }
  if (DATABASE_CONFIG.type === "mssql") {
    const sqlRequiredVars = ["SQL_CONNECTION_STRING"];
    const sqlMissing = sqlRequiredVars.filter((envVar) => !process.env[envVar]);
    if (sqlMissing.length > 0) {
      logger2.warn(
        `SQL Server configuration incomplete. Missing: ${sqlMissing.join(
          ", "
        )}. Falling back to SQLite.`
      );
      DATABASE_CONFIG.type = "sqlite";
    } else {
      logger2.debug("\u2705 SQL Server configuration validated");
    }
  }
  logger2.debug(`\u{1F4E6} Using database: ${DATABASE_CONFIG.type}`);
  logger2.debug("\u2705 Environment validation passed");
}
function logModelConfigs(logger2) {
  logger2.debug("\u{1F527} AI Model Configuration:");
  logger2.debug(`  Manager Capability: ${AI_MODELS.MANAGER.model}`);
  logger2.debug(`  Summarizer Capability: ${AI_MODELS.SUMMARIZER.model}`);
  logger2.debug(`  Action Items Capability: ${AI_MODELS.ACTION_ITEMS.model}`);
  logger2.debug(`  Search Capability: ${AI_MODELS.SEARCH.model}`);
  logger2.debug(`  Meeting Notes Capability: ${AI_MODELS.MEETING_NOTES.model}`);
  logger2.debug(`  Default Model: ${AI_MODELS.DEFAULT.model}`);
}
function loadConfig(logger2) {
  if (appConfigInstance) {
    return appConfigInstance;
  }
  const config = {
    botEndpoint: process.env.BOT_ENDPOINT || `http://localhost:${process.env.PORT || 3978}`,
    meetingMediaBotUrl: process.env.MEETING_MEDIA_BOT_URL || "http://localhost:4000",
    meetingMediaBotSharedSecret: process.env.MEETING_MEDIA_BOT_SHARED_SECRET || "dev-secret",
    databaseType: DATABASE_CONFIG.type,
    speechKey: process.env.AZURE_SPEECH_KEY,
    speechRegion: process.env.AZURE_SPEECH_REGION
  };
  logger2?.debug("\u{1F4CB} App Configuration loaded:");
  logger2?.debug(`  Bot Endpoint: ${config.botEndpoint}`);
  logger2?.debug(`  Meeting Media Bot URL: ${config.meetingMediaBotUrl}`);
  logger2?.debug(`  Database Type: ${config.databaseType}`);
  appConfigInstance = config;
  return config;
}
function getAppConfig() {
  if (!appConfigInstance) {
    throw new Error("App config not loaded. Call loadConfig() first.");
  }
  return appConfigInstance;
}
var DATABASE_CONFIG, AI_MODELS, appConfigInstance;
var init_config = __esm({
  "src/utils/config.ts"() {
    DATABASE_CONFIG = {
      type: process.env.RUNNING_ON_AZURE === "1" ? "mssql" : "sqlite",
      connectionString: process.env.SQL_CONNECTION_STRING,
      server: process.env.SQL_SERVER,
      database: process.env.SQL_DATABASE,
      username: process.env.SQL_USERNAME,
      password: process.env.SQL_PASSWORD,
      sqlitePath: process.env.CONVERSATIONS_DB_PATH
    };
    AI_MODELS = {
      // Manager Capability - Uses lighter, faster model for routing decisions
      MANAGER: {
        model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o-mini",
        apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY,
        endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT,
        apiVersion: "2025-04-01-preview"
      },
      // Summarizer Capability - Uses more capable model for complex analysis
      SUMMARIZER: {
        model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
        apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY,
        endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT,
        apiVersion: "2025-04-01-preview"
      },
      // Action Items Capability - Uses capable model for analysis and task management
      ACTION_ITEMS: {
        model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
        apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY,
        endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT,
        apiVersion: "2025-04-01-preview"
      },
      // Search Capability - Uses capable model for semantic search and deep linking
      SEARCH: {
        model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
        apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY,
        endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT,
        apiVersion: "2025-04-01-preview"
      },
      // Meeting Notes Capability - Uses capable model for transcript analysis and structured summaries
      MEETING_NOTES: {
        model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
        apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY,
        endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT,
        apiVersion: "2025-04-01-preview"
      },
      // Default model configuration (fallback)
      DEFAULT: {
        model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
        apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY,
        endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT,
        apiVersion: "2025-04-01-preview"
      }
    };
    appConfigInstance = null;
  }
});

// src/capabilities/capability.ts
init_config();
var BaseCapability = class {
  constructor(logger2) {
    this.logger = logger2;
  }
  /**
   * Default implementation of processRequest that creates a prompt and sends the request
   */
  async processRequest(context) {
    try {
      const prompt = this.createPrompt(context);
      const response = await prompt.send(context.text);
      return {
        response: response.content || "No response generated"
      };
    } catch (error) {
      return {
        response: "",
        error: error instanceof Error ? error.message : "Unknown error"
      };
    }
  }
  /**
   * Helper method to get model configuration
   */
  getModelConfig(configKey) {
    return getModelConfig(configKey);
  }
};

// src/capabilities/actionItems/prompt.ts
var ACTION_ITEMS_PROMPT = `
You are the Action Items capability of the Missa bot. Your role is to analyze team conversations and extract a list of clear action items based on what people said.

<GOAL>
Your job is to generate a concise, readable list of action items mentioned in the conversation. Focus on identifying:
- What needs to be done
- Who will do it (if mentioned)

<EXAMPLES OF ACTION ITEM CLUES>
- "I'll take care of this"
- "Can you follow up on..."
- "Let's finish this by tomorrow"
- "We still need to decide..."
- "Assign this to Alex"
- "We should check with finance"

<OUTPUT FORMAT>
- Return a plain text list of bullet points
- Each item should include a clear task and a person (if known)

<EXAMPLE OUTPUT>
- \u2705 Sarah will create the draft proposal by Friday
- \u2705 Alex will check budget numbers before the meeting
- \u2705 Follow up with IT on access issues
- \u2705 Decide final presenters by end of week

<NOTES>
- If no one is assigned, just describe the task
- Skip greetings or summary text \u2014 just the action items
- Do not assign tasks unless the conversation suggests it

Be clear, helpful, and concise.
`;

// src/capabilities/actionItems/actionItems.ts
var ActionItemsCapability = class extends BaseCapability {
  name = "action_items";
  createPrompt(context) {
    const actionItemsModelConfig = this.getModelConfig("actionItems");
    const prompt = new teams_ai.ChatPrompt({
      instructions: ACTION_ITEMS_PROMPT,
      model: new teams_openai.OpenAIChatModel({
        model: actionItemsModelConfig.model,
        apiKey: actionItemsModelConfig.apiKey,
        endpoint: actionItemsModelConfig.endpoint,
        apiVersion: actionItemsModelConfig.apiVersion
      })
    }).function(
      "generate_action_items",
      "Generate a list of action items based on the conversation",
      async () => {
        const allMessages = await context.memory.getMessagesByTimeRange(
          context.startTime,
          context.endTime
        );
        return JSON.stringify({
          messages: allMessages.map((msg) => ({
            timestamp: msg.timestamp,
            name: msg.name,
            content: msg.content
          }))
        });
      }
    );
    this.logger.debug(
      `Initialized Action Items Capability using ${context.members.length} members from context`
    );
    return prompt;
  }
};
var ACTION_ITEMS_CAPABILITY_DEFINITION = {
  name: "action_items",
  manager_desc: `**Action Items**: Use for requests like:
- "next steps", "to-do", "assign task", "my tasks", "what needs to be done"`,
  handler: async (context, logger2) => {
    const actionItemsCapability = new ActionItemsCapability(logger2);
    const result = await actionItemsCapability.processRequest(context);
    if (result.error) {
      logger2.error(`Error in Action Items Capability: ${result.error}`);
      return `Error in Action Items Capability: ${result.error}`;
    }
    return result.response || "No response from Action Items Capability";
  }
};

// src/services/meetingMediaBotClient.ts
init_config();
var MeetingMediaBotClient = class {
  logger;
  baseUrl;
  sharedSecret;
  constructor(logger2) {
    this.logger = logger2;
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
  async startMeetingCapture(joinUrl, meetingId) {
    this.logger.debug(`Requesting meeting capture for: ${joinUrl.substring(0, 50)}...`);
    try {
      const response = await fetch(`${this.baseUrl}/api/meetings/join`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shared-Secret": this.sharedSecret
        },
        body: JSON.stringify({
          joinUrl,
          meetingId
        })
      });
      const data = await response.json();
      if (!response.ok) {
        this.logger.error(`Failed to start meeting capture: ${data.error}`);
        return {
          success: false,
          error: data.error || `HTTP ${response.status}`
        };
      }
      this.logger.info(`Meeting capture started, callId: ${data.callId}`);
      return data;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`Error starting meeting capture: ${errorMessage}`);
      return {
        success: false,
        error: errorMessage
      };
    }
  }
  /**
   * Request the meeting-media-bot to leave a meeting
   * 
   * @param callId - The call ID returned from startMeetingCapture
   * @returns Result indicating success or failure
   */
  async stopMeetingCapture(callId) {
    this.logger.debug(`Requesting to stop meeting capture for callId: ${callId}`);
    try {
      const response = await fetch(`${this.baseUrl}/api/meetings/leave`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Shared-Secret": this.sharedSecret
        },
        body: JSON.stringify({ callId })
      });
      const data = await response.json();
      if (!response.ok) {
        this.logger.error(`Failed to stop meeting capture: ${data.error}`);
        return {
          success: false,
          error: data.error || `HTTP ${response.status}`
        };
      }
      this.logger.info(`Meeting capture stopped for callId: ${callId}`);
      return data;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`Error stopping meeting capture: ${errorMessage}`);
      return {
        success: false,
        error: errorMessage
      };
    }
  }
  /**
   * Get the status of an active meeting capture
   * 
   * @param callId - The call ID to check
   * @returns Meeting status or error
   */
  async getMeetingCaptureStatus(callId) {
    this.logger.debug(`Getting status for callId: ${callId}`);
    try {
      const response = await fetch(`${this.baseUrl}/api/meetings/${callId}/status`, {
        method: "GET",
        headers: {
          "X-Shared-Secret": this.sharedSecret
        }
      });
      if (!response.ok) {
        const data = await response.json();
        return {
          success: false,
          error: data.error || `HTTP ${response.status}`
        };
      }
      const status = await response.json();
      return {
        success: true,
        status
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";
      this.logger.error(`Error getting meeting status: ${errorMessage}`);
      return {
        success: false,
        error: errorMessage
      };
    }
  }
  /**
   * Check if the meeting-media-bot service is available
   * Retries once after a delay to handle Azure cold starts
   */
  async checkHealth() {
    const url = `${this.baseUrl}/api/health`;
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        this.logger.info(`[HealthCheck] Attempt ${attempt}: ${url}`);
        const response = await fetch(url, {
          method: "GET",
          signal: AbortSignal.timeout(15e3)
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
        await new Promise((r) => setTimeout(r, 3e3));
      }
    }
    this.logger.error(`[HealthCheck] Meeting media bot unreachable after 2 attempts: ${url}`);
    return false;
  }
};
var clientInstance = null;
function getMeetingMediaBotClient(logger2) {
  if (!clientInstance) {
    clientInstance = new MeetingMediaBotClient(logger2);
  }
  return clientInstance;
}

// src/capabilities/meeting-notes/prompt.ts
var MEETING_NOTES_BASE_PROMPT = `
You are the Meeting Notes capability of the Missa bot. Your role is to help users manage meeting transcripts, generate structured summaries, and distribute meeting notes to participants.

You have access to five main commands:
1. **start_meeting_capture**: Join a Teams meeting and start real-time transcription
2. **stop_meeting_capture**: Leave a Teams meeting and stop transcription
3. **read_transcript**: Fetch and store a meeting transcript from Microsoft Graph
4. **summarize_meeting**: Generate a structured JSON summary from a transcript or conversation
5. **send_summary**: Email the summary to participants and optionally post to Teams chat

You must be precise, professional, and extract actionable insights from meetings.
`;
var MEETING_NOTES_PROMPT = `
${MEETING_NOTES_BASE_PROMPT}

<ROUTING LOGIC>
Determine which command the user wants:
- Keywords like "start capture", "join meeting", "start transcribing", "capture this meeting", "record meeting" + meeting URL \u2192 start_meeting_capture
- Keywords like "stop capture", "leave meeting", "stop transcribing", "stop recording" \u2192 stop_meeting_capture  
- Keywords like "fetch", "get", "retrieve", "read transcript" \u2192 read_transcript
- Keywords like "summarize", "summary", "create notes", "analyze" \u2192 summarize_meeting
- Keywords like "send", "email", "share", "distribute" \u2192 send_summary

<MEETING CAPTURE GUIDANCE>
When users want to start meeting capture ("join meeting", "join this meeting", "start recording", "capture this"):
1. Look for a Teams meeting URL in their message (contains teams.microsoft.com or teams.live.com)
2. If URL found, call start_meeting_capture with the joinUrl
3. If NO URL found, STILL call start_meeting_capture with an empty joinUrl \u2014 the system will automatically detect the active meeting from the chat context (like Otter.ai does)
4. Only ask the user for a URL if the automatic detection also fails (the function will return success: false with an appropriate error message)

When stopping capture:
1. If the user says "stop capture", "stop recording", or "leave meeting", check if you know the callId from a previous start
2. If you don't have the callId, ask the user for it
3. After calling stop_meeting_capture, check the response:
   - If hasTranscript=true, IMMEDIATELY call summarize_meeting with the transcriptText from the response
   - If hasTranscript=false, let the user know and offer to summarize the chat conversation instead

Example user requests:
- "Start capturing https://teams.microsoft.com/l/meetup-join/..." \u2192 Extract URL, call start_meeting_capture
- "Join this meeting and take notes: <URL>" \u2192 Extract URL, call start_meeting_capture  
- "Stop the meeting capture" \u2192 Ask for callId if unknown, then call stop_meeting_capture

<GENERAL BEHAVIOR>
- Be proactive: After reading a transcript, offer to summarize
- Be proactive: After stopping capture, offer to summarize the transcript
- Be helpful: After summarizing, offer to send
- Be precise: Use structured JSON for summaries
- Be professional: Meeting notes should be clear and actionable
`;

// src/capabilities/meeting-notes/meeting-notes.ts
async function getActiveMeetingFromChat(context) {
  try {
    const chatId = context.conversationId;
    if (!chatId || !context.appGraph) return null;
    const response = await context.appGraph.call(
      (_userId) => ({
        method: "get",
        path: `/chats/${encodeURIComponent(chatId)}`,
        query: { "$select": "onlineMeetingInfo" }
      }),
      context.userAadId || context.userId || ""
    );
    const joinWebUrl = response?.onlineMeetingInfo;
    return joinWebUrl?.joinWebUrl || null;
  } catch {
    return null;
  }
}
var MeetingNotesCapability = class extends BaseCapability {
  name = "meeting_notes";
  createPrompt(context) {
    const meetingNotesModelConfig = this.getModelConfig("meetingNotes");
    const prompt = new teams_ai.ChatPrompt({
      instructions: MEETING_NOTES_PROMPT,
      model: new teams_openai.OpenAIChatModel({
        model: meetingNotesModelConfig.model,
        apiKey: meetingNotesModelConfig.apiKey,
        endpoint: meetingNotesModelConfig.endpoint,
        apiVersion: meetingNotesModelConfig.apiVersion
      })
    }).function(
      "read_transcript",
      "Fetch a meeting transcript from Microsoft Graph and store it in the database",
      {
        type: "object",
        properties: {
          meetingId: {
            type: "string",
            description: "The meeting ID or join URL to fetch the transcript for"
          }
        },
        required: ["meetingId"]
      },
      async ({ meetingId }) => {
        try {
          this.logger.debug(`Fetching transcript for meeting: ${meetingId}`);
          const transcript = await this.fetchMeetingTranscript(meetingId, context);
          await this.storeMeetingTranscript(transcript, context);
          return JSON.stringify({
            success: true,
            meeting: {
              id: transcript.meetingId,
              title: transcript.title,
              startDateTime: transcript.startDateTime,
              endDateTime: transcript.endDateTime,
              participants: transcript.participants,
              transcriptLength: transcript.transcriptText.length
            },
            message: "Transcript retrieved and stored successfully"
          });
        } catch (error) {
          this.logger.error("Error fetching transcript:", error);
          return JSON.stringify({
            success: false,
            error: error instanceof Error ? error.message : "Unknown error",
            message: "Failed to fetch meeting transcript"
          });
        }
      }
    ).function(
      "summarize_meeting",
      "Generate a structured JSON summary from a meeting transcript or conversation history. Returns JSON matching the MeetingSummary schema.",
      {
        type: "object",
        properties: {
          meetingIdOrTranscriptText: {
            type: "string",
            description: "Meeting ID to look up stored transcript, or raw transcript text to summarize directly"
          }
        },
        required: ["meetingIdOrTranscriptText"]
      },
      async ({ meetingIdOrTranscriptText }) => {
        try {
          this.logger.debug("Generating meeting summary");
          const transcriptText = await this.getTranscriptText(
            meetingIdOrTranscriptText,
            context
          );
          const conversationHistory = await context.memory.getMessagesByTimeRange(
            context.startTime,
            context.endTime
          );
          const summary = await this.generateStructuredSummary(
            transcriptText,
            conversationHistory,
            context
          );
          await this.storeMeetingSummary(summary, context);
          return JSON.stringify({
            success: true,
            summary,
            message: "Meeting summary generated and stored successfully"
          });
        } catch (error) {
          this.logger.error("Error generating summary:", error);
          return JSON.stringify({
            success: false,
            error: error instanceof Error ? error.message : "Unknown error",
            message: "Failed to generate meeting summary"
          });
        }
      }
    ).function(
      "send_summary",
      "Send meeting summary to participants via email and optionally post to Teams chat",
      {
        type: "object",
        properties: {
          summaryJson: {
            type: "string",
            description: "JSON string of the MeetingSummary object to send"
          },
          recipients: {
            type: "string",
            description: "Comma-separated list of recipient email addresses"
          },
          includeTeamsPost: {
            type: "boolean",
            description: "Whether to also post the summary to the current Teams chat"
          }
        },
        required: ["summaryJson", "recipients"]
      },
      async (args) => {
        try {
          this.logger.debug("Sending meeting summary");
          const summary = JSON.parse(args.summaryJson);
          const recipientList = args.recipients.split(",").map((r) => r.trim());
          const includeTeamsPost = args.includeTeamsPost || false;
          const emailRequest = {
            recipients: recipientList,
            subject: `Meeting Notes: ${summary.title}`,
            summary,
            includeTeamsPost,
            conversationId: includeTeamsPost ? context.conversationId : void 0
          };
          const emailResult = await this.sendSummaryEmail(emailRequest, context);
          let teamsPostResult = null;
          if (includeTeamsPost) {
            teamsPostResult = await this.postSummaryToTeams(summary, context);
          }
          await this.recordSummarySent(summary, recipientList, context);
          return JSON.stringify({
            success: true,
            emailSent: emailResult.success,
            emailRecipients: recipientList,
            teamsPostSent: teamsPostResult?.success || false,
            message: `Summary sent to ${recipientList.length} recipient(s)`
          });
        } catch (error) {
          this.logger.error("Error sending summary:", error);
          return JSON.stringify({
            success: false,
            error: error instanceof Error ? error.message : "Unknown error",
            message: "Failed to send meeting summary"
          });
        }
      }
    ).function(
      "start_meeting_capture",
      "Join a Teams meeting and start real-time transcription. If no URL is provided, the bot will try to detect the active meeting from the current chat automatically.",
      {
        type: "object",
        properties: {
          joinUrl: {
            type: "string",
            description: "The full Teams meeting join URL (contains teams.microsoft.com or teams.live.com). Leave empty to auto-detect from current chat."
          },
          meetingId: {
            type: "string",
            description: "Optional custom meeting ID for tracking. Auto-generated if not provided."
          }
        },
        required: []
      },
      async (args) => {
        try {
          let joinUrl = args.joinUrl?.trim() || "";
          if (!joinUrl) {
            this.logger.debug("No joinUrl provided \u2014 checking chat for active meeting...");
            joinUrl = await getActiveMeetingFromChat(context) || "";
            if (!joinUrl) {
              return JSON.stringify({
                success: false,
                error: "No active meeting found",
                message: "No meeting URL provided and no active meeting was detected in this chat. Please share the Teams meeting join link."
              });
            }
            this.logger.debug(`Auto-detected meeting URL from chat: ${joinUrl.substring(0, 60)}...`);
          }
          this.logger.debug(`Starting meeting capture for: ${joinUrl.substring(0, 50)}...`);
          if (!joinUrl.includes("teams.microsoft.com") && !joinUrl.includes("teams.live.com")) {
            return JSON.stringify({
              success: false,
              error: "Invalid Teams meeting URL",
              message: "Please provide a valid Microsoft Teams meeting join URL"
            });
          }
          args = { ...args, joinUrl };
          const client = getMeetingMediaBotClient(this.logger);
          const isAvailable = await client.checkHealth();
          if (!isAvailable) {
            const config = await Promise.resolve().then(() => (init_config(), config_exports));
            const appConfig = config.getAppConfig();
            this.logger.error(`Meeting media bot unreachable at: ${appConfig.meetingMediaBotUrl}`);
            return JSON.stringify({
              success: false,
              error: "Meeting capture service unavailable",
              message: `The meeting capture service at ${appConfig.meetingMediaBotUrl} is not responding. Please verify the service is running and try again.`
            });
          }
          const result = await client.startMeetingCapture(joinUrl);
          if (!result.success || !result.callId) {
            return JSON.stringify({
              success: false,
              error: result.error,
              message: "Failed to start meeting capture"
            });
          }
          const meetingId = result.callId;
          await context.database.upsertMeeting({
            meetingId,
            conversationId: context.conversationId,
            joinUrl,
            title: args.meetingId ? `Meeting: ${args.meetingId}` : `Meeting capture ${(/* @__PURE__ */ new Date()).toLocaleString()}`,
            organizerAadId: context.userAadId || context.userId || "unknown"
          });
          return JSON.stringify({
            success: true,
            meetingId,
            callId: meetingId,
            message: `I've joined the meeting and am transcribing in real-time. Use the callId **${meetingId}** when you want to stop recording.`
          });
        } catch (error) {
          this.logger.error("Error starting meeting capture:", error);
          return JSON.stringify({
            success: false,
            error: error instanceof Error ? error.message : "Unknown error",
            message: "Failed to start meeting capture"
          });
        }
      }
    ).function(
      "stop_meeting_capture",
      "Leave a Teams meeting and stop real-time transcription. Requires the call ID from start_meeting_capture.",
      {
        type: "object",
        properties: {
          callId: {
            type: "string",
            description: "The call ID returned by start_meeting_capture when the meeting capture was started"
          }
        },
        required: ["callId"]
      },
      async (args) => {
        try {
          this.logger.debug(`Stopping meeting capture for callId: ${args.callId}`);
          const client = getMeetingMediaBotClient(this.logger);
          const result = await client.stopMeetingCapture(args.callId);
          if (!result.success) {
            return JSON.stringify({
              success: false,
              error: result.error,
              message: "Failed to stop meeting capture"
            });
          }
          await context.database.updateMeetingStatus({
            meetingId: args.callId,
            status: "ended",
            endedAt: (/* @__PURE__ */ new Date()).toISOString()
          });
          let transcriptData = null;
          try {
            const stored = await context.database.getTranscriptByMeetingId(args.callId);
            if (stored && stored.chunks.length > 0) {
              transcriptData = {
                text: stored.chunks.map((c) => `[${c.speaker}]: ${c.text}`).join("\n"),
                chunkCount: stored.chunks.length,
                participants: stored.participants.map((p) => p.displayName)
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
              message: "Meeting capture stopped. Transcript retrieved \u2014 now call summarize_meeting with the transcriptText above to generate the summary."
            });
          }
          return JSON.stringify({
            success: true,
            callId: args.callId,
            hasTranscript: false,
            message: "Meeting capture stopped. No transcript chunks were recorded (the meeting may not have produced audio yet). You can still ask for a summary of the chat conversation."
          });
        } catch (error) {
          this.logger.error("Error stopping meeting capture:", error);
          return JSON.stringify({
            success: false,
            error: error instanceof Error ? error.message : "Unknown error",
            message: "Failed to stop meeting capture"
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
  async fetchMeetingTranscript(meetingId, context) {
    this.logger.debug(`Fetching transcript for meeting: ${meetingId}`);
    try {
      const storedTranscript = await context.database.getTranscriptByMeetingId(meetingId);
      if (storedTranscript && storedTranscript.chunks.length > 0) {
        this.logger.debug(`Found stored transcript with ${storedTranscript.chunks.length} chunks`);
        const fullText = storedTranscript.chunks.map((chunk) => `[${chunk.speaker}]: ${chunk.text}`).join("\n");
        return {
          meetingId: storedTranscript.meetingId,
          title: storedTranscript.title || "Meeting Transcript",
          startDateTime: storedTranscript.startedAt || (/* @__PURE__ */ new Date()).toISOString(),
          endDateTime: storedTranscript.endedAt || (/* @__PURE__ */ new Date()).toISOString(),
          organizer: storedTranscript.participants[0]?.displayName || "Unknown",
          participants: storedTranscript.participants.map((p) => p.displayName),
          transcriptText: fullText,
          retrievedAt: (/* @__PURE__ */ new Date()).toISOString()
        };
      }
    } catch (error) {
      this.logger.warn(`No stored transcript found for ${meetingId}, will try Graph API fallback`);
    }
    this.logger.warn("Graph API transcript fallback not yet implemented - returning stub");
    return {
      meetingId,
      title: "Meeting (Graph API transcript pending)",
      startDateTime: (/* @__PURE__ */ new Date()).toISOString(),
      endDateTime: (/* @__PURE__ */ new Date()).toISOString(),
      organizer: "organizer@example.com",
      participants: ["participant1@example.com", "participant2@example.com"],
      transcriptText: "No transcript data available. The meeting may not have been captured yet, or Graph API access is pending configuration.",
      retrievedAt: (/* @__PURE__ */ new Date()).toISOString()
    };
  }
  /**
   * Send meeting summary email via Microsoft Graph API.
   * Uses the app's credentials (Mail.Send Application permission granted in Azure Portal).
   * Sends on behalf of the requesting user's identity.
   */
  async sendSummaryEmail(emailRequest, context) {
    const graph = context.appGraph;
    if (!graph) {
      this.logger.warn("[sendSummaryEmail] No appGraph in context \u2014 cannot send email");
      return { success: false };
    }
    const senderId = context.userAadId || context.userId;
    if (!senderId) {
      this.logger.warn("[sendSummaryEmail] No sender user ID \u2014 cannot send email");
      return { success: false };
    }
    const s = emailRequest.summary;
    const decisionsHtml = s.decisions.length ? `<h3>\u2705 Decisions</h3><ul>${s.decisions.map((d) => `<li>${d}</li>`).join("")}</ul>` : "";
    const actionItemsHtml = s.actionItems.length ? `<h3>\u{1F4CC} Action Items</h3><ul>${s.actionItems.map((a) => `<li><b>${a.owner}</b>: ${a.task}${a.due ? ` (Due: ${a.due})` : ""}</li>`).join("")}</ul>` : "";
    const risksHtml = s.risks.length ? `<h3>\u26A0\uFE0F Risks / Blockers</h3><ul>${s.risks.map((r) => `<li>${r}</li>`).join("")}</ul>` : "";
    const openQHtml = s.openQuestions.length ? `<h3>\u2753 Open Questions</h3><ul>${s.openQuestions.map((q) => `<li>${q}</li>`).join("")}</ul>` : "";
    const htmlBody = `
      <h2>\u{1F4CB} ${s.title}</h2>
      <p><b>Date:</b> ${s.dateTime}</p>
      <p><b>Participants:</b> ${s.participants.join(", ")}</p>
      <h3>\u{1F4DD} Summary</h3><p>${s.shortSummary}</p>
      ${decisionsHtml}
      ${actionItemsHtml}
      ${risksHtml}
      ${openQHtml}
      ${s.detailedSummary ? `<h3>Detailed Summary</h3><p>${s.detailedSummary}</p>` : ""}
    `.trim();
    try {
      await graph.call(
        (userId) => ({
          method: "post",
          path: `/users/${userId}/sendMail`,
          body: {
            message: {
              subject: emailRequest.subject,
              body: { contentType: "HTML", content: htmlBody },
              toRecipients: emailRequest.recipients.map((addr) => ({
                emailAddress: { address: addr }
              }))
            },
            saveToSentItems: false
          }
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
  async postSummaryToTeams(_summary, context) {
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
  async storeMeetingTranscript(transcript, context) {
    this.logger.debug(
      `Storing transcript for meeting ${transcript.meetingId} in conversation ${context.conversationId}`
    );
    this.logger.warn("TODO: Implement dedicated meeting_transcripts table");
  }
  /**
   * Store meeting summary in database
   * TODO: Extend database schema to support meeting_summaries table
   */
  async storeMeetingSummary(summary, context) {
    this.logger.debug(
      `Storing summary for meeting "${summary.title}" in conversation ${context.conversationId}`
    );
    this.logger.warn("TODO: Implement dedicated meeting_summaries table");
  }
  /**
   * Record that a summary was sent to recipients
   * TODO: Extend database schema to support meeting_summary_sends table
   */
  async recordSummarySent(summary, recipients, _context) {
    this.logger.debug(
      `Recording send action for summary "${summary.title}" to ${recipients.length} recipients`
    );
    this.logger.warn("TODO: Implement dedicated meeting_summary_sends tracking table");
  }
  /**
   * Get transcript text from database or use provided text directly.
   * If meetingIdOrText looks like a meeting/call ID, tries to load stored chunks.
   */
  async getTranscriptText(meetingIdOrText, context) {
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
    return meetingIdOrText;
  }
  /**
   * Build a seed MeetingSummary structure.
   * The heavy lifting (actual content generation) is done by the AI model
   * in the ChatPrompt based on the SUMMARIZE_MEETING_PROMPT instructions.
   * The transcript text is passed to the AI via the function call arguments
   * and is visible in the prompt context.
   */
  async generateStructuredSummary(transcriptText, conversationHistory, _context) {
    this.logger.debug(
      `[generateStructuredSummary] transcript=${transcriptText.length} chars, history=${conversationHistory.length} messages`
    );
    return {
      title: "Meeting Summary",
      dateTime: (/* @__PURE__ */ new Date()).toISOString(),
      participants: [],
      decisions: [],
      actionItems: [],
      risks: [],
      openQuestions: [],
      shortSummary: "",
      detailedSummary: ""
    };
  }
};
var MEETING_NOTES_CAPABILITY_DEFINITION = {
  name: "meeting_notes",
  manager_desc: `**Meeting Notes**: Use for requests like:
- "join meeting", "start recording", "capture meeting", "transcribe meeting", "join this meeting" + any Teams URL
- "stop recording", "stop transcription", "leave meeting", "end capture"
- "read transcript", "get meeting transcript", "fetch meeting notes"
- "summarize meeting", "create meeting summary", "analyze meeting"
- "send summary", "email notes", "share meeting notes", "distribute summary to participants"
- Managing meeting transcripts, summaries, action items from meetings`,
  handler: async (context, logger2) => {
    const meetingNotesCapability = new MeetingNotesCapability(logger2);
    const result = await meetingNotesCapability.processRequest(context);
    if (result.error) {
      logger2.error(`Error in Meeting Notes Capability: ${result.error}`);
      return `Error in Meeting Notes Capability: ${result.error}`;
    }
    return result.response || "No response from Meeting Notes Capability";
  }
};

// src/capabilities/search/prompt.ts
var SEARCH_PROMPT = `
You are the Search capability of the Missa bot. Your role is to help users find specific conversations or messages from their chat history.

You can search through message history to find:
- Conversations between specific people
- Messages about specific topics
- Messages from specific time periods (time ranges will be pre-calculated by the Manager)
- Messages containing specific keywords

When a user asks you to find something, use the search_messages function to search the database.

RESPONSE FORMAT:
- Your search_messages function returns just the text associated with the search results
- Focus on creating a helpful, conversational summary that complements the citations
- Be specific about what was found and provide context about timing and participants
- If no results are found, suggest alternative search terms or broader criteria

Be helpful and conversational in your responses. The user will see both your text response and interactive cards that let them jump to the original messages.
`;

// src/capabilities/search/schema.ts
var SEARCH_MESSAGES_SCHEMA = {
  type: "object",
  properties: {
    keywords: {
      type: "array",
      items: { type: "string" },
      description: "Keywords to search for in the message content"
    },
    participants: {
      type: "array",
      items: { type: "string" },
      description: "Optional: list of participant names to filter messages by who said them"
    },
    max_results: {
      type: "number",
      description: "Optional: maximum number of results to return (default is 5)"
    }
  },
  required: ["keywords"]
};

// src/capabilities/search/search.ts
var dateFormat = new Intl.DateTimeFormat("en-US");
var SearchCapability = class extends BaseCapability {
  name = "search";
  createPrompt(context) {
    const searchModelConfig = this.getModelConfig("search");
    const prompt = new teams_ai.ChatPrompt({
      instructions: SEARCH_PROMPT,
      model: new teams_openai.OpenAIChatModel({
        model: searchModelConfig.model,
        apiKey: searchModelConfig.apiKey,
        endpoint: searchModelConfig.endpoint,
        apiVersion: searchModelConfig.apiVersion
      })
    }).function(
      "search_messages",
      "Search the conversation for relevant messages",
      SEARCH_MESSAGES_SCHEMA,
      async ({ keywords, participants, max_results }) => {
        const selected = await context.memory.getFilteredMessages(
          context.conversationId,
          keywords,
          context.startTime,
          context.endTime,
          participants,
          max_results
        );
        this.logger.debug(selected);
        if (selected.length === 0) {
          return "No matching messages found.";
        }
        const citations = selected.map(
          (msg) => createCitationFromRecord(msg, context.conversationId)
        );
        context.citations.push(...citations);
        return selected.map((msg) => {
          const date = new Date(msg.timestamp).toLocaleString();
          const preview = msg.content.slice(0, 100);
          const citation = citations.find((c) => c.keywords?.includes(msg.name));
          const link = citation?.url || "#";
          return `\u2022 [${msg.name}](${link}) at ${date}: "${preview}"`;
        }).join("\n");
      }
    );
    this.logger.debug("Initialized Search Capability!");
    return prompt;
  }
};
function createDeepLink(activityId, conversationId) {
  const contextParam = encodeURIComponent(JSON.stringify({ contextType: "chat" }));
  return `https://teams.microsoft.com/l/message/${encodeURIComponent(
    conversationId
  )}/${activityId}?context=${contextParam}`;
}
function createCitationFromRecord(message, conversationId) {
  const date = new Date(message.timestamp);
  const formatted = dateFormat.format(date);
  const preview = message.content.length > 120 ? message.content.slice(0, 120) + "..." : message.content;
  const activityId = message.activity_id ?? message.name.replace(/\s+/g, "-").toLowerCase();
  const deepLink = createDeepLink(activityId, conversationId);
  return {
    name: `Message from ${message.name}`,
    url: deepLink,
    abstract: `${formatted}: "${preview}"`,
    keywords: [message.name]
  };
}
var SEARCH_CAPABILITY_DEFINITION = {
  name: "search",
  manager_desc: `**Search**: Use for:
- "find", "search", "show me", "conversation with", "where did [person] say", "messages from last week"`,
  handler: async (context, logger2) => {
    const searchCapability = new SearchCapability(logger2);
    const result = await searchCapability.processRequest(context);
    if (result.error) {
      logger2.error(`\u274C Error in Search Capability: ${result.error}`);
      return `Error in Search Capability: ${result.error}`;
    }
    return result.response || "No response from Search Capability";
  }
};

// src/capabilities/summarizer/prompt.ts
var SUMMARY_PROMPT = `
You are the Summarizer capability of the Missa bot that specializes in analyzing conversations between groups of people.
Your job is to retrieve and analyze conversation messages, then provide structured summaries with proper attribution.

<TIMEZONE AWARENESS>
The system uses the user's actual timezone from Microsoft Teams for all time calculations.
Time ranges will be pre-calculated by the Manager and passed to you as ISO timestamps when needed.

<INSTRUCTIONS>
1. Use the appropriate function to retrieve the messages you need based on the user's request
2. If time ranges are specified in the request, they will be pre-calculated and provided as ISO timestamps
3. If no specific timespan is mentioned, default to the last 24 hours using get_messages_by_time_range
4. Analyze the retrieved messages and identify participants and topics
5. Return a BRIEF summary with proper participant attribution
6. Include participant names in your analysis and summary points
7. Be concise and focus on the key topics discussed

<OUTPUT FORMAT>
- Use bullet points for main topics
- Include participant names when attributing ideas or statements
- Provide a brief overview if requested
`;

// src/capabilities/summarizer/summarize.ts
var SummarizerCapability = class extends BaseCapability {
  name = "summarizer";
  createPrompt(context) {
    const summarizerModelConfig = this.getModelConfig("summarizer");
    const prompt = new teams_ai.ChatPrompt({
      instructions: SUMMARY_PROMPT,
      model: new teams_openai.OpenAIChatModel({
        model: summarizerModelConfig.model,
        apiKey: summarizerModelConfig.apiKey,
        endpoint: summarizerModelConfig.endpoint,
        apiVersion: summarizerModelConfig.apiVersion
      })
    }).function("summarize_conversation", "Summarize the conversation history", async () => {
      const allMessages = await context.memory.getMessagesByTimeRange(
        context.startTime,
        context.endTime
      );
      return JSON.stringify({
        messages: allMessages.map((msg) => ({
          timestamp: msg.timestamp,
          name: msg.name,
          content: msg.content
        }))
      });
    });
    this.logger.debug("Initialized Summarizer Capability!");
    return prompt;
  }
};
var SUMMARIZER_CAPABILITY_DEFINITION = {
  name: "summarizer",
  manager_desc: `**Summarizer**: Use for keywords like:
- "summarize", "overview", "recap", "conversation history"
- "what did we discuss", "catch me up", "who said what", "recent messages"`,
  handler: async (context, logger2) => {
    const summarizerCapability = new SummarizerCapability(logger2);
    const result = await summarizerCapability.processRequest(context);
    if (result.error) {
      logger2.error(`Error in Summarizer Capability: ${result.error}`);
      return `Error in Summarizer Capability: ${result.error}`;
    }
    return result.response || "No response from Summarizer Capability";
  }
};

// src/capabilities/registry.ts
var CAPABILITY_DEFINITIONS = [
  SUMMARIZER_CAPABILITY_DEFINITION,
  ACTION_ITEMS_CAPABILITY_DEFINITION,
  SEARCH_CAPABILITY_DEFINITION,
  MEETING_NOTES_CAPABILITY_DEFINITION
];

// src/agent/manager.ts
init_config();
function finalizePromptResponse(text, context, logger2) {
  const messageActivity = new teams_api.MessageActivity(text).addAiGenerated().addFeedback();
  if (context.citations && context.citations.length > 0) {
    logger2.debug(`Adding ${context.citations.length} context.citations to message activity`);
    context.citations.forEach((citation, index) => {
      const citationNumber = index + 1;
      messageActivity.addCitation(citationNumber, citation);
      logger2.debug(`Citation number ${citationNumber}`);
      logger2.debug(citation);
      messageActivity.text += ` [${citationNumber}]`;
    });
  }
  return messageActivity;
}
function extractTimeRange(phrase, now = /* @__PURE__ */ new Date()) {
  const results = chrono__namespace.parse(phrase, now);
  if (!results.length || !results[0].start) {
    return null;
  }
  const { start, end } = results[0];
  const from = start.date();
  const to = end?.date() ?? new Date(from.getTime() + 24 * 60 * 60 * 1e3);
  return { from, to };
}
function createMessageRecords(activities) {
  if (!activities || activities.length === 0) return [];
  const conversation_id = activities[0].conversation.id;
  return activities.map((activity) => ({
    conversation_id,
    role: activity.entities?.some((e) => e.additionalType?.includes("AIGeneratedContent")) ? "model" : "user",
    content: activity.text?.replace(/<\/?at>/g, "") || "",
    timestamp: activity.timestamp?.toString() || (/* @__PURE__ */ new Date()).toISOString(),
    activity_id: activity.id,
    name: activity.from?.name || "Missa"
  }));
}

// src/agent/prompt.ts
function generateManagerPrompt(capabilities) {
  const namesList = capabilities.map((cap, i) => `${i + 1}. **${cap.name}**`).join("\n");
  const capabilityDescriptions = capabilities.map((cap) => `${cap.manager_desc}`).join("\n");
  return `
You are the Manager for Missa \u2014 a Microsoft Teams bot. You coordinate requests by deciding which specialized capability should handle each @mention.

<AVAILABLE CAPABILITIES>
${namesList}

<INSTRUCTIONS>
1. Analyze the request's intent and route it to the best-matching capability.
2. **If the request includes a time expression**, call calculate_time_range first using the exact phrase (e.g., "last week", "past 2 days").
3. If no capability applies, respond conversationally and describe what Missa *can* help with.

<WHEN TO USE EACH CAPABILITY>
Use the following descriptions to determine routing logic. Match based on intent, not just keywords.

${capabilityDescriptions}

<RESPONSE RULE>
When using a function call to delegate, return the capability\u2019s response **as-is**, with no added commentary or explanation. MAKE SURE TO NOT WRAP THE RESPONSE IN QUOTES.

\u2705 GOOD: [capability response]  
\u274C BAD: Here\u2019s what the Summarizer found: [capability response]

<GENERAL RESPONSES>
Be warm and helpful when the request is casual or unclear. Mention your abilities naturally.

\u2705 Hi there! I can help with summaries, task tracking, or finding specific messages.
\u2705 Interesting! I specialize in conversation analysis and action items. Want help with that?
`;
}

// src/agent/manager.ts
var ManagerPrompt = class {
  constructor(context, logger2) {
    this.context = context;
    this.logger = logger2;
  }
  prompt;
  isInitialized = false;
  async createManagerPrompt() {
    const managerModelConfig = getModelConfig("manager");
    const prompt = new teams_ai.ChatPrompt({
      instructions: generateManagerPrompt(CAPABILITY_DEFINITIONS),
      model: new teams_openai.OpenAIChatModel({
        model: managerModelConfig.model,
        apiKey: managerModelConfig.apiKey,
        endpoint: managerModelConfig.endpoint,
        apiVersion: managerModelConfig.apiVersion
      }),
      messages: await this.context.memory.values()
    }).function(
      "calculate_time_range",
      "Parse natural language time expressions and calculate exact start/end times for time-based queries",
      {
        type: "object",
        properties: {
          time_phrase: {
            type: "string",
            description: 'Natural language time expression extracted from the user request (e.g., "yesterday", "last week", "2 days ago", "past 3 hours")'
          }
        },
        required: ["time_phrase"]
      },
      async (time_phrase) => {
        this.logger.debug(`\u{1F552} FUNCTION CALL: calculate_time_range - parsing "${time_phrase}"`);
        const timeRange = extractTimeRange(time_phrase);
        this.context.startTime = timeRange ? timeRange?.from.toISOString() : this.context.endTime;
        this.context.endTime = timeRange ? timeRange?.to.toISOString() : this.context.endTime;
        this.logger.debug(this.context.startTime);
        this.logger.debug(this.context.endTime);
      }
    ).function(
      "clear_conversation_history",
      "Clear conversation history in the database for the current conversation",
      async () => {
        await this.context.memory.clear();
        this.logger.debug("The conversation history has been cleared!");
      }
    );
    return prompt;
  }
  addCapabilities() {
    for (const capability of CAPABILITY_DEFINITIONS) {
      this.prompt.function(
        `delegate_to_${capability.name}`,
        `Delegate to ${capability.name} capability`,
        async () => {
          return capability.handler(this.context, this.logger.child(capability.name));
        }
      );
    }
  }
  async initialize() {
    if (!this.isInitialized) {
      this.prompt = await this.createManagerPrompt();
      this.addCapabilities();
      this.isInitialized = true;
    }
  }
  async processRequest() {
    try {
      await this.initialize();
      const response = await this.prompt.send(this.context.text);
      return {
        response: response.content || "No response generated"
      };
    } catch (error) {
      this.logger.error("\u274C Error in Manager:", error);
      return {
        response: `Sorry, I encountered an error processing your request: ${error instanceof Error ? error.message : "Unknown error"}`
      };
    }
  }
};
var dbInstance = null;
var configInstance = null;
var loggerInstance = null;
function initializeMeetingRoutes(db, config, logger2) {
  dbInstance = db;
  configInstance = config;
  loggerInstance = logger2;
}
var router = express.Router();
function verifySharedSecret(req, res, next) {
  if (!configInstance) {
    res.status(500).json({ error: "Server not initialized" });
    return;
  }
  const providedSecret = req.headers["x-shared-secret"];
  const expectedSecret = configInstance.meetingMediaBotSharedSecret;
  if (!providedSecret || !expectedSecret || providedSecret !== expectedSecret) {
    loggerInstance?.warn(`Unauthorized request to ${req.path}`);
    res.status(401).json({ error: "Unauthorized" });
    return;
  }
  next();
}
router.post("/meeting-transcripts/chunk", verifySharedSecret, async (req, res) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }
  const {
    callId,
    text,
    speakerId,
    timestamp,
    source
  } = req.body;
  loggerInstance?.debug(`Received transcript chunk for call ${callId}: "${text?.substring(0, 50)}..."`);
  if (!callId || !text) {
    res.status(400).json({ error: "callId and text are required" });
    return;
  }
  try {
    let dbSource = "speech";
    if (source === "azure_speech" || source === "speech") {
      dbSource = "speech";
    } else if (source === "graph_transcript" || source === "graphTranscript") {
      dbSource = "graphTranscript";
    }
    const existingMeeting = await dbInstance.getMeeting(callId);
    if (!existingMeeting) {
      await dbInstance.upsertMeeting({
        meetingId: callId,
        conversationId: "api-upload",
        // Placeholder for API-uploaded transcripts
        joinUrl: `api://meeting/${callId}`,
        // Placeholder URL for API-uploaded meetings
        title: `Meeting ${callId}`,
        organizerAadId: speakerId || "unknown"
      });
      loggerInstance?.debug(`Created meeting record for ${callId}`);
    }
    const chunkId = await dbInstance.appendTranscriptChunk({
      meetingId: callId,
      speaker: speakerId || "Unknown",
      text,
      timestampUtc: timestamp || (/* @__PURE__ */ new Date()).toISOString(),
      confidence: 1,
      // Azure Speech provides high confidence final results
      source: dbSource
    });
    loggerInstance?.debug(`Stored transcript chunk ${chunkId} for meeting ${callId}`);
    res.json({
      success: true,
      chunkId,
      meetingId: callId
    });
  } catch (error) {
    loggerInstance?.error(`Error storing transcript chunk: ${error}`);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error"
    });
  }
});
router.post("/meeting-capture/status", verifySharedSecret, async (req, res) => {
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
    let meetingStatus;
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
      endedAt: meetingStatus === "ended" ? (/* @__PURE__ */ new Date()).toISOString() : void 0
    });
    loggerInstance?.debug(`Updated meeting ${callId} status to ${meetingStatus}`);
    res.json({
      success: true,
      meetingId: callId,
      mappedStatus: meetingStatus
    });
  } catch (error2) {
    loggerInstance?.error(`Error updating meeting status: ${error2}`);
    res.status(500).json({
      success: false,
      error: error2 instanceof Error ? error2.message : "Unknown error"
    });
  }
});
router.get("/meeting-transcripts/:meetingId", async (req, res) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }
  const meetingId = req.params.meetingId;
  loggerInstance?.debug(`Retrieving transcript for meeting ${meetingId}`);
  try {
    const result = await dbInstance.getTranscriptByMeetingId(meetingId);
    if (!result) {
      res.status(404).json({ error: "Meeting not found" });
      return;
    }
    const fullText = result.chunks.map((chunk) => `[${chunk.speaker}]: ${chunk.text}`).join("\n");
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
        source: chunk.source
      })),
      participants: result.participants
    });
  } catch (error) {
    loggerInstance?.error(`Error retrieving transcript: ${error}`);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : "Unknown error"
    });
  }
});
router.get("/meetings/:meetingId", async (req, res) => {
  if (!dbInstance) {
    res.status(500).json({ error: "Database not initialized" });
    return;
  }
  const meetingId = req.params.meetingId;
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
      error: error instanceof Error ? error.message : "Unknown error"
    });
  }
});
router.get("/health", (_req, res) => {
  res.json({
    status: "ok",
    service: "project-missa",
    meetingApiEnabled: true,
    timestamp: (/* @__PURE__ */ new Date()).toISOString()
  });
});
var meetingApi_default = router;

// src/storage/storageFactory.ts
init_config();
var MssqlKVStore = class {
  constructor(logger2, config) {
    this.logger = logger2;
    this.config = config;
  }
  pool = null;
  isInitialized = false;
  async initialize() {
    if (this.isInitialized) return;
    try {
      let sqlConfig;
      if (this.config.connectionString) {
        this.pool = new mssql__namespace.ConnectionPool(this.config.connectionString);
      } else {
        sqlConfig = {
          server: this.config.server,
          database: this.config.database,
          user: this.config.username,
          password: this.config.password,
          options: {
            encrypt: true,
            trustServerCertificate: false
          }
        };
        this.pool = new mssql__namespace.ConnectionPool(sqlConfig);
      }
      await this.pool.connect();
      await this.initializeDatabase();
      this.isInitialized = true;
      this.logger.debug("\u2705 Connected to MSSQL database");
    } catch (error) {
      this.logger.error("\u274C Error connecting to MSSQL database:", error);
      throw error;
    }
  }
  async initializeDatabase() {
    if (!this.pool) throw new Error("Database not connected");
    try {
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='conversations' AND xtype='U')
        BEGIN
          CREATE TABLE conversations (
            id INT IDENTITY(1,1) PRIMARY KEY,
            conversation_id NVARCHAR(255) NOT NULL,
            role NVARCHAR(50) NOT NULL,
            name NVARCHAR(255) NOT NULL,
            content NVARCHAR(MAX) NOT NULL,
            activity_id NVARCHAR(255) NOT NULL,
            timestamp NVARCHAR(50) NOT NULL,
            blob NVARCHAR(MAX) NOT NULL
          )
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_conversation_id' AND object_id = OBJECT_ID('conversations'))
        BEGIN
          CREATE INDEX idx_conversation_id ON conversations(conversation_id)
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='feedback' AND xtype='U')
        BEGIN
          CREATE TABLE feedback (
            id INT IDENTITY(1,1) PRIMARY KEY,
            reply_to_id NVARCHAR(255) NOT NULL,
            reaction NVARCHAR(50) NOT NULL CHECK (reaction IN ('like','dislike')),
            feedback NVARCHAR(MAX),
            created_at DATETIME NOT NULL DEFAULT GETDATE()
          )
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='meetings' AND xtype='U')
        BEGIN
          CREATE TABLE meetings (
            meeting_id NVARCHAR(255) PRIMARY KEY,
            join_url NVARCHAR(2048) NOT NULL,
            organizer_aad_id NVARCHAR(255) NOT NULL,
            organizer_display_name NVARCHAR(255),
            organizer_email NVARCHAR(255),
            started_at NVARCHAR(50),
            ended_at NVARCHAR(50),
            status NVARCHAR(50) NOT NULL DEFAULT 'joining',
            title NVARCHAR(500),
            conversation_id NVARCHAR(255),
            requested_by_aad_id NVARCHAR(255),
            created_at DATETIME NOT NULL DEFAULT GETDATE(),
            updated_at DATETIME NOT NULL DEFAULT GETDATE()
          )
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_meetings_status' AND object_id = OBJECT_ID('meetings'))
        BEGIN
          CREATE INDEX idx_meetings_status ON meetings(status)
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_meetings_conversation' AND object_id = OBJECT_ID('meetings'))
        BEGIN
          CREATE INDEX idx_meetings_conversation ON meetings(conversation_id)
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='meeting_participants' AND xtype='U')
        BEGIN
          CREATE TABLE meeting_participants (
            id INT IDENTITY(1,1) PRIMARY KEY,
            meeting_id NVARCHAR(255) NOT NULL,
            participant_aad_id NVARCHAR(255) NOT NULL,
            display_name NVARCHAR(255) NOT NULL,
            email NVARCHAR(255),
            joined_at NVARCHAR(50),
            left_at NVARCHAR(50),
            CONSTRAINT FK_participants_meeting FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id),
            CONSTRAINT UQ_participant UNIQUE (meeting_id, participant_aad_id)
          )
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_participants_meeting' AND object_id = OBJECT_ID('meeting_participants'))
        BEGIN
          CREATE INDEX idx_participants_meeting ON meeting_participants(meeting_id)
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='transcript_chunks' AND xtype='U')
        BEGIN
          CREATE TABLE transcript_chunks (
            id INT IDENTITY(1,1) PRIMARY KEY,
            meeting_id NVARCHAR(255) NOT NULL,
            timestamp_utc NVARCHAR(50) NOT NULL,
            speaker NVARCHAR(255) NOT NULL,
            speaker_aad_id NVARCHAR(255),
            text NVARCHAR(MAX) NOT NULL,
            confidence FLOAT NOT NULL DEFAULT 0.0,
            source NVARCHAR(50) NOT NULL CHECK (source IN ('speech', 'graphTranscript')),
            sequence_number INT,
            created_at DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT FK_chunks_meeting FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
          )
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_chunks_meeting' AND object_id = OBJECT_ID('transcript_chunks'))
        BEGIN
          CREATE INDEX idx_chunks_meeting ON transcript_chunks(meeting_id)
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_chunks_timestamp' AND object_id = OBJECT_ID('transcript_chunks'))
        BEGIN
          CREATE INDEX idx_chunks_timestamp ON transcript_chunks(timestamp_utc)
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='meeting_summaries' AND xtype='U')
        BEGIN
          CREATE TABLE meeting_summaries (
            id INT IDENTITY(1,1) PRIMARY KEY,
            meeting_id NVARCHAR(255) NOT NULL,
            title NVARCHAR(500) NOT NULL,
            summary_json NVARCHAR(MAX) NOT NULL,
            short_summary NVARCHAR(MAX) NOT NULL,
            generated_at DATETIME NOT NULL DEFAULT GETDATE(),
            generated_by_aad_id NVARCHAR(255),
            CONSTRAINT FK_summaries_meeting FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
          )
        END
      `);
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_summaries_meeting' AND object_id = OBJECT_ID('meeting_summaries'))
        BEGIN
          CREATE INDEX idx_summaries_meeting ON meeting_summaries(meeting_id)
        END
      `);
      this.logger.debug("\u2705 Database tables initialized");
    } catch (error) {
      this.logger.error("\u274C Error initializing database tables:", error);
      throw error;
    }
  }
  async clearAll() {
    if (!this.pool) throw new Error("Database not connected");
    try {
      await this.pool.request().query("DELETE FROM conversations");
      this.logger.debug("\u{1F9F9} Cleared all conversations from MSSQL store.");
    } catch (error) {
      this.logger.error("\u274C Error clearing all conversations:", error);
      throw error;
    }
  }
  async get(conversationId) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("conversationId", mssql__namespace.NVarChar, conversationId).query(
        "SELECT blob FROM conversations WHERE conversation_id = @conversationId ORDER BY timestamp ASC"
      );
      return result.recordset.map((row) => JSON.parse(row.blob));
    } catch (error) {
      this.logger.error("\u274C Error getting messages:", error);
      return [];
    }
  }
  async getMessagesByTimeRange(conversationId, startTime, endTime) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("conversationId", mssql__namespace.NVarChar, conversationId).input("startTime", mssql__namespace.NVarChar, startTime).input("endTime", mssql__namespace.NVarChar, endTime).query(`
          SELECT blob FROM conversations 
          WHERE conversation_id = @conversationId 
            AND timestamp >= @startTime 
            AND timestamp <= @endTime 
          ORDER BY timestamp ASC
        `);
      return result.recordset.map((row) => JSON.parse(row.blob));
    } catch (error) {
      this.logger.error("\u274C Error getting messages by time range:", error);
      return [];
    }
  }
  async getRecentMessages(conversationId, limit = 10) {
    const messages = await this.get(conversationId);
    return messages.slice(-limit);
  }
  async clearConversation(conversationId) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      await this.pool.request().input("conversationId", mssql__namespace.NVarChar, conversationId).query("DELETE FROM conversations WHERE conversation_id = @conversationId");
    } catch (error) {
      this.logger.error("\u274C Error clearing conversation:", error);
      throw error;
    }
  }
  async addMessages(messages) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const transaction = new mssql__namespace.Transaction(this.pool);
      await transaction.begin();
      try {
        for (const message of messages) {
          await transaction.request().input("conversationId", mssql__namespace.NVarChar, message.conversation_id).input("role", mssql__namespace.NVarChar, message.role).input("name", mssql__namespace.NVarChar, message.name).input("content", mssql__namespace.NVarChar, message.content).input("activityId", mssql__namespace.NVarChar, message.activity_id).input("timestamp", mssql__namespace.NVarChar, message.timestamp).input("blob", mssql__namespace.NVarChar, JSON.stringify(message)).query(`
              INSERT INTO conversations (conversation_id, role, name, content, activity_id, timestamp, blob)
              VALUES (@conversationId, @role, @name, @content, @activityId, @timestamp, @blob)
            `);
        }
        await transaction.commit();
      } catch (error) {
        await transaction.rollback();
        throw error;
      }
    } catch (error) {
      this.logger.error("\u274C Error adding messages:", error);
      throw error;
    }
  }
  async countMessages(conversationId) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("conversationId", mssql__namespace.NVarChar, conversationId).query(
        "SELECT COUNT(*) as count FROM conversations WHERE conversation_id = @conversationId"
      );
      return result.recordset[0].count;
    } catch (error) {
      this.logger.error("\u274C Error counting messages:", error);
      return 0;
    }
  }
  async clearAllMessages() {
    await this.clearAll();
  }
  async getFilteredMessages(conversationId, keywords, startTime, endTime, participants, maxResults = 5) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const request = this.pool.request();
      let whereClause = "conversation_id = @conversationId AND timestamp >= @startTime AND timestamp <= @endTime";
      request.input("conversationId", mssql__namespace.NVarChar, conversationId);
      request.input("startTime", mssql__namespace.NVarChar, startTime);
      request.input("endTime", mssql__namespace.NVarChar, endTime);
      request.input("maxResults", mssql__namespace.Int, maxResults);
      if (keywords.length > 0) {
        const keywordConditions = keywords.map((_, index) => {
          request.input(`keyword${index}`, mssql__namespace.NVarChar, `%${keywords[index].toLowerCase()}%`);
          return `content LIKE @keyword${index}`;
        }).join(" OR ");
        whereClause += ` AND (${keywordConditions})`;
      }
      if (participants && participants.length > 0) {
        const participantConditions = participants.map((_, index) => {
          request.input(
            `participant${index}`,
            mssql__namespace.NVarChar,
            `%${participants[index].toLowerCase()}%`
          );
          return `name LIKE @participant${index}`;
        }).join(" OR ");
        whereClause += ` AND (${participantConditions})`;
      }
      const query = `
        SELECT TOP (@maxResults) blob FROM conversations
        WHERE ${whereClause}
        ORDER BY timestamp DESC
      `;
      const result = await request.query(query);
      return result.recordset.map((row) => JSON.parse(row.blob));
    } catch (error) {
      this.logger.error("\u274C Error getting filtered messages:", error);
      return [];
    }
  }
  async recordFeedback(replyToId, reaction, feedbackJson) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      await this.pool.request().input("replyToId", mssql__namespace.NVarChar, replyToId).input("reaction", mssql__namespace.NVarChar, reaction).input("feedback", mssql__namespace.NVarChar, feedbackJson ? JSON.stringify(feedbackJson) : null).query(`
          INSERT INTO feedback (reply_to_id, reaction, feedback)
          VALUES (@replyToId, @reaction, @feedback)
        `);
      return true;
    } catch (error) {
      this.logger.error("\u274C Error recording feedback:", error);
      return false;
    }
  }
  // ============================================================================
  // MEETING STORAGE METHODS
  // ============================================================================
  async upsertMeeting(input) {
    if (!this.pool) throw new Error("Database not connected");
    const now = (/* @__PURE__ */ new Date()).toISOString();
    try {
      await this.pool.request().input("meetingId", mssql__namespace.NVarChar, input.meetingId).input("joinUrl", mssql__namespace.NVarChar, input.joinUrl).input("organizerAadId", mssql__namespace.NVarChar, input.organizerAadId).input("organizerDisplayName", mssql__namespace.NVarChar, input.organizerDisplayName || null).input("organizerEmail", mssql__namespace.NVarChar, input.organizerEmail || null).input("startedAt", mssql__namespace.NVarChar, now).input("title", mssql__namespace.NVarChar, input.title || null).input("conversationId", mssql__namespace.NVarChar, input.conversationId || null).input("requestedByAadId", mssql__namespace.NVarChar, input.requestedByAadId || null).input("now", mssql__namespace.DateTime, /* @__PURE__ */ new Date()).query(`
          MERGE meetings AS target
          USING (SELECT @meetingId AS meeting_id) AS source
          ON (target.meeting_id = source.meeting_id)
          WHEN MATCHED THEN
            UPDATE SET
              join_url = @joinUrl,
              organizer_aad_id = @organizerAadId,
              organizer_display_name = @organizerDisplayName,
              organizer_email = @organizerEmail,
              title = @title,
              conversation_id = @conversationId,
              requested_by_aad_id = @requestedByAadId,
              updated_at = @now
          WHEN NOT MATCHED THEN
            INSERT (meeting_id, join_url, organizer_aad_id, organizer_display_name, organizer_email, started_at, status, title, conversation_id, requested_by_aad_id, created_at, updated_at)
            VALUES (@meetingId, @joinUrl, @organizerAadId, @organizerDisplayName, @organizerEmail, @startedAt, 'joining', @title, @conversationId, @requestedByAadId, @now, @now);
        `);
      const meeting = await this.getMeeting(input.meetingId);
      if (!meeting) {
        throw new Error(`Failed to retrieve meeting after upsert: ${input.meetingId}`);
      }
      this.logger.debug(`\u2705 Upserted meeting: ${input.meetingId}`);
      return meeting;
    } catch (error) {
      this.logger.error("\u274C upsertMeeting error:", error);
      throw error;
    }
  }
  async updateMeetingStatus(input) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const request = this.pool.request().input("meetingId", mssql__namespace.NVarChar, input.meetingId).input("status", mssql__namespace.NVarChar, input.status).input("now", mssql__namespace.DateTime, /* @__PURE__ */ new Date());
      if (input.endedAt) {
        request.input("endedAt", mssql__namespace.NVarChar, input.endedAt);
        await request.query(`
          UPDATE meetings 
          SET status = @status, ended_at = @endedAt, updated_at = @now
          WHERE meeting_id = @meetingId
        `);
      } else {
        await request.query(`
          UPDATE meetings 
          SET status = @status, updated_at = @now
          WHERE meeting_id = @meetingId
        `);
      }
      this.logger.debug(`\u2705 Updated meeting ${input.meetingId} status to: ${input.status}`);
      return true;
    } catch (error) {
      this.logger.error("\u274C updateMeetingStatus error:", error);
      return false;
    }
  }
  async getMeeting(meetingId) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("meetingId", mssql__namespace.NVarChar, meetingId).query(`SELECT * FROM meetings WHERE meeting_id = @meetingId`);
      if (result.recordset.length === 0) return null;
      const row = result.recordset[0];
      return {
        meetingId: row.meeting_id,
        joinUrl: row.join_url,
        organizerAadId: row.organizer_aad_id,
        organizerDisplayName: row.organizer_display_name || void 0,
        organizerEmail: row.organizer_email || void 0,
        startedAt: row.started_at,
        endedAt: row.ended_at || void 0,
        status: row.status,
        title: row.title || void 0,
        conversationId: row.conversation_id || void 0,
        requestedByAadId: row.requested_by_aad_id || void 0,
        createdAt: row.created_at?.toISOString() || (/* @__PURE__ */ new Date()).toISOString(),
        updatedAt: row.updated_at?.toISOString() || (/* @__PURE__ */ new Date()).toISOString()
      };
    } catch (error) {
      this.logger.error("\u274C getMeeting error:", error);
      return null;
    }
  }
  async upsertParticipants(meetingId, participants) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      for (const participant of participants) {
        await this.pool.request().input("meetingId", mssql__namespace.NVarChar, meetingId).input("participantAadId", mssql__namespace.NVarChar, participant.participantAadId).input("displayName", mssql__namespace.NVarChar, participant.displayName).input("email", mssql__namespace.NVarChar, participant.email || null).input("joinedAt", mssql__namespace.NVarChar, participant.joinedAt || null).input("leftAt", mssql__namespace.NVarChar, participant.leftAt || null).query(`
            MERGE meeting_participants AS target
            USING (SELECT @meetingId AS meeting_id, @participantAadId AS participant_aad_id) AS source
            ON (target.meeting_id = source.meeting_id AND target.participant_aad_id = source.participant_aad_id)
            WHEN MATCHED THEN
              UPDATE SET
                display_name = @displayName,
                email = @email,
                joined_at = COALESCE(@joinedAt, target.joined_at),
                left_at = @leftAt
            WHEN NOT MATCHED THEN
              INSERT (meeting_id, participant_aad_id, display_name, email, joined_at, left_at)
              VALUES (@meetingId, @participantAadId, @displayName, @email, @joinedAt, @leftAt);
          `);
      }
      this.logger.debug(`\u2705 Upserted ${participants.length} participants for meeting: ${meetingId}`);
    } catch (error) {
      this.logger.error("\u274C upsertParticipants error:", error);
      throw error;
    }
  }
  async getParticipants(meetingId) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("meetingId", mssql__namespace.NVarChar, meetingId).query(`SELECT * FROM meeting_participants WHERE meeting_id = @meetingId`);
      return result.recordset.map((row) => ({
        id: row.id,
        meetingId: row.meeting_id,
        participantAadId: row.participant_aad_id,
        displayName: row.display_name,
        email: row.email || void 0,
        joinedAt: row.joined_at || void 0,
        leftAt: row.left_at || void 0
      }));
    } catch (error) {
      this.logger.error("\u274C getParticipants error:", error);
      return [];
    }
  }
  async appendTranscriptChunk(input) {
    if (!this.pool) throw new Error("Database not connected");
    const now = (/* @__PURE__ */ new Date()).toISOString();
    try {
      const seqResult = await this.pool.request().input("meetingId", mssql__namespace.NVarChar, input.meetingId).query(`SELECT ISNULL(MAX(sequence_number), -1) as max_seq FROM transcript_chunks WHERE meeting_id = @meetingId`);
      const sequenceNumber = (seqResult.recordset[0].max_seq ?? -1) + 1;
      const insertResult = await this.pool.request().input("meetingId", mssql__namespace.NVarChar, input.meetingId).input("timestampUtc", mssql__namespace.NVarChar, input.timestampUtc).input("speaker", mssql__namespace.NVarChar, input.speaker).input("speakerAadId", mssql__namespace.NVarChar, input.speakerAadId || null).input("text", mssql__namespace.NVarChar, input.text).input("confidence", mssql__namespace.Float, input.confidence).input("source", mssql__namespace.NVarChar, input.source).input("sequenceNumber", mssql__namespace.Int, sequenceNumber).query(`
          INSERT INTO transcript_chunks (meeting_id, timestamp_utc, speaker, speaker_aad_id, text, confidence, source, sequence_number)
          OUTPUT INSERTED.id
          VALUES (@meetingId, @timestampUtc, @speaker, @speakerAadId, @text, @confidence, @source, @sequenceNumber)
        `);
      const insertedId = insertResult.recordset[0].id;
      this.logger.debug(`\u2705 Appended transcript chunk for meeting: ${input.meetingId}, seq: ${sequenceNumber}`);
      return {
        id: insertedId,
        meetingId: input.meetingId,
        timestampUtc: input.timestampUtc,
        speaker: input.speaker,
        speakerAadId: input.speakerAadId,
        text: input.text,
        confidence: input.confidence,
        source: input.source,
        sequenceNumber,
        createdAt: now
      };
    } catch (error) {
      this.logger.error("\u274C appendTranscriptChunk error:", error);
      throw error;
    }
  }
  async getTranscriptChunks(meetingId) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("meetingId", mssql__namespace.NVarChar, meetingId).query(`SELECT * FROM transcript_chunks WHERE meeting_id = @meetingId ORDER BY sequence_number ASC, timestamp_utc ASC`);
      return result.recordset.map((row) => ({
        id: row.id,
        meetingId: row.meeting_id,
        timestampUtc: row.timestamp_utc,
        speaker: row.speaker,
        speakerAadId: row.speaker_aad_id || void 0,
        text: row.text,
        confidence: row.confidence,
        source: row.source,
        sequenceNumber: row.sequence_number || void 0,
        createdAt: row.created_at?.toISOString() || (/* @__PURE__ */ new Date()).toISOString()
      }));
    } catch (error) {
      this.logger.error("\u274C getTranscriptChunks error:", error);
      return [];
    }
  }
  async getTranscriptByMeetingId(meetingId) {
    try {
      const meeting = await this.getMeeting(meetingId);
      if (!meeting) {
        this.logger.warn(`Meeting not found: ${meetingId}`);
        return null;
      }
      const participants = await this.getParticipants(meetingId);
      const chunks = await this.getTranscriptChunks(meetingId);
      return {
        meetingId: meeting.meetingId,
        title: meeting.title,
        startedAt: meeting.startedAt,
        endedAt: meeting.endedAt,
        participants,
        chunks,
        totalChunks: chunks.length
      };
    } catch (error) {
      this.logger.error("\u274C getTranscriptByMeetingId error:", error);
      return null;
    }
  }
  async getMeetingsByStatus(status) {
    if (!this.pool) throw new Error("Database not connected");
    try {
      const result = await this.pool.request().input("status", mssql__namespace.NVarChar, status).query(`SELECT * FROM meetings WHERE status = @status`);
      return result.recordset.map((row) => ({
        meetingId: row.meeting_id,
        joinUrl: row.join_url,
        organizerAadId: row.organizer_aad_id,
        organizerDisplayName: row.organizer_display_name || void 0,
        organizerEmail: row.organizer_email || void 0,
        startedAt: row.started_at,
        endedAt: row.ended_at || void 0,
        status: row.status,
        title: row.title || void 0,
        conversationId: row.conversation_id || void 0,
        requestedByAadId: row.requested_by_aad_id || void 0,
        createdAt: row.created_at?.toISOString() || (/* @__PURE__ */ new Date()).toISOString(),
        updatedAt: row.updated_at?.toISOString() || (/* @__PURE__ */ new Date()).toISOString()
      }));
    } catch (error) {
      this.logger.error("\u274C getMeetingsByStatus error:", error);
      return [];
    }
  }
  async close() {
    if (this.pool) {
      await this.pool.close();
      this.pool = null;
      this.isInitialized = false;
      this.logger.debug("\u{1F50C} Closed MSSQL database connection");
    }
  }
};
var SqliteKVStore = class {
  constructor(logger2, dbPath) {
    this.logger = logger2;
    const resolvedDbPath = process.env.CONVERSATIONS_DB_PATH ? path__default.default.resolve(process.env.CONVERSATIONS_DB_PATH) : dbPath ? dbPath : path__default.default.resolve(__dirname, "../../src/storage/conversations.db");
    this.db = new Database__default.default(resolvedDbPath);
    this.initializeDatabase();
  }
  db;
  async initialize() {
    return Promise.resolve();
  }
  initializeDatabase() {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS conversations (
        conversation_id TEXT NOT NULL,
        role TEXT NOT NULL,
        name TEXT NOT NULL,
        content TEXT NOT NULL,
        activity_id TEXT NOT NULL,
        timestamp TEXT NOT NULL,
        blob TEXT NOT NULL
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_conversation_id ON conversations(conversation_id);
    `);
    this.db.exec(`
    CREATE TABLE IF NOT EXISTS feedback (
    reply_to_id  TEXT    NOT NULL,                -- the Teams message ID you replied to
    reaction     TEXT    NOT NULL CHECK (reaction IN ('like','dislike')),
    feedback     TEXT,                           -- JSON or plain text
    created_at   TEXT    NOT NULL DEFAULT (CURRENT_TIMESTAMP)
  );
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meetings (
        meeting_id TEXT PRIMARY KEY,
        join_url TEXT NOT NULL,
        organizer_aad_id TEXT NOT NULL,
        organizer_display_name TEXT,
        organizer_email TEXT,
        started_at TEXT,
        ended_at TEXT,
        status TEXT NOT NULL DEFAULT 'joining',
        title TEXT,
        conversation_id TEXT,
        requested_by_aad_id TEXT,
        created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
        updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_meetings_status ON meetings(status);
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_meetings_conversation ON meetings(conversation_id);
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meeting_participants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        meeting_id TEXT NOT NULL,
        participant_aad_id TEXT NOT NULL,
        display_name TEXT NOT NULL,
        email TEXT,
        joined_at TEXT,
        left_at TEXT,
        FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id),
        UNIQUE(meeting_id, participant_aad_id)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_participants_meeting ON meeting_participants(meeting_id);
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS transcript_chunks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        meeting_id TEXT NOT NULL,
        timestamp_utc TEXT NOT NULL,
        speaker TEXT NOT NULL,
        speaker_aad_id TEXT,
        text TEXT NOT NULL,
        confidence REAL NOT NULL DEFAULT 0.0,
        source TEXT NOT NULL CHECK (source IN ('speech', 'graphTranscript')),
        sequence_number INTEGER,
        created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
        FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_chunks_meeting ON transcript_chunks(meeting_id);
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_chunks_timestamp ON transcript_chunks(timestamp_utc);
    `);
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meeting_summaries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        meeting_id TEXT NOT NULL,
        title TEXT NOT NULL,
        summary_json TEXT NOT NULL,
        short_summary TEXT NOT NULL,
        generated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
        generated_by_aad_id TEXT,
        FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_summaries_meeting ON meeting_summaries(meeting_id);
    `);
  }
  clearAll() {
    this.db.exec("DELETE FROM conversations; VACUUM;");
    this.logger.debug("\u{1F9F9} Cleared all conversations from SQLite store.");
  }
  get(conversationId) {
    const stmt = this.db.prepare(
      "SELECT blob FROM conversations WHERE conversation_id = ? ORDER BY timestamp ASC"
    );
    return stmt.all(conversationId).map((row) => JSON.parse(row.blob));
  }
  getMessagesByTimeRange(conversationId, startTime, endTime) {
    const stmt = this.db.prepare(
      "SELECT blob FROM conversations WHERE conversation_id = ? AND timestamp >= ? AND timestamp <= ? ORDER BY timestamp ASC"
    );
    return stmt.all([conversationId, startTime, endTime]).map((row) => JSON.parse(row.blob));
  }
  getRecentMessages(conversationId, limit = 10) {
    const messages = this.get(conversationId);
    return messages.slice(-limit);
  }
  clearConversation(conversationId) {
    const stmt = this.db.prepare("DELETE FROM conversations WHERE conversation_id = ?");
    stmt.run(conversationId);
  }
  addMessages(messages) {
    const stmt = this.db.prepare(
      "INSERT INTO conversations (conversation_id, role, name, content, activity_id, timestamp, blob) VALUES (?, ?, ?, ?, ?, ?, ?)"
    );
    for (const message of messages) {
      stmt.run(
        message.conversation_id,
        message.role,
        message.name,
        message.content,
        message.activity_id,
        message.timestamp,
        JSON.stringify(message)
      );
    }
  }
  countMessages(conversationId) {
    const stmt = this.db.prepare(
      "SELECT COUNT(*) as count FROM conversations WHERE conversation_id = ?"
    );
    const result = stmt.get(conversationId);
    return result.count;
  }
  // Clear all messages for debugging (optional utility method)
  clearAllMessages() {
    try {
      const stmt = this.db.prepare("DELETE FROM conversations");
      const result = stmt.run();
      this.logger.debug(
        `\u{1F9F9} Cleared all conversations from database. Deleted ${result.changes} records.`
      );
    } catch (error) {
      this.logger.error("\u274C Error clearing all conversations:", error);
    }
  }
  getFilteredMessages(conversationId, keywords, startTime, endTime, participants, maxResults) {
    const keywordClauses = keywords.map(() => `content LIKE ?`).join(" OR ");
    const participantClauses = participants?.map(() => `name LIKE ?`).join(" OR ");
    const whereClauses = [
      `conversation_id = ?`,
      `timestamp >= ?`,
      `timestamp <= ?`,
      `(${keywordClauses})`
    ];
    const values = [
      conversationId,
      startTime,
      endTime,
      ...keywords.map((k) => `%${k.toLowerCase()}%`)
    ];
    if (participants && participants.length > 0) {
      whereClauses.push(`(${participantClauses})`);
      values.push(...participants.map((p) => `%${p.toLowerCase()}%`));
    }
    const limit = maxResults && typeof maxResults === "number" ? maxResults : 5;
    values.push(limit);
    const query = `
  SELECT blob FROM conversations
  WHERE ${whereClauses.join(" AND ")}
  ORDER BY timestamp DESC
  LIMIT ?
`;
    const stmt = this.db.prepare(query);
    const rows = stmt.all(...values);
    return rows.map((row) => JSON.parse(row.blob));
  }
  // ===== FEEDBACK MANAGEMENT =====
  // Initialize feedback record for a message with optional delegated capability
  // Insert one row per submission
  recordFeedback(replyToId, reaction, feedbackJson) {
    try {
      const stmt = this.db.prepare(`
      INSERT INTO feedback (reply_to_id, reaction, feedback)
      VALUES (?, ?, ?)
    `);
      const result = stmt.run(
        replyToId,
        reaction,
        feedbackJson ? JSON.stringify(feedbackJson) : null
      );
      return result.changes > 0;
    } catch (err) {
      this.logger.error(`\u274C recordFeedback error:`, err);
      return false;
    }
  }
  // ============================================================================
  // MEETING STORAGE METHODS
  // ============================================================================
  async upsertMeeting(input) {
    const now = (/* @__PURE__ */ new Date()).toISOString();
    try {
      const stmt = this.db.prepare(`
        INSERT INTO meetings (
          meeting_id, join_url, organizer_aad_id, organizer_display_name, organizer_email,
          started_at, status, title, conversation_id, requested_by_aad_id, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, 'joining', ?, ?, ?, ?, ?)
        ON CONFLICT(meeting_id) DO UPDATE SET
          join_url = excluded.join_url,
          organizer_aad_id = excluded.organizer_aad_id,
          organizer_display_name = excluded.organizer_display_name,
          organizer_email = excluded.organizer_email,
          title = excluded.title,
          conversation_id = excluded.conversation_id,
          requested_by_aad_id = excluded.requested_by_aad_id,
          updated_at = excluded.updated_at
      `);
      stmt.run(
        input.meetingId,
        input.joinUrl,
        input.organizerAadId,
        input.organizerDisplayName || null,
        input.organizerEmail || null,
        now,
        input.title || null,
        input.conversationId || null,
        input.requestedByAadId || null,
        now,
        now
      );
      const meeting = await this.getMeeting(input.meetingId);
      if (!meeting) {
        throw new Error(`Failed to retrieve meeting after upsert: ${input.meetingId}`);
      }
      this.logger.debug(`\u2705 Upserted meeting: ${input.meetingId}`);
      return meeting;
    } catch (err) {
      this.logger.error(`\u274C upsertMeeting error:`, err);
      throw err;
    }
  }
  async updateMeetingStatus(input) {
    const now = (/* @__PURE__ */ new Date()).toISOString();
    try {
      let stmt;
      if (input.endedAt) {
        stmt = this.db.prepare(`
          UPDATE meetings SET status = ?, ended_at = ?, updated_at = ?
          WHERE meeting_id = ?
        `);
        stmt.run(input.status, input.endedAt, now, input.meetingId);
      } else {
        stmt = this.db.prepare(`
          UPDATE meetings SET status = ?, updated_at = ?
          WHERE meeting_id = ?
        `);
        stmt.run(input.status, now, input.meetingId);
      }
      this.logger.debug(`\u2705 Updated meeting ${input.meetingId} status to: ${input.status}`);
      return true;
    } catch (err) {
      this.logger.error(`\u274C updateMeetingStatus error:`, err);
      return false;
    }
  }
  async getMeeting(meetingId) {
    try {
      const stmt = this.db.prepare(`SELECT * FROM meetings WHERE meeting_id = ?`);
      const row = stmt.get(meetingId);
      if (!row) return null;
      return {
        meetingId: row.meeting_id,
        joinUrl: row.join_url,
        organizerAadId: row.organizer_aad_id,
        organizerDisplayName: row.organizer_display_name || void 0,
        organizerEmail: row.organizer_email || void 0,
        startedAt: row.started_at,
        endedAt: row.ended_at || void 0,
        status: row.status,
        title: row.title || void 0,
        conversationId: row.conversation_id || void 0,
        requestedByAadId: row.requested_by_aad_id || void 0,
        createdAt: row.created_at,
        updatedAt: row.updated_at
      };
    } catch (err) {
      this.logger.error(`\u274C getMeeting error:`, err);
      return null;
    }
  }
  async upsertParticipants(meetingId, participants) {
    try {
      const stmt = this.db.prepare(`
        INSERT INTO meeting_participants (
          meeting_id, participant_aad_id, display_name, email, joined_at, left_at
        ) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(meeting_id, participant_aad_id) DO UPDATE SET
          display_name = excluded.display_name,
          email = excluded.email,
          joined_at = COALESCE(excluded.joined_at, meeting_participants.joined_at),
          left_at = excluded.left_at
      `);
      for (const participant of participants) {
        stmt.run(
          meetingId,
          participant.participantAadId,
          participant.displayName,
          participant.email || null,
          participant.joinedAt || null,
          participant.leftAt || null
        );
      }
      this.logger.debug(`\u2705 Upserted ${participants.length} participants for meeting: ${meetingId}`);
    } catch (err) {
      this.logger.error(`\u274C upsertParticipants error:`, err);
      throw err;
    }
  }
  async getParticipants(meetingId) {
    try {
      const stmt = this.db.prepare(`SELECT * FROM meeting_participants WHERE meeting_id = ?`);
      const rows = stmt.all(meetingId);
      return rows.map((row) => ({
        id: row.id,
        meetingId: row.meeting_id,
        participantAadId: row.participant_aad_id,
        displayName: row.display_name,
        email: row.email || void 0,
        joinedAt: row.joined_at || void 0,
        leftAt: row.left_at || void 0
      }));
    } catch (err) {
      this.logger.error(`\u274C getParticipants error:`, err);
      return [];
    }
  }
  async appendTranscriptChunk(input) {
    const now = (/* @__PURE__ */ new Date()).toISOString();
    try {
      const seqStmt = this.db.prepare(
        `SELECT MAX(sequence_number) as max_seq FROM transcript_chunks WHERE meeting_id = ?`
      );
      const seqResult = seqStmt.get(input.meetingId);
      const sequenceNumber = (seqResult?.max_seq ?? -1) + 1;
      const stmt = this.db.prepare(`
        INSERT INTO transcript_chunks (
          meeting_id, timestamp_utc, speaker, speaker_aad_id, text, confidence, source, sequence_number, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      const result = stmt.run(
        input.meetingId,
        input.timestampUtc,
        input.speaker,
        input.speakerAadId || null,
        input.text,
        input.confidence,
        input.source,
        sequenceNumber,
        now
      );
      this.logger.debug(`\u2705 Appended transcript chunk for meeting: ${input.meetingId}, seq: ${sequenceNumber}`);
      return {
        id: Number(result.lastInsertRowid),
        meetingId: input.meetingId,
        timestampUtc: input.timestampUtc,
        speaker: input.speaker,
        speakerAadId: input.speakerAadId,
        text: input.text,
        confidence: input.confidence,
        source: input.source,
        sequenceNumber,
        createdAt: now
      };
    } catch (err) {
      this.logger.error(`\u274C appendTranscriptChunk error:`, err);
      throw err;
    }
  }
  async getTranscriptChunks(meetingId) {
    try {
      const stmt = this.db.prepare(`SELECT * FROM transcript_chunks WHERE meeting_id = ? ORDER BY sequence_number ASC, timestamp_utc ASC`);
      const rows = stmt.all(meetingId);
      return rows.map((row) => ({
        id: row.id,
        meetingId: row.meeting_id,
        timestampUtc: row.timestamp_utc,
        speaker: row.speaker,
        speakerAadId: row.speaker_aad_id || void 0,
        text: row.text,
        confidence: row.confidence,
        source: row.source,
        sequenceNumber: row.sequence_number || void 0,
        createdAt: row.created_at
      }));
    } catch (err) {
      this.logger.error(`\u274C getTranscriptChunks error:`, err);
      return [];
    }
  }
  async getTranscriptByMeetingId(meetingId) {
    try {
      const meeting = await this.getMeeting(meetingId);
      if (!meeting) {
        this.logger.warn(`Meeting not found: ${meetingId}`);
        return null;
      }
      const participants = await this.getParticipants(meetingId);
      const chunks = await this.getTranscriptChunks(meetingId);
      return {
        meetingId: meeting.meetingId,
        title: meeting.title,
        startedAt: meeting.startedAt,
        endedAt: meeting.endedAt,
        participants,
        chunks,
        totalChunks: chunks.length
      };
    } catch (err) {
      this.logger.error(`\u274C getTranscriptByMeetingId error:`, err);
      return null;
    }
  }
  async getMeetingsByStatus(status) {
    try {
      const stmt = this.db.prepare(`SELECT * FROM meetings WHERE status = ?`);
      const rows = stmt.all(status);
      return rows.map((row) => ({
        meetingId: row.meeting_id,
        joinUrl: row.join_url,
        organizerAadId: row.organizer_aad_id,
        organizerDisplayName: row.organizer_display_name || void 0,
        organizerEmail: row.organizer_email || void 0,
        startedAt: row.started_at,
        endedAt: row.ended_at || void 0,
        status: row.status,
        title: row.title || void 0,
        conversationId: row.conversation_id || void 0,
        requestedByAadId: row.requested_by_aad_id || void 0,
        createdAt: row.created_at,
        updatedAt: row.updated_at
      }));
    } catch (err) {
      this.logger.error(`\u274C getMeetingsByStatus error:`, err);
      return [];
    }
  }
  close() {
    if (this.db) {
      this.db.close();
      this.logger.debug("\u{1F50C} Closed SQLite database connection");
    }
  }
};

// src/storage/storageFactory.ts
var StorageFactory = class {
  static async createStorage(logger2, config) {
    const dbConfig = config || DATABASE_CONFIG;
    let storage2;
    if (dbConfig.type === "mssql") {
      try {
        logger2.debug("\u{1F527} Initializing MSSQL storage...");
        storage2 = new MssqlKVStore(logger2.child("mssql"), dbConfig);
        await storage2.initialize();
        logger2.debug("\u2705 MSSQL storage initialized successfully");
        return storage2;
      } catch (error) {
        logger2.warn("\u26A0\uFE0F Failed to initialize MSSQL storage, falling back to SQLite:", error);
      }
    }
    logger2.debug("\u{1F527} Initializing SQLite storage...");
    storage2 = new SqliteKVStore(logger2.child("sqlite"), dbConfig.sqlitePath);
    await storage2.initialize();
    logger2.debug("\u2705 SQLite storage initialized successfully");
    return storage2;
  }
};

// src/index.ts
init_config();

// src/utils/chatHistory.ts
function stripHtml(html) {
  return html.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, " ").trim();
}
function toRecord(msg, conversationId) {
  if (msg.deletedDateTime || msg.messageType !== "message") return null;
  const name = msg.from?.user?.displayName || msg.from?.application?.displayName || "Unknown";
  const content = msg.body.contentType === "html" ? stripHtml(msg.body.content) : (msg.body.content || "").trim();
  if (!content) return null;
  return {
    conversation_id: conversationId,
    role: "user",
    content,
    timestamp: msg.createdDateTime,
    activity_id: msg.id,
    name
  };
}
async function fetchGroupChatHistory(appGraph, chatId, logger2, limit = 50) {
  try {
    logger2.debug(`[chatHistory] Fetching group chat history for ${chatId} (limit=${limit})`);
    const data = await appGraph.call(
      (id) => ({
        method: "get",
        path: `/chats/${id}/messages?$top=${limit}`
      }),
      chatId
    );
    const records = (data?.value ?? []).map((m) => toRecord(m, chatId)).filter((r) => r !== null);
    logger2.debug(`[chatHistory] Got ${records.length} messages from group chat ${chatId}`);
    return records;
  } catch (err) {
    logger2.warn(
      `[chatHistory] Group chat fetch failed for ${chatId}: ${err instanceof Error ? err.message : String(err)}`
    );
    return [];
  }
}
async function fetchChannelHistory(appGraph, teamId, channelId, conversationId, logger2, limit = 50) {
  try {
    logger2.debug(`[chatHistory] Fetching channel history for team=${teamId} channel=${channelId}`);
    const data = await appGraph.call(
      (tId, cId) => ({
        method: "get",
        path: `/teams/${tId}/channels/${cId}/messages?$top=${limit}`
      }),
      teamId,
      channelId
    );
    const records = (data?.value ?? []).map((m) => toRecord(m, conversationId)).filter((r) => r !== null);
    logger2.debug(`[chatHistory] Got ${records.length} messages from channel ${channelId}`);
    return records;
  } catch (err) {
    logger2.warn(
      `[chatHistory] Channel fetch failed for team=${teamId} channel=${channelId}: ${err instanceof Error ? err.message : String(err)}`
    );
    return [];
  }
}

// src/storage/conversationMemory.ts
var ConversationMemory = class {
  constructor(store, conversationId) {
    this.store = store;
    this.conversationId = conversationId;
  }
  async addMessages(messages) {
    await this.store.addMessages(messages);
  }
  async values() {
    const result = this.store.get(this.conversationId);
    return Promise.resolve(result).then((messages) => messages || []);
  }
  async length() {
    const result = this.store.countMessages(this.conversationId);
    return Promise.resolve(result);
  }
  async clear() {
    await this.store.clearConversation(this.conversationId);
  }
  async getMessagesByTimeRange(startTime, endTime) {
    const result = this.store.getMessagesByTimeRange(this.conversationId, startTime, endTime);
    return Promise.resolve(result);
  }
  async getRecentMessages(limit) {
    const result = this.store.getRecentMessages(this.conversationId, limit);
    return Promise.resolve(result);
  }
  async getFilteredMessages(conversationId, keywords, startTime, endTime, participants, maxResults) {
    const result = this.store.getFilteredMessages(
      conversationId,
      keywords,
      startTime,
      endTime,
      participants,
      maxResults
    );
    return Promise.resolve(result);
  }
};

// src/utils/messageContext.ts
async function getConversationParticipantsFromAPI(api, conversationId) {
  try {
    const members = await api.conversations.members(conversationId).get();
    if (Array.isArray(members)) {
      const participants = members.map((member) => ({
        name: member.name || "Unknown",
        id: member.objectId || member.id
      }));
      return participants;
    } else {
      return [];
    }
  } catch (error) {
    return [];
  }
}
async function createMessageContext(storage2, activity, api, appGraph) {
  const text = (activity.text || "").replace(/<at>[^<]*<\/at>\s*/g, "").trim();
  const conversationId = `${activity.conversation.id}`;
  const userId = activity.from.id;
  const userAadId = activity.from.aadObjectId;
  const userName = activity.from.name || "User";
  const timestamp = activity.timestamp?.toString() || "Unknown";
  const isPersonalChat = activity.conversation.conversationType === "personal";
  const activityId = activity.id;
  let members = [];
  if (api) {
    members = await getConversationParticipantsFromAPI(api, conversationId);
  }
  const memory = new ConversationMemory(storage2, conversationId);
  const now = /* @__PURE__ */ new Date();
  const startTime = new Date(now.getTime() - 24 * 60 * 60 * 1e3).toISOString();
  const endTime = now.toISOString();
  const citations = [];
  const context = {
    text,
    conversationId,
    userId,
    userAadId,
    userName,
    timestamp,
    isPersonalChat,
    activityId,
    members,
    memory,
    database: storage2,
    appGraph,
    startTime,
    endTime,
    citations
  };
  return context;
}

// src/index.ts
var autoJoinedMeetings = /* @__PURE__ */ new Map();
var logger = new teams_common.ConsoleLogger("missa", { level: "debug" });
var createTokenFactory = () => {
  return async (scope, tenantId) => {
    const managedIdentityCredential = new identity.ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId
    });
    return tokenResponse.token;
  };
};
var tokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory()
};
var options = process.env.BOT_TYPE === "UserAssignedMsi" ? { ...tokenCredentials } : { plugins: [new teams_dev.DevtoolsPlugin()] };
var app = new teams_apps.App({
  ...options,
  logger
});
var storage;
var feedbackStorage;
app.on("message.submit.feedback", async ({ activity }) => {
  try {
    if (!feedbackStorage) {
      logger.warn("feedbackStorage not yet initialized \u2014 ignoring feedback event");
      return;
    }
    const { reaction, feedback: feedbackJson } = activity.value.actionValue;
    if (!activity.replyToId) {
      logger.warn(`No replyToId found for messageId ${activity.id}`);
      return;
    }
    const success = await feedbackStorage.recordFeedback(
      activity.replyToId,
      reaction,
      feedbackJson
    );
    if (success) {
      logger.debug(`\u2705 Successfully recorded feedback for message ${activity.replyToId}`);
    } else {
      logger.warn(`Failed to record feedback for message ${activity.replyToId}`);
    }
  } catch (error) {
    logger.error(
      `Error processing feedback: ${error instanceof Error ? error.message : "Unknown error"}`
    );
  }
});
app.on("message", async ({ send, activity, api, appGraph }) => {
  if (!storage) {
    logger.warn("storage not yet initialized \u2014 ignoring message event");
    return;
  }
  const botMentioned = activity.entities?.some((e) => e.type === "mention");
  const context = botMentioned ? await createMessageContext(storage, activity, api, appGraph) : await createMessageContext(storage, activity, void 0, appGraph);
  let trackedMessages;
  if (!activity.conversation.isGroup || botMentioned) {
    await send({ type: "typing" });
    const manager = new ManagerPrompt(context, logger.child("manager"));
    const result = await manager.processRequest();
    const formattedResult = finalizePromptResponse(result.response, context, logger);
    const sent = await send(formattedResult);
    formattedResult.id = sent.id;
    trackedMessages = createMessageRecords([activity, formattedResult]);
  } else {
    trackedMessages = createMessageRecords([activity]);
  }
  logger.debug(trackedMessages);
  await context.memory.addMessages(trackedMessages);
});
app.on("install.add", async ({ send, activity, appGraph }) => {
  let historyCount = 0;
  if (storage) {
    try {
      const conversationId = activity.conversation.id;
      const channelData = activity.channelData;
      const teamId = channelData?.team?.id;
      const channelId = channelData?.channel?.id;
      let records;
      if (teamId && channelId) {
        records = await fetchChannelHistory(appGraph, teamId, channelId, conversationId, logger);
      } else if (activity.conversation.isGroup) {
        records = await fetchGroupChatHistory(appGraph, conversationId, logger);
      }
      if (records && records.length > 0) {
        await storage.addMessages(records);
        historyCount = records.length;
        logger.info(`[install.add] Loaded ${historyCount} historical messages for ${conversationId}`);
      }
    } catch (err) {
      logger.warn(
        `[install.add] History backfill failed: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  }
  const historyNote = historyCount > 0 ? `I've loaded **${historyCount}** recent messages so I already have context on your conversation.

` : "";
  await send(
    "\u{1F44B} Hi! I'm **Missa**, your intelligent meeting assistant!\n\n" + historyNote + "Here's what I can do:\n\n\u{1F4DD} **Summarize** conversations and meetings\n\u2705 **Action Items** \u2014 find and track tasks from your chats\n\u{1F50D} **Search** through conversation history\n\u{1F4CB} **Meeting Notes** \u2014 get structured transcripts and summaries\n\u{1F399}\uFE0F **Auto-Join** \u2014 I automatically join meetings when they start in this chat\n\u{1F517} **Smart Join** \u2014 just say `@Missa join meeting` (no URL needed if a meeting is active)\n\u23F9\uFE0F **Stop Recording** \u2014 stop live transcription and save the notes\n\nUse the command menu or @mention me with your request!"
  );
});
app.on("event", async ({ activity, send }) => {
  const eventName = activity.name;
  if (eventName === "application/vnd.microsoft.meetingStart") {
    const details = activity.value?.details;
    const joinUrl = details?.joinWebUrl || details?.joinUrl;
    const meetingTitle = details?.title || "Teams Meeting";
    const conversationId = activity.conversation.id;
    logger.info(`[MeetingEvent] Meeting started in ${conversationId}: "${meetingTitle}"`);
    if (!joinUrl) {
      logger.warn("[MeetingEvent] No joinWebUrl in meetingStart event \u2014 cannot auto-join");
      return;
    }
    if (autoJoinedMeetings.has(conversationId)) {
      logger.info(`[MeetingEvent] Already joined meeting for ${conversationId}, skipping`);
      return;
    }
    try {
      await send("\u{1F399}\uFE0F Meeting started \u2014 Missa is joining to capture the transcript...");
      const client = getMeetingMediaBotClient(logger);
      const result = await client.startMeetingCapture(joinUrl);
      if (result.success && result.callId) {
        autoJoinedMeetings.set(conversationId, result.callId);
        logger.info(`[MeetingEvent] Auto-joined meeting, callId: ${result.callId}`);
        await send(
          `\u2705 I've joined **"${meetingTitle}"** and will transcribe in real-time.

\u{1F4CC} **Tip:** Ask the meeting organizer to enable **Teams transcription** (meeting controls \u2192 ... \u2192 Start transcription) for best results.

When the meeting ends I'll automatically leave and offer a summary.`
        );
      } else {
        logger.error(`[MeetingEvent] Failed to auto-join: ${result.error}`);
        await send(`\u26A0\uFE0F Could not auto-join the meeting: ${result.error || "Unknown error"}. You can still join manually with \`@Missa join meeting <url>\`.`);
      }
    } catch (err) {
      logger.error("[MeetingEvent] Error during auto-join:", err);
    }
  }
  if (eventName === "application/vnd.microsoft.meetingEnd") {
    const conversationId = activity.conversation.id;
    const callId = autoJoinedMeetings.get(conversationId);
    logger.info(`[MeetingEvent] Meeting ended in ${conversationId}, callId: ${callId || "none"}`);
    if (!callId) {
      return;
    }
    try {
      const client = getMeetingMediaBotClient(logger);
      await client.stopMeetingCapture(callId);
      autoJoinedMeetings.delete(conversationId);
      logger.info(`[MeetingEvent] Left meeting ${callId} after meetingEnd event`);
      await send(
        "\u{1F4CB} Meeting ended \u2014 I've left and saved the transcript.\n\nUse **@Missa summarize meeting** to get a structured summary with action items, decisions, and key points."
      );
    } catch (err) {
      logger.error("[MeetingEvent] Error during auto-leave:", err);
      autoJoinedMeetings.delete(conversationId);
    }
  }
});
(async () => {
  const port = process.env.PORT || process.env.port || 3978;
  const internalApiPort = process.env.INTERNAL_API_PORT || 3980;
  try {
    validateEnvironment(logger);
    logModelConfigs(logger);
    const config = loadConfig(logger);
    storage = await StorageFactory.createStorage(logger.child("storage"));
    feedbackStorage = storage;
    logger.debug("\u2705 Storage initialized successfully");
    initializeMeetingRoutes(storage, config, logger.child("meeting-api"));
    app.http.use(express__default.default.json({ type: "application/json" }));
    app.http.use("/api", meetingApi_default);
    const internalApp = express__default.default();
    internalApp.use(express__default.default.json());
    internalApp.use("/api", meetingApi_default);
    internalApp.listen(internalApiPort, () => {
      logger.info(`\u{1F4E1} Internal API server started on port ${internalApiPort}`);
    });
  } catch (error) {
    logger.error("\u274C Configuration error:", error);
    process.exit(1);
  }
  await app.start(port);
  logger.debug(`\u{1F680} Collab Agent started on port ${port}`);
  const appConfig = loadConfig(logger);
  (async () => {
    try {
      const meetingBotClient = getMeetingMediaBotClient(logger.child("health"));
      const reachable = await meetingBotClient.checkHealth();
      if (reachable) {
        logger.info(`Meeting media bot reachable at ${appConfig.meetingMediaBotUrl}`);
      } else {
        logger.warn(`Meeting media bot UNREACHABLE at ${appConfig.meetingMediaBotUrl} \u2014 meeting capture will fail`);
      }
    } catch (err) {
      logger.warn(`Meeting media bot health check error: ${err instanceof Error ? err.message : err}`);
    }
  })();
})();
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map