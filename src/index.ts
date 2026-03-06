import { ManagedIdentityCredential } from "@azure/identity";
import { TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import express from "express";
import { ManagerPrompt } from "./agent/manager";
import meetingApiRoutes, { initializeMeetingRoutes } from "./routes/meetingApi";
import { getMeetingMediaBotClient } from "./services/meetingMediaBotClient";
import { IDatabase } from "./storage/database";
import { StorageFactory } from "./storage/storageFactory";
import { loadConfig, logModelConfigs, validateEnvironment } from "./utils/config";
import { fetchChannelHistory, fetchGroupChatHistory } from "./utils/chatHistory";
import { createMessageContext } from "./utils/messageContext";
import { createMessageRecords, finalizePromptResponse } from "./utils/utils";

// Tracks auto-joined meetings: conversationId → callId
// Used for auto-leave when meetingEnd event is received
const autoJoinedMeetings = new Map<string, string>();

const logger = new ConsoleLogger("missa", { level: "debug" });

const createTokenFactory = () => {
  return async (scope: string | string[], tenantId?: string): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

// Use managed identity in cloud environment, otherwise use devtools plugin for local development
const options =
  process.env.BOT_TYPE === "UserAssignedMsi"
    ? { ...tokenCredentials }
    : { plugins: [new DevtoolsPlugin()] };

const app = new App({
  ...options,
  logger,
});

// Initialize storage
let storage: IDatabase;
let feedbackStorage: IDatabase;

app.on("message.submit.feedback", async ({ activity }) => {
  try {
    if (!feedbackStorage) {
      logger.warn("feedbackStorage not yet initialized — ignoring feedback event");
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
      logger.debug(`✅ Successfully recorded feedback for message ${activity.replyToId}`);
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
    logger.warn("storage not yet initialized — ignoring message event");
    return;
  }
  const botMentioned = activity.entities?.some((e) => e.type === "mention");
  const context = botMentioned
    ? await createMessageContext(storage, activity, api, appGraph)
    : await createMessageContext(storage, activity, undefined, appGraph);

  let trackedMessages;

  if (!activity.conversation.isGroup || botMentioned) {
    // process request if One-on-One chat or if @mentioned in Groupchat
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const channelData = (activity as any).channelData as {
        team?: { id?: string };
        channel?: { id?: string };
      } | undefined;

      const teamId = channelData?.team?.id;
      const channelId = channelData?.channel?.id;

      let records;
      if (teamId && channelId) {
        // Installed in a Teams channel
        records = await fetchChannelHistory(appGraph, teamId, channelId, conversationId, logger);
      } else if (activity.conversation.isGroup) {
        // Installed in a group chat
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

  const historyNote = historyCount > 0
    ? `I've loaded **${historyCount}** recent messages so I already have context on your conversation.\n\n`
    : "";

  await send(
    "👋 Hi! I'm **Missa**, your intelligent meeting assistant!\n\n" +
    historyNote +
    "Here's what I can do:\n\n" +
    "📝 **Summarize** conversations and meetings\n" +
    "✅ **Action Items** — find and track tasks from your chats\n" +
    "🔍 **Search** through conversation history\n" +
    "📋 **Meeting Notes** — get structured transcripts and summaries\n" +
    "🎙️ **Auto-Join** — I automatically join meetings when they start in this chat\n" +
    "🔗 **Smart Join** — just say `@Missa join meeting` (no URL needed if a meeting is active)\n" +
    "⏹️ **Stop Recording** — stop live transcription and save the notes\n\n" +
    "Use the command menu or @mention me with your request!"
  );
});

/**
 * Teams Meeting Lifecycle Events (Tier 1 auto-join)
 * Received when a meeting starts/ends in a conversation where the bot is installed.
 * Requires manifest: bots[].meetingEventSubscription.onlineMeetingStarted/Ended = true
 */
app.on("event", async ({ activity, send }) => {
  const eventName = activity.name as string;

  if (eventName === "application/vnd.microsoft.meetingStart") {
    const details = ((activity as unknown as Record<string, unknown>).value as Record<string, unknown>)?.details as Record<string, string> | undefined;
    const joinUrl = details?.joinWebUrl || details?.joinUrl;
    const meetingTitle = details?.title || "Teams Meeting";
    const conversationId = activity.conversation.id;

    logger.info(`[MeetingEvent] Meeting started in ${conversationId}: "${meetingTitle}"`);

    if (!joinUrl) {
      logger.warn("[MeetingEvent] No joinWebUrl in meetingStart event — cannot auto-join");
      return;
    }

    // Don't double-join if already active
    if (autoJoinedMeetings.has(conversationId)) {
      logger.info(`[MeetingEvent] Already joined meeting for ${conversationId}, skipping`);
      return;
    }

    try {
      await send("🎙️ Meeting started — Missa is joining to capture the transcript...");

      const client = getMeetingMediaBotClient(logger);
      const result = await client.startMeetingCapture(joinUrl);

      if (result.success && result.callId) {
        autoJoinedMeetings.set(conversationId, result.callId);
        logger.info(`[MeetingEvent] Auto-joined meeting, callId: ${result.callId}`);
        await send(
          `✅ I've joined **"${meetingTitle}"** and will transcribe in real-time.\n\n` +
          `📌 **Tip:** Ask the meeting organizer to enable **Teams transcription** (meeting controls → ... → Start transcription) for best results.\n\n` +
          `When the meeting ends I'll automatically leave and offer a summary.`
        );
      } else {
        logger.error(`[MeetingEvent] Failed to auto-join: ${result.error}`);
        await send(`⚠️ Could not auto-join the meeting: ${result.error || "Unknown error"}. You can still join manually with \`@Missa join meeting <url>\`.`);
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
      // Meeting wasn't auto-joined (user may have manually joined or bot wasn't active)
      return;
    }

    try {
      const client = getMeetingMediaBotClient(logger);
      await client.stopMeetingCapture(callId);
      autoJoinedMeetings.delete(conversationId);
      logger.info(`[MeetingEvent] Left meeting ${callId} after meetingEnd event`);

      await send(
        "📋 Meeting ended — I've left and saved the transcript.\n\n" +
        "Use **@Missa summarize meeting** to get a structured summary with action items, decisions, and key points."
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

    // Load config
    const config = loadConfig(logger);

    // Initialize storage
    storage = await StorageFactory.createStorage(logger.child("storage"));
    feedbackStorage = storage;

    logger.debug("✅ Storage initialized successfully");

    // Initialize and start internal API server for meeting-media-bot communication
    initializeMeetingRoutes(storage, config, logger.child("meeting-api"));

    // Expose meeting API routes on the main Teams bot port (for Azure - port 3980 is internal only)
    app.http.use(express.json({ type: "application/json" }));
    app.http.use("/api", meetingApiRoutes);

    // Also start dedicated internal API server (for local dev where port 3980 is preferred)
    const internalApp = express();
    internalApp.use(express.json());
    internalApp.use("/api", meetingApiRoutes);

    internalApp.listen(internalApiPort, () => {
      logger.info(`📡 Internal API server started on port ${internalApiPort}`);
    });
  } catch (error) {
    logger.error("❌ Configuration error:", error);
    process.exit(1);
  }

  await app.start(port);

  logger.debug(`🚀 Collab Agent started on port ${port}`);

  // Non-blocking startup check: verify meeting-media-bot connectivity
  const appConfig = loadConfig(logger);
  (async () => {
    try {
      const meetingBotClient = getMeetingMediaBotClient(logger.child("health"));
      const reachable = await meetingBotClient.checkHealth();
      if (reachable) {
        logger.info(`Meeting media bot reachable at ${appConfig.meetingMediaBotUrl}`);
      } else {
        logger.warn(`Meeting media bot UNREACHABLE at ${appConfig.meetingMediaBotUrl} — meeting capture will fail`);
      }
    } catch (err) {
      logger.warn(`Meeting media bot health check error: ${err instanceof Error ? err.message : err}`);
    }
  })();
})();
