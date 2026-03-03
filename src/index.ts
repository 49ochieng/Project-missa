import { ManagedIdentityCredential } from "@azure/identity";
import { TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import express from "express";
import { ManagerPrompt } from "./agent/manager";
import meetingApiRoutes, { initializeMeetingRoutes } from "./routes/meetingApi";
import { IDatabase } from "./storage/database";
import { StorageFactory } from "./storage/storageFactory";
import { loadConfig, logModelConfigs, validateEnvironment } from "./utils/config";
import { fetchChannelHistory, fetchGroupChatHistory } from "./utils/chatHistory";
import { createMessageContext } from "./utils/messageContext";
import { createMessageRecords, finalizePromptResponse } from "./utils/utils";

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
    "🎙️ **Join Meeting** — I can join a Teams meeting and transcribe it live\n" +
    "⏹️ **Stop Recording** — stop live transcription and save the notes\n\n" +
    "Use the command menu or @mention me with your request!"
  );
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
})();
