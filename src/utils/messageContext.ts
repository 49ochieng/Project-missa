import { CitationAppearance, Client, IMessageActivity } from "@microsoft/teams.api";
import type { Client as GraphClient } from "@microsoft/teams.graph";
import { ConversationMemory } from "../storage/conversationMemory";
import { IDatabase } from "../storage/database";

/**
 * Context object that stores all important information for processing a message
 */
export interface MessageContext {
  text: string;
  conversationId: string;
  userId?: string;
  userAadId?: string;
  userName: string;
  timestamp: string;
  isPersonalChat: boolean;
  activityId: string;
  members: Array<{ name: string; id: string }>; // Available conversation members
  memory: ConversationMemory; // get convo memory by agent type
  database: IDatabase; // Database instance for direct storage operations
  appGraph?: GraphClient; // SDK Graph client (app credentials) for Graph API calls
  startTime: string;
  endTime: string;
  citations: CitationAppearance[];
}

async function getConversationParticipantsFromAPI(
  api: Client,
  conversationId: string
): Promise<Array<{ name: string; id: string }>> {
  try {
    const members = await api.conversations.members(conversationId).get();

    if (Array.isArray(members)) {
      const participants = members.map((member) => ({
        name: member.name || "Unknown",
        id: member.objectId || member.id,
      }));
      return participants;
    } else {
      return [];
    }
  } catch (error) {
    return [];
  }
}

/**
 * Factory function to create a MessageContext from a Teams activity
 */
export async function createMessageContext(
  storage: IDatabase,
  activity: IMessageActivity,
  api?: Client,
  appGraph?: GraphClient
): Promise<MessageContext> {
  // Strip @mention markup so the AI sees clean text (e.g. "<at>Missa</at> join meeting" → "join meeting")
  const text = (activity.text || "").replace(/<at>[^<]*<\/at>\s*/g, "").trim();
  const conversationId = `${activity.conversation.id}`;
  const userId = activity.from.id;
  const userAadId = activity.from.aadObjectId;
  const userName = activity.from.name || "User";
  const timestamp = activity.timestamp?.toString() || "Unknown";
  const isPersonalChat = activity.conversation.conversationType === "personal";
  const activityId = activity.id;

  // Fetch members for group conversations
  let members: Array<{ name: string; id: string }> = [];
  if (api) {
    members = await getConversationParticipantsFromAPI(api, conversationId);
  }

  const memory = new ConversationMemory(storage, conversationId);

  const now = new Date();

  const startTime = new Date(now.getTime() - 24 * 60 * 60 * 1000).toISOString();
  const endTime = now.toISOString();
  const citations: CitationAppearance[] = [];

  const context: MessageContext = {
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
    database: storage,
    appGraph,
    startTime,
    endTime,
    citations,
  };

  return context;
}
