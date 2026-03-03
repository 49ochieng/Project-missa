/**
 * Utilities for backfilling conversation history from Microsoft Graph
 * when the bot is first added to a group chat or channel.
 *
 * Requires RSC permissions:
 *   - ChatMessage.Read.Chat       (group chats)
 *   - ChannelMessage.Read.Group   (team channels)
 */

import type { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import type { MessageRecord } from "../storage/types";

// ─── Graph API response shapes ──────────────────────────────────────────────

interface GraphMessage {
  id: string;
  createdDateTime: string;
  from?: {
    user?: { displayName?: string; id?: string };
    application?: { displayName?: string; id?: string };
  };
  body: { contentType: "html" | "text"; content: string };
  /** "message" | "chatEvent" | "unknownFutureValue" */
  messageType?: string;
  deletedDateTime?: string | null;
}

interface GraphPagedResponse {
  value: GraphMessage[];
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .trim();
}

function toRecord(msg: GraphMessage, conversationId: string): MessageRecord | null {
  // Skip deleted messages and system events
  if (msg.deletedDateTime || msg.messageType !== "message") return null;

  const name =
    msg.from?.user?.displayName ||
    msg.from?.application?.displayName ||
    "Unknown";

  const content =
    msg.body.contentType === "html"
      ? stripHtml(msg.body.content)
      : (msg.body.content || "").trim();

  if (!content) return null;

  return {
    conversation_id: conversationId,
    role: "user",
    content,
    timestamp: msg.createdDateTime,
    activity_id: msg.id,
    name,
  };
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Fetch recent messages from a group chat via Graph API.
 * Requires RSC permission: ChatMessage.Read.Chat
 */
export async function fetchGroupChatHistory(
  appGraph: GraphClient,
  chatId: string,
  logger: ILogger,
  limit = 50
): Promise<MessageRecord[]> {
  try {
    logger.debug(`[chatHistory] Fetching group chat history for ${chatId} (limit=${limit})`);

    const data = (await appGraph.call(
      (id: string) => ({
        method: "get" as const,
        path: `/chats/${id}/messages?$top=${limit}`,
      }),
      chatId
    )) as GraphPagedResponse;

    const records = (data?.value ?? [])
      .map((m) => toRecord(m, chatId))
      .filter((r): r is MessageRecord => r !== null);

    logger.debug(`[chatHistory] Got ${records.length} messages from group chat ${chatId}`);
    return records;
  } catch (err) {
    logger.warn(
      `[chatHistory] Group chat fetch failed for ${chatId}: ${
        err instanceof Error ? err.message : String(err)
      }`
    );
    return [];
  }
}

/**
 * Fetch recent messages from a Teams channel via Graph API.
 * Requires RSC permission: ChannelMessage.Read.Group
 */
export async function fetchChannelHistory(
  appGraph: GraphClient,
  teamId: string,
  channelId: string,
  conversationId: string,
  logger: ILogger,
  limit = 50
): Promise<MessageRecord[]> {
  try {
    logger.debug(`[chatHistory] Fetching channel history for team=${teamId} channel=${channelId}`);

    const data = (await appGraph.call(
      (tId: string, cId: string) => ({
        method: "get" as const,
        path: `/teams/${tId}/channels/${cId}/messages?$top=${limit}`,
      }),
      teamId,
      channelId
    )) as GraphPagedResponse;

    const records = (data?.value ?? [])
      .map((m) => toRecord(m, conversationId))
      .filter((r): r is MessageRecord => r !== null);

    logger.debug(`[chatHistory] Got ${records.length} messages from channel ${channelId}`);
    return records;
  } catch (err) {
    logger.warn(
      `[chatHistory] Channel fetch failed for team=${teamId} channel=${channelId}: ${
        err instanceof Error ? err.message : String(err)
      }`
    );
    return [];
  }
}
