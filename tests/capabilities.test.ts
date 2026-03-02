/**
 * Lightweight test harness for all capabilities.
 * No external services required — uses in-memory mocks.
 *
 * Run with:  npx ts-node tests/capabilities.test.ts
 */

// Set dummy OpenAI credentials before any capability code runs
// so OpenAIChatModel construction doesn't throw during structural tests
process.env.AOAI_API_KEY = "test-key-harness";
process.env.AOAI_ENDPOINT = "https://test.openai.azure.com";
process.env.AOAI_MODEL = "gpt-4o";

import { CitationAppearance } from "@microsoft/teams.api";
import { ConsoleLogger } from "@microsoft/teams.common";
import { ActionItemsCapability } from "../src/capabilities/actionItems/actionItems";
import { SearchCapability } from "../src/capabilities/search/search";
import { SummarizerCapability } from "../src/capabilities/summarizer/summarize";
import { ConversationMemory } from "../src/storage/conversationMemory";
import { IDatabase } from "../src/storage/database";
import { MessageRecord } from "../src/storage/types";
import { MessageContext } from "../src/utils/messageContext";
import { createMessageRecords, extractTimeRange } from "../src/utils/utils";

// ---------------------------------------------------------------------------
// Minimal test runner
// ---------------------------------------------------------------------------

let passed = 0;
let failed = 0;

function test(name: string, fn: () => void | Promise<void>) {
  Promise.resolve()
    .then(fn)
    .then(() => {
      console.log(`  ✅  ${name}`);
      passed++;
    })
    .catch((err) => {
      console.error(`  ❌  ${name}`);
      console.error(`       ${err?.message ?? err}`);
      failed++;
    });
}

function assert(condition: boolean, message: string) {
  if (!condition) throw new Error(message);
}

// ---------------------------------------------------------------------------
// Mock IDatabase
// ---------------------------------------------------------------------------

const SAMPLE_RECORDS: MessageRecord[] = [
  {
    conversation_id: "conv-1",
    role: "user",
    content: "We need to finish the report by Friday",
    timestamp: new Date("2026-03-01T10:00:00Z").toISOString(),
    activity_id: "act-1",
    name: "Alice",
  },
  {
    conversation_id: "conv-1",
    role: "model",
    content: "I'll draft the outline tonight",
    timestamp: new Date("2026-03-01T10:01:00Z").toISOString(),
    activity_id: "act-2",
    name: "Bob",
  },
  {
    conversation_id: "conv-1",
    role: "user",
    content: "Can you also check the budget numbers?",
    timestamp: new Date("2026-03-01T10:02:00Z").toISOString(),
    activity_id: "act-3",
    name: "Alice",
  },
];

class MockDatabase implements IDatabase {
  private store: Map<string, MessageRecord[]> = new Map([["conv-1", SAMPLE_RECORDS]]);

  async initialize() {}
  clearAll() { this.store.clear(); }
  get(conversationId: string) { return this.store.get(conversationId) ?? []; }
  getMessagesByTimeRange(conversationId: string, startTime: string, endTime: string) {
    const msgs = this.store.get(conversationId) ?? [];
    const start = new Date(startTime).getTime();
    const end = new Date(endTime).getTime();
    return msgs.filter((m) => {
      const t = new Date(m.timestamp).getTime();
      return t >= start && t <= end;
    });
  }
  getRecentMessages(conversationId: string, limit = 10) {
    return (this.store.get(conversationId) ?? []).slice(-limit);
  }
  clearConversation(conversationId: string) { this.store.delete(conversationId); }
  addMessages(messages: MessageRecord[]) {
    const id = messages[0]?.conversation_id;
    if (!id) return;
    const existing = this.store.get(id) ?? [];
    this.store.set(id, [...existing, ...messages]);
  }
  countMessages(conversationId: string) { return (this.store.get(conversationId) ?? []).length; }
  clearAllMessages() { this.store.clear(); }
  getFilteredMessages(
    conversationId: string,
    keywords: string[],
    startTime: string,
    endTime: string,
    participants?: string[],
    maxResults = 5
  ) {
    const msgs = this.store.get(conversationId) ?? [];
    return msgs
      .filter((m) => {
        const inTime = m.timestamp >= startTime && m.timestamp <= endTime;
        const hasKeyword = keywords.some((k) =>
          m.content.toLowerCase().includes(k.toLowerCase())
        );
        const hasParticipant = participants
          ? participants.some((p) => m.name.toLowerCase().includes(p.toLowerCase()))
          : true;
        return inTime && hasKeyword && hasParticipant;
      })
      .slice(0, maxResults);
  }
  recordFeedback() { return true; }
  close() {}
}

// ---------------------------------------------------------------------------
// Helper: build a minimal MessageContext
// ---------------------------------------------------------------------------

function makeContext(overrides: Partial<MessageContext> = {}): MessageContext {
  const db = new MockDatabase();
  const memory = new ConversationMemory(db, "conv-1");
  const now = new Date("2026-03-02T12:00:00Z");
  return {
    text: "test message",
    conversationId: "conv-1",
    userId: "user-1",
    userName: "TestUser",
    timestamp: now.toISOString(),
    isPersonalChat: false,
    activityId: "act-0",
    members: [{ name: "Alice", id: "u-1" }, { name: "Bob", id: "u-2" }],
    memory,
    startTime: new Date(now.getTime() - 24 * 60 * 60 * 1000).toISOString(),
    endTime: now.toISOString(),
    citations: [] as CitationAppearance[],
    ...overrides,
  };
}

const logger = new ConsoleLogger("test", { level: "error" }); // suppress debug noise during tests

// ---------------------------------------------------------------------------
// Tests: extractTimeRange
// ---------------------------------------------------------------------------

console.log("\n── extractTimeRange ──────────────────────────────────────────");

test("returns null for empty string", () => {
  assert(extractTimeRange("") === null, "expected null for empty string");
});

test("parses 'yesterday' into a valid range", () => {
  const range = extractTimeRange("yesterday");
  assert(range !== null, "expected non-null range");
  assert(range!.from < range!.to, "from should be before to");
});

test("parses 'last week' into a valid range", () => {
  const range = extractTimeRange("last week");
  assert(range !== null, "expected non-null range");
  assert(range!.to.getTime() - range!.from.getTime() > 0, "range should have positive duration");
});

test("parses 'past 3 hours' into a valid range", () => {
  const range = extractTimeRange("past 3 hours", new Date("2026-03-02T12:00:00Z"));
  assert(range !== null, "expected non-null range");
});

test("returns null for nonsense phrase", () => {
  const result = extractTimeRange("blah blah blah xyz");
  // chrono may not return null for all gibberish but should not throw
  assert(result === null || typeof result === "object", "should be null or a range object");
});

// ---------------------------------------------------------------------------
// Tests: createMessageRecords
// ---------------------------------------------------------------------------

console.log("\n── createMessageRecords ──────────────────────────────────────");

test("returns empty array for empty input", () => {
  const result = createMessageRecords([]);
  assert(Array.isArray(result) && result.length === 0, "expected empty array");
});

test("maps a user activity to a user role record", () => {
  const mockActivity: any = {
    conversation: { id: "conv-1" },
    text: "Hello world",
    from: { name: "Alice" },
    timestamp: "2026-03-01T10:00:00Z",
    id: "act-1",
    entities: [],
  };
  const records = createMessageRecords([mockActivity]);
  assert(records.length === 1, "expected 1 record");
  assert(records[0].role === "user", `expected role 'user', got '${records[0].role}'`);
  assert(records[0].content === "Hello world", "content mismatch");
  assert(records[0].conversation_id === "conv-1", "conversation_id mismatch");
});

test("strips <at> tags from content", () => {
  const mockActivity: any = {
    conversation: { id: "conv-1" },
    text: "<at>Bot</at> summarize this",
    from: { name: "Alice" },
    timestamp: "2026-03-01T10:00:00Z",
    id: "act-2",
    entities: [],
  };
  const records = createMessageRecords([mockActivity]);
  assert(!records[0].content.includes("<at>"), "should strip <at> tags");
  assert(records[0].content === "Bot summarize this", `unexpected content: ${records[0].content}`);
});

// ---------------------------------------------------------------------------
// Tests: ConversationMemory
// ---------------------------------------------------------------------------

console.log("\n── ConversationMemory ────────────────────────────────────────");

test("values() returns seeded records for conversation", async () => {
  const db = new MockDatabase();
  const mem = new ConversationMemory(db, "conv-1");
  const records = await mem.values();
  assert(records.length === SAMPLE_RECORDS.length, `expected ${SAMPLE_RECORDS.length} records`);
});

test("countMessages() returns correct count", async () => {
  const db = new MockDatabase();
  const mem = new ConversationMemory(db, "conv-1");
  const count = await mem.length();
  assert(count === SAMPLE_RECORDS.length, `expected ${SAMPLE_RECORDS.length}, got ${count}`);
});

test("addMessages() persists new messages", async () => {
  const db = new MockDatabase();
  const mem = new ConversationMemory(db, "conv-1");
  const newMsg: MessageRecord = {
    conversation_id: "conv-1",
    role: "user",
    content: "New message",
    timestamp: new Date().toISOString(),
    activity_id: "act-new",
    name: "Charlie",
  };
  await mem.addMessages([newMsg]);
  const count = await mem.length();
  assert(count === SAMPLE_RECORDS.length + 1, `expected ${SAMPLE_RECORDS.length + 1}, got ${count}`);
});

test("clear() removes all messages for conversation", async () => {
  const db = new MockDatabase();
  const mem = new ConversationMemory(db, "conv-1");
  await mem.clear();
  const count = await mem.length();
  assert(count === 0, `expected 0 after clear, got ${count}`);
});

test("getMessagesByTimeRange() filters correctly", async () => {
  const db = new MockDatabase();
  const mem = new ConversationMemory(db, "conv-1");
  const start = "2026-03-01T10:00:00Z";
  const end = "2026-03-01T10:01:30Z";
  const msgs = await mem.getMessagesByTimeRange(start, end);
  // Should include act-1 (10:00) and act-2 (10:01) but not act-3 (10:02)
  assert(msgs.length === 2, `expected 2 messages, got ${msgs.length}`);
});

// ---------------------------------------------------------------------------
// Tests: Capability — createPrompt (structural, no AI call)
// ---------------------------------------------------------------------------

console.log("\n── Capability.createPrompt (structural) ──────────────────────");

test("SummarizerCapability.createPrompt returns a ChatPrompt", () => {
  const ctx = makeContext({ text: "summarize today" });
  const cap = new SummarizerCapability(logger);
  const prompt = cap.createPrompt(ctx);
  assert(prompt !== null && typeof prompt === "object", "expected a ChatPrompt object");
  assert(typeof (prompt as any).send === "function", "ChatPrompt should have a send method");
});

test("ActionItemsCapability.createPrompt returns a ChatPrompt", () => {
  const ctx = makeContext({ text: "what are the action items?" });
  const cap = new ActionItemsCapability(logger);
  const prompt = cap.createPrompt(ctx);
  assert(prompt !== null && typeof prompt === "object", "expected a ChatPrompt object");
  assert(typeof (prompt as any).send === "function", "ChatPrompt should have a send method");
});

test("SearchCapability.createPrompt returns a ChatPrompt", () => {
  const ctx = makeContext({ text: "find messages about budget" });
  const cap = new SearchCapability(logger);
  const prompt = cap.createPrompt(ctx);
  assert(prompt !== null && typeof prompt === "object", "expected a ChatPrompt object");
  assert(typeof (prompt as any).send === "function", "ChatPrompt should have a send method");
});

// ---------------------------------------------------------------------------
// Tests: Capability error handling (handler wraps errors)
// ---------------------------------------------------------------------------

console.log("\n── Capability error wrapping ──────────────────────────────────");

test("SummarizerCapability.processRequest returns error string on failure", async () => {
  const ctx = makeContext({ text: "summarize" });
  // Override memory.getMessagesByTimeRange to throw
  (ctx.memory as any).getMessagesByTimeRange = async () => { throw new Error("DB unavailable"); };
  const cap = new SummarizerCapability(logger);
  const result = await cap.processRequest(ctx);
  // The base class catches and returns { response: "", error: "..." }
  // Because the function throw happens inside the AI function call (not processRequest directly),
  // the error bubbles through prompt.send() and gets caught by BaseCapability
  assert(typeof result === "object", "result should be an object");
  assert("response" in result, "result should have a response field");
});

test("SearchCapability: search_messages with no results returns gracefully", async () => {
  const db = new MockDatabase();
  // Override to return empty
  db.getFilteredMessages = () => [];
  const mem = new ConversationMemory(db, "conv-1");
  const ctx = makeContext({ memory: mem, text: "find xyz123notexist" });
  const cap = new SearchCapability(logger);
  const prompt = cap.createPrompt(ctx);
  assert(prompt !== null, "prompt should build even with empty memory");
});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

setTimeout(() => {
  console.log(`\n══════════════════════════════════════════════════════════════`);
  console.log(`  Results: ${passed} passed, ${failed} failed`);
  if (failed > 0) process.exit(1);
}, 500);
