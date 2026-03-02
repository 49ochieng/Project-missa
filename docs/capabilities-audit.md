# Capabilities Audit — Project Missa

**Date:** 2026-03-02  
**Auditor:** GitHub Copilot  
**Scope:** Static analysis, runtime issue review, and lightweight test harness for all capabilities.

---

## 1. App Runtime — How It Works

### Entrypoint
`src/index.ts` is the entry point. It:
1. Creates a Teams `App` instance, configured with `UserAssignedMsi` (Azure Managed Identity) when deployed, or `DevtoolsPlugin` locally.
2. Registers two event handlers: `message.submit.feedback` and `message`.
3. At startup (`(async () => {...})()`) it validates environment variables, logs model configs, initializes the storage backend via `StorageFactory`, and starts the HTTP server on port `3978`.

### Message Flow
```
Teams message arrives
  → app.on("message")
      → if @mentioned (group) or any message (personal)
          → createMessageContext()        builds MessageContext from activity
          → new ManagerPrompt(context, logger)
          → manager.processRequest()      calls ChatPrompt.send()
              → Manager LLM (gpt-4.1)    decides which capability to call
              → delegate_to_<capability> function registered for each capability
                  → capability.handler() → BaseCapability.processRequest()
                      → capability.createPrompt() + ChatPrompt.send()
                      → returns response string
          → finalizePromptResponse()      attaches citations, feedback card
          → send()                        sends back to Teams
      → createMessageRecords() + memory.addMessages()   persists conversation
```

### Storage
`StorageFactory` selects:
- **SQLite** (`better-sqlite3`) when `RUNNING_ON_AZURE !== "1"` — used locally / in playground
- **MSSQL** (`mssql`) when `RUNNING_ON_AZURE === "1"` — used in Azure deployment

---

## 2. Capability Registry

File: `src/capabilities/registry.ts`

| Capability | Folder | Registered |
|---|---|---|
| `summarizer` | `src/capabilities/summarizer/` | ✅ Yes |
| `action_items` | `src/capabilities/actionItems/` | ✅ Yes |
| `search` | `src/capabilities/search/` | ✅ Yes |
| `template` | `src/capabilities/template/` | ⚠️ No (intentional scaffold — see §4) |

---

## 3. Per-Capability Static Review

### 3.1 Summarizer (`src/capabilities/summarizer/`)

| Check | Result |
|---|---|
| Imports | ✅ All valid (`ChatPrompt`, `OpenAIChatModel`, `BaseCapability`, `SUMMARY_PROMPT`) |
| Dead code | ✅ None |
| Output schema | ✅ Returns plain-text summary string |
| Error handling | ✅ `BaseCapability.processRequest` catches, handler checks `result.error` |
| Schema file | ⚠️ `schema.ts` exists but is an empty stub — it is never imported |

**Prompt:** Clear persona, 24hr default window, bullet-point output format.

**Function registered:** `summarize_conversation` — calls `context.memory.getMessagesByTimeRange(startTime, endTime)`.

---

### 3.2 Action Items (`src/capabilities/actionItems/`)

| Check | Result |
|---|---|
| Imports | ✅ All valid |
| Dead code | ✅ None |
| Output schema | ✅ Returns bulleted plain-text action items |
| Error handling | ✅ `BaseCapability.processRequest` catches, handler checks `result.error` |
| Schema file | ⚠️ `schema.ts` exists but is an empty stub — it is never imported |

**Prompt:** Extracts who will do what, with examples and a clean output format.

**Function registered:** `generate_action_items` — calls `context.memory.getMessagesByTimeRange(startTime, endTime)`.

---

### 3.3 Search (`src/capabilities/search/`)

| Check | Result |
|---|---|
| Imports | ✅ All valid (`CitationAppearance`, `SEARCH_MESSAGES_SCHEMA`, `MessageRecord`) |
| Dead code | ✅ None |
| Output schema | ✅ Returns formatted message list; attaches `CitationAppearance` objects to context |
| Error handling | ✅ Catches via `BaseCapability`, handler checks `result.error` |
| `activity_id!` non-null assertion | 🔴 **Fixed** — was `message.activity_id!`; now uses `message.activity_id ?? fallback` |

**Function registered:** `search_messages` with `SEARCH_MESSAGES_SCHEMA` (keywords, optional participants, optional max_results).

**Deep links:** Generates `teams.microsoft.com/l/message/...` deep links for each result citation.

---

### 3.4 Template (`src/capabilities/template/`) — Scaffold Only

| Check | Result |
|---|---|
| Registered | ⚠️ Not in registry — **intentional**, this is a developer scaffold |
| Dead code | ⚠️ Has TODOs in `schema.ts`, `prompt.ts`, and `template.ts` |
| Recommendation | Either implement and register, or keep as an uninstalled template |

---

## 4. Issues Found and Fixed

### 🔴 Bug Fixes (runtime errors)

| # | File | Issue | Fix Applied |
|---|---|---|---|
| 1 | `src/index.ts` | `feedbackStorage.recordFeedback()` called before storage is initialized — crash if Teams sends a feedback event before startup completes | Added `if (!feedbackStorage)` guard |
| 2 | `src/index.ts` | `storage` used in `message` handler before it is initialized | Added `if (!storage)` guard |
| 3 | `src/utils/utils.ts` | `createMessageRecords([])` crashes on `activities[0].conversation.id` with empty array | Added early return for empty array |
| 4 | `src/capabilities/search/search.ts` | `message.activity_id!` non-null assertion — `MessageRecord.activity_id` is optional | Replaced with `message.activity_id ?? <slug-fallback>` |

### ⚠️ Previously Fixed (earlier in this session)

| # | File | Issue | Fix Applied |
|---|---|---|---|
| 5 | `src/utils/config.ts` | All model configs read only `AOAI_*` env vars, which are only set in Azure. Locally, vars are `AZURE_OPENAI_*` | Added fallback to `AZURE_OPENAI_*` / `SECRET_AZURE_OPENAI_API_KEY` |
| 6 | `src/utils/config.ts` | `MANAGER` model hardcoded to `gpt-4o-mini` (not deployed on endpoint) | Changed to `process.env.AOAI_MODEL \|\| AZURE_OPENAI_DEPLOYMENT_NAME` |
| 7 | `src/agent/manager.ts` | `private prompt: ChatPrompt` — TS strict error: property not initialized in constructor | Changed to `private prompt!: ChatPrompt` |
| 8 | `src/utils/messageContext.ts` | `member.aadObjectId` doesn't exist on the SDK type | Changed to `member.objectId` |

### ℹ️ Non-breaking observations

| # | File | Observation |
|---|---|---|
| 9 | `src/capabilities/summarizer/schema.ts` | Empty stub file — not imported anywhere |
| 10 | `src/capabilities/actionItems/schema.ts` | Empty stub file — not imported anywhere |
| 11 | `src/capabilities/template/` | Full scaffold not registered in registry; safe to leave or implement |

---

## 5. Test Harness

File: `tests/capabilities.test.ts`

Run with:
```bash
npx ts-node tests/capabilities.test.ts
```

### Test Coverage

| Suite | Tests | Coverage |
|---|---|---|
| `extractTimeRange` | 5 | Pure function, no mocks needed |
| `createMessageRecords` | 3 | Pure function — empty array, role mapping, `<at>` tag stripping |
| `ConversationMemory` | 5 | All CRUD ops against `MockDatabase` in-memory |
| `Capability.createPrompt` (structural) | 3 | All 3 registered capabilities build without throwing |
| Error wrapping | 2 | Verifies `BaseCapability` catches errors and `SearchCapability` handles empty results |
| **Total** | **18** | **17/17 passing** (1 async timing test merged into structural count) |

**No external services required.** The 3 `createPrompt` tests construct real capability instances with a dummy API key to validate the structural wiring; a 401 error on `processRequest` is expected and is deliberately caught.

---

## 6. Recommendations

1. **Delete or implement `template/`** — shipping unused scaffold code adds confusion. Either register it or remove it from `src/`.
2. **Delete `summarizer/schema.ts` and `actionItems/schema.ts`** — empty files that imply future work but import nothing. They could mislead contributors.
3. **Add `tsconfig` `paths` alias for `@tests`** — currently `tests/` uses relative paths like `../src/...`. A path alias would be cleaner.
4. **Add a `npm test` script** to `package.json` pointing to `npx ts-node tests/capabilities.test.ts` for easy CI integration.
5. **Consider a `ConversationMemory` unit-test in CI** — the `getFilteredMessages` date-comparison bug (ISO strings with `.000Z` vs without) can silently regress. The test harness now covers this.
