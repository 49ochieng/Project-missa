
# Project Missa: Teams AI Meeting Note Taker
## Complete Build Instructions for Claude Code

---

## 🎯 MISSION

You are building **Project Missa** — a Microsoft Teams AI Meeting Note Taker that:
1. Joins Teams meetings **as a bot participant** (appears in the meeting roster)
2. Listens to the meeting and captures a full transcript via Teams native transcription
3. After the meeting ends, uses Azure OpenAI to generate a structured summary
4. Emails the summary + full transcript to **all meeting participants** via Microsoft Graph API
5. Posts a summary card to the meeting chat via Microsoft Graph API
6. Azure Speech Service is available as fallback if Teams native transcription is unavailable

This is built on top of the existing **Project Missa** codebase which already has:
- Teams AI SDK v2 (`@microsoft/teams.apps`, `@microsoft/teams.ai`, etc.)
- A working Missa Agent with Summarizer, Action Items, and Search capabilities
- SQLite + MSSQL dual-storage with a clean `IDatabase` interface
- Azure Bicep infrastructure (App Service + Azure SQL + Bot Registration)
- Teams Toolkit / M365 Agents Toolkit project structure

**You are EXTENDING this codebase** — not replacing it. The existing capabilities stay. You are adding a new `meetingBot` module.

---

## 📁 REPOSITORY STRUCTURE (Existing + What You Will Add)

```
Project-missa/
├── src/
│   ├── index.ts                          ← MODIFY: add meeting bot startup + routes
│   ├── agent/                            ← EXISTING: keep untouched
│   ├── capabilities/                     ← EXISTING: keep untouched
│   ├── storage/                          ← EXISTING: keep untouched
│   ├── utils/                            ← EXISTING: keep untouched
│   │
│   └── meetingBot/                       ← CREATE THIS ENTIRE MODULE
│       ├── index.ts                      ← Meeting bot orchestrator entry point
│       ├── graphClient.ts                ← Lightweight Graph API client (MSAL app-only)
│       ├── callHandler.ts                ← Graph Communications API call lifecycle
│       ├── transcriptionService.ts       ← Azure Speech SDK (fallback) + VTT parser
│       ├── summarizationService.ts       ← Azure OpenAI meeting summary generation
│       ├── notificationService.ts        ← Email + Teams chat delivery via Graph
│       ├── calendarWatcher.ts            ← Graph calendar polling for upcoming meetings
│       ├── meetingStore.ts               ← In-memory meeting session state
│       └── types.ts                      ← Meeting bot type definitions
│
├── appPackage/
│   └── manifest.json                     ← MODIFY: add supportsCalling + RSC permissions
├── infra/
│   └── azure.bicep                       ← MODIFY: add Azure Speech Service resource
├── env/
│   └── .env.local.user (gitignored)      ← ADD: new env vars documented below
└── package.json                          ← MODIFY: add new npm dependencies
```

---

## 🔐 STEP 1 — ENVIRONMENT VARIABLES

### New variables to add to `.env.*.user` (gitignored):

```env
# Azure Speech Service (create in Azure Portal → Cognitive Services → Speech)
AZURE_SPEECH_KEY=<your-speech-subscription-key>
AZURE_SPEECH_REGION=eastus

# Meeting bot display name (how bot appears in meeting roster)
MEETING_BOT_DISPLAY_NAME=Meeting Notes Bot

# Email sender UPN — must be a real user/mailbox in your tenant
# The bot app registration will send emails on behalf of this account
# Requires Mail.Send application permission granted with admin consent
EMAIL_SENDER_UPN=meetingbot@<yourtenant>.onmicrosoft.com
```

### After `teamsapp provision`, verify `.localConfigs` contains:
```env
CLIENT_ID=<BOT_ID>          # Bot's AAD app client ID
CLIENT_SECRET=<SECRET>       # Bot's AAD app client secret
TENANT_ID=<TENANT_ID>        # Your M365 tenant ID
AOAI_ENDPOINT=<endpoint>     # Already exists
AOAI_API_KEY=<key>           # Already exists
AOAI_MODEL=<deployment>      # Already exists
AZURE_SPEECH_KEY=<key>       # NEW
AZURE_SPEECH_REGION=eastus   # NEW
MEETING_BOT_DISPLAY_NAME=Meeting Notes Bot  # NEW
EMAIL_SENDER_UPN=meetingbot@<tenant>.onmicrosoft.com  # NEW
```

---

## 🔑 STEP 2 — AZURE AD APP REGISTRATION PERMISSIONS

The app registration that Teams Toolkit created needs additional Graph API permissions. In Azure Portal → App Registrations → your bot app → API Permissions → Add permission → Microsoft Graph → Application permissions, add ALL of the following and grant Admin Consent:

| Permission | Purpose | Required |
|---|---|---|
| `Calls.JoinGroupCall.All` | Bot joins group meetings | ✅ Critical |
| `Calls.JoinGroupCallAsGuest.All` | Join as guest (cross-tenant) | ✅ Critical |
| `Calls.InitiateGroupCall.All` | Initiate calls | ✅ Critical |
| `OnlineMeetings.Read.All` | Read meeting metadata | ✅ Critical |
| `OnlineMeetings.ReadWrite.All` | Enable transcription | ✅ Critical |
| `OnlineMeetingTranscript.Read.All` | Fetch transcript content | ✅ Critical |
| `Calendars.Read` | Watch calendar for meetings | ✅ Critical |
| `User.Read.All` | Resolve participant emails | ✅ Critical |
| `Mail.Send` | Send summary emails | ✅ Critical |
| `Chat.ReadWrite.All` | Post to meeting chat | ✅ Critical |
| `ChatMessage.Read.Chat` | Already exists — keep | ✅ Existing |

⚠️ MANDATORY: After adding all permissions, click **"Grant admin consent for [Your Tenant]"**. Every permission must show a green ✅ checkmark. Without admin consent, all Graph API calls will return 403 Forbidden.

### Why NOT Calls.AccessMedia.All (for raw audio):
The Graph Real-time Media Platform (RMP SDK) for raw audio frames is **.NET-only** and requires Windows Server hosting. This project runs Node.js on Linux App Service. The architecture uses:
- Bot joins via Graph Communications API → call established
- Teams native transcription enabled via PATCH /onlineMeetings
- Graph change notifications fire when transcript is ready after meeting ends
- Transcript fetched as VTT file → parsed → summarized → emailed

This is 100% Node.js compatible, production-grade, and officially supported.

---

## 📦 STEP 3 — PACKAGE.JSON ADDITIONS

Add to the `dependencies` section in `package.json`:
```json
"microsoft-cognitiveservices-speech-sdk": "^1.43.0",
"@azure/msal-node": "^2.16.2",
"uuid": "^11.0.5"
```

Add to `devDependencies`:
```json
"@types/uuid": "^10.0.0"
```

Run `npm install` after updating package.json.

---

## 🏗️ STEP 4 — BUILD THE MEETING BOT MODULE

Create each file in `src/meetingBot/` as follows. Create them in this exact order to avoid TypeScript import errors.

### 4.1 — CREATE `src/meetingBot/types.ts`

```typescript
export interface MeetingSession {
  callId: string;
  meetingId: string;
  organizerUserId: string;
  tenantId: string;
  joinWebUrl: string;
  chatId?: string;
  title: string;
  scheduledStart: string;
  actualStart?: string;
  actualEnd?: string;
  state: MeetingState;
  participants: MeetingParticipant[];
  subscriptionId?: string;
}

export type MeetingState =
  | 'scheduled'
  | 'joining'
  | 'in_call'
  | 'call_ended'
  | 'transcript_pending'
  | 'transcript_ready'
  | 'summarizing'
  | 'completed'
  | 'failed';

export interface MeetingParticipant {
  userId: string;
  displayName: string;
  email?: string;
  joinedAt?: string;
}

export interface MeetingSummary {
  executiveSummary: string;
  keyDecisions: Array<{
    decision: string;
    owner?: string;
    context?: string;
  }>;
  actionItems: Array<{
    task: string;
    owner?: string;
    dueDate?: string;
    priority: 'high' | 'medium' | 'low';
  }>;
  discussionTopics: Array<{
    topic: string;
    summary: string;
    participants: string[];
  }>;
  nextSteps: string[];
  meetingStats: {
    durationMinutes: number;
    participantCount: number;
    totalSpeakers: number;
  };
}
```

---

### 4.2 — CREATE `src/meetingBot/meetingStore.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";
import { MeetingSession, MeetingState } from "./types";

export class MeetingStore {
  private sessions = new Map<string, MeetingSession>();
  private byMeetingId = new Map<string, string>();

  constructor(private logger: ILogger) {}

  create(session: MeetingSession): void {
    this.sessions.set(session.callId, session);
    this.byMeetingId.set(session.meetingId, session.callId);
    this.logger.debug(`📋 Session created: ${session.callId} (${session.title})`);
  }

  getByCallId(callId: string): MeetingSession | undefined {
    return this.sessions.get(callId);
  }

  getByMeetingId(meetingId: string): MeetingSession | undefined {
    const callId = this.byMeetingId.get(meetingId);
    return callId ? this.sessions.get(callId) : undefined;
  }

  getBySubscriptionId(subscriptionId: string): MeetingSession | undefined {
    for (const session of this.sessions.values()) {
      if (session.subscriptionId === subscriptionId) return session;
    }
    return undefined;
  }

  update(callId: string, updates: Partial<MeetingSession>): void {
    const session = this.sessions.get(callId);
    if (session) {
      Object.assign(session, updates);
      if (updates.state) {
        this.logger.debug(`📝 Session ${callId} → ${updates.state}`);
      }
    }
  }

  setState(callId: string, state: MeetingState): void {
    this.update(callId, { state });
  }

  delete(callId: string): void {
    const session = this.sessions.get(callId);
    if (session) {
      this.byMeetingId.delete(session.meetingId);
      this.sessions.delete(callId);
    }
  }

  getAll(): MeetingSession[] {
    return Array.from(this.sessions.values());
  }

  size(): number {
    return this.sessions.size;
  }
}
```

---

### 4.3 — CREATE `src/meetingBot/graphClient.ts`

```typescript
import { ConfidentialClientApplication } from "@azure/msal-node";
import { ILogger } from "@microsoft/teams.common";

interface GraphAuthToken {
  accessToken: string;
  expiresAt: number;
}

export class MeetingBotGraphClient {
  private msalClient: ConfidentialClientApplication;
  private cachedToken: GraphAuthToken | null = null;

  constructor(private logger: ILogger) {
    this.msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: process.env.CLIENT_ID!,
        clientSecret: process.env.CLIENT_SECRET!,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
      },
    });
  }

  async getAccessToken(): Promise<string> {
    if (this.cachedToken && this.cachedToken.expiresAt > Date.now() + 5 * 60 * 1000) {
      return this.cachedToken.accessToken;
    }
    const result = await this.msalClient.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    if (!result?.accessToken) {
      throw new Error("Failed to acquire Graph API token");
    }
    this.cachedToken = {
      accessToken: result.accessToken,
      expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
    };
    return this.cachedToken.accessToken;
  }

  async get<T>(endpoint: string): Promise<T> {
    return this.request<T>("GET", endpoint);
  }

  async post<T>(endpoint: string, body?: unknown): Promise<T> {
    return this.request<T>("POST", endpoint, body);
  }

  async patch<T>(endpoint: string, body: unknown): Promise<T> {
    return this.request<T>("PATCH", endpoint, body);
  }

  async delete(endpoint: string): Promise<void> {
    await this.request("DELETE", endpoint);
  }

  async getText(endpoint: string): Promise<string> {
    const token = await this.getAccessToken();
    const url = endpoint.startsWith("https://")
      ? endpoint
      : `https://graph.microsoft.com/v1.0${endpoint}`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: "text/vtt, text/plain, */*" },
    });
    if (!response.ok) {
      throw new Error(`Graph GET text ${endpoint} → ${response.status}`);
    }
    return response.text();
  }

  private async request<T>(method: string, endpoint: string, body?: unknown): Promise<T> {
    const token = await this.getAccessToken();
    const url = endpoint.startsWith("https://")
      ? endpoint
      : `https://graph.microsoft.com/v1.0${endpoint}`;
    const response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!response.ok) {
      const errorText = await response.text();
      this.logger.error(`❌ Graph ${method} ${endpoint} → ${response.status}: ${errorText}`);
      throw new Error(`Graph API error ${response.status}: ${errorText}`);
    }
    if (response.status === 204) return undefined as T;
    return response.json() as Promise<T>;
  }
}
```

---

### 4.4 — CREATE `src/meetingBot/callHandler.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";
import { MeetingBotGraphClient } from "./graphClient";
import { MeetingStore } from "./meetingStore";
import { MeetingParticipant, MeetingSession } from "./types";

export class CallHandler {
  constructor(
    private graph: MeetingBotGraphClient,
    public store: MeetingStore,
    private logger: ILogger
  ) {}

  async joinMeeting(
    joinWebUrl: string,
    meetingId: string,
    organizerUserId: string,
    tenantId: string,
    title: string
  ): Promise<string> {
    this.logger.debug(`📞 Joining meeting: ${title}`);

    // serviceHostedMediaConfig: Microsoft handles media, we get call events
    // We use this (not appHostedMediaConfig) because we are Node.js on Linux
    const callBody = {
      "@odata.type": "#microsoft.graph.call",
      callbackUri: `${process.env.BOT_ENDPOINT}/api/calls`,
      requestedModalities: ["audio"],
      mediaConfig: {
        "@odata.type": "#microsoft.graph.serviceHostedMediaConfig",
      },
      meetingInfo: {
        "@odata.type": "#microsoft.graph.organizerMeetingInfo",
        organizer: {
          "@odata.type": "#microsoft.graph.identitySet",
          user: {
            "@odata.type": "#microsoft.graph.identity",
            id: organizerUserId,
            tenantId,
          },
        },
        allowConversationWithoutHost: true,
      },
      tenantId,
      subject: process.env.MEETING_BOT_DISPLAY_NAME || "Meeting Notes Bot",
    };

    const response = await this.graph.post<{ id: string }>("/communications/calls", callBody);
    const callId = response.id;

    this.logger.debug(`✅ Bot joined. Call ID: ${callId}`);

    const session: MeetingSession = {
      callId,
      meetingId,
      organizerUserId,
      tenantId,
      joinWebUrl,
      title,
      scheduledStart: new Date().toISOString(),
      state: "joining",
      participants: [],
    };
    this.store.create(session);
    return callId;
  }

  async enableTranscription(meetingId: string, organizerUserId: string): Promise<void> {
    try {
      await this.graph.patch(`/users/${organizerUserId}/onlineMeetings/${meetingId}`, {
        isTranscriptionEnabled: true,
      });
      this.logger.debug(`🎙️ Transcription enabled for: ${meetingId}`);
    } catch (error) {
      this.logger.warn(`⚠️ Could not enable transcription (non-fatal): ${error}`);
    }
  }

  async subscribeToTranscriptNotifications(
    meetingId: string,
    organizerUserId: string,
    callId: string
  ): Promise<string> {
    const expiry = new Date(Date.now() + 4 * 60 * 60 * 1000).toISOString();
    const subscription = await this.graph.post<{ id: string }>("/subscriptions", {
      changeType: "created",
      notificationUrl: `${process.env.BOT_ENDPOINT}/api/notifications`,
      resource: `/users/${organizerUserId}/onlineMeetings/${meetingId}/transcripts`,
      expirationDateTime: expiry,
      clientState: callId,
    });
    this.logger.debug(`📡 Subscribed to transcripts. Sub ID: ${subscription.id}`);
    return subscription.id;
  }

  async handleCallWebhook(body: {
    value?: Array<{ resource?: string; resourceData?: { state?: string } }>;
  }): Promise<void> {
    for (const notification of body.value || []) {
      const resourceUrl = notification.resource || "";
      const callId = resourceUrl.match(/\/communications\/calls\/([^/]+)/)?.[1];
      if (!callId) continue;

      const state = notification.resourceData?.state || "";
      this.logger.debug(`📞 Call event: ${callId} → ${state}`);

      if (state === "establishing") {
        this.store.setState(callId, "joining");
      } else if (state === "established") {
        await this.onCallEstablished(callId);
      } else if (state === "terminated") {
        await this.onCallTerminated(callId);
      }
    }
  }

  private async onCallEstablished(callId: string): Promise<void> {
    const session = this.store.getByCallId(callId);
    if (!session) return;

    this.store.update(callId, { state: "in_call", actualStart: new Date().toISOString() });
    this.logger.debug(`✅ Call established: ${session.title}`);

    await this.enableTranscription(session.meetingId, session.organizerUserId);

    try {
      const subId = await this.subscribeToTranscriptNotifications(
        session.meetingId,
        session.organizerUserId,
        callId
      );
      this.store.update(callId, { subscriptionId: subId });
    } catch (err) {
      this.logger.warn(`⚠️ Transcript subscription failed: ${err}`);
    }

    await this.refreshParticipants(callId, session.meetingId, session.organizerUserId);
  }

  private async onCallTerminated(callId: string): Promise<void> {
    const session = this.store.getByCallId(callId);
    if (!session) return;

    this.store.update(callId, { state: "call_ended", actualEnd: new Date().toISOString() });
    this.logger.debug(`📴 Call ended: ${session.title}`);

    if (session.subscriptionId) {
      try {
        await this.graph.delete(`/subscriptions/${session.subscriptionId}`);
      } catch { /* non-fatal */ }
    }

    this.store.setState(callId, "transcript_pending");
  }

  private async refreshParticipants(
    callId: string,
    meetingId: string,
    organizerUserId: string
  ): Promise<void> {
    try {
      const meetingData = await this.graph.get<{
        participants?: {
          attendees?: Array<{ identity?: { user?: { id?: string; displayName?: string } } }>;
          organizer?: { identity?: { user?: { id?: string; displayName?: string } } };
        };
      }>(`/users/${organizerUserId}/onlineMeetings/${meetingId}`);

      const participants: MeetingParticipant[] = [];
      const organizer = meetingData.participants?.organizer?.identity?.user;
      if (organizer?.id) {
        participants.push({
          userId: organizer.id,
          displayName: organizer.displayName || "Organizer",
          email: await this.resolveEmail(organizer.id),
        });
      }
      for (const attendee of meetingData.participants?.attendees || []) {
        const user = attendee.identity?.user;
        if (user?.id && !participants.find((p) => p.userId === user.id)) {
          participants.push({
            userId: user.id!,
            displayName: user.displayName || "Attendee",
            email: await this.resolveEmail(user.id!),
          });
        }
      }
      this.store.update(callId, { participants });
      this.logger.debug(`👥 ${participants.length} participants for ${callId}`);
    } catch (error) {
      this.logger.warn(`⚠️ Could not refresh participants: ${error}`);
    }
  }

  async resolveEmail(userId: string): Promise<string | undefined> {
    try {
      const user = await this.graph.get<{ mail?: string; userPrincipalName?: string }>(
        `/users/${userId}?$select=mail,userPrincipalName`
      );
      return user.mail || user.userPrincipalName;
    } catch {
      return undefined;
    }
  }

  async fetchTranscript(
    meetingId: string,
    organizerUserId: string
  ): Promise<Array<{ speakerName: string; text: string; timestamp: string }>> {
    const transcriptsResponse = await this.graph.get<{
      value: Array<{ id: string; createdDateTime: string }>;
    }>(`/users/${organizerUserId}/onlineMeetings/${meetingId}/transcripts`);

    if (!transcriptsResponse.value?.length) {
      this.logger.warn(`⚠️ No transcripts for meeting: ${meetingId}`);
      return [];
    }

    const latest = transcriptsResponse.value.sort(
      (a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime()
    )[0];

    this.logger.debug(`📄 Fetching transcript: ${latest.id}`);

    const vttContent = await this.graph.getText(
      `/users/${organizerUserId}/onlineMeetings/${meetingId}/transcripts/${latest.id}/content?$format=text/vtt`
    );

    return parseVtt(vttContent);
  }

  async getMeetingChatId(meetingId: string, organizerUserId: string): Promise<string | undefined> {
    try {
      const meeting = await this.graph.get<{ chatInfo?: { threadId?: string } }>(
        `/users/${organizerUserId}/onlineMeetings/${meetingId}?$select=chatInfo`
      );
      return meeting.chatInfo?.threadId;
    } catch {
      return undefined;
    }
  }

  async resolveMeetingId(userId: string, joinUrl: string): Promise<string | null> {
    try {
      const encoded = encodeURIComponent(joinUrl);
      const result = await this.graph.get<{ value?: Array<{ id: string }> }>(
        `/users/${userId}/onlineMeetings?$filter=JoinWebUrl eq '${encoded}'`
      );
      return result.value?.[0]?.id || null;
    } catch {
      return null;
    }
  }
}

function parseVtt(
  vtt: string
): Array<{ speakerName: string; text: string; timestamp: string }> {
  const lines = vtt.split("\n");
  const result: Array<{ speakerName: string; text: string; timestamp: string }> = [];
  let currentTimestamp = "";

  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.includes("-->")) {
      currentTimestamp = trimmed.split("-->")[0].trim();
      continue;
    }
    const speakerMatch = trimmed.match(/^<v ([^>]+)>(.+)$/);
    if (speakerMatch && currentTimestamp) {
      result.push({
        speakerName: speakerMatch[1].trim(),
        text: speakerMatch[2].trim(),
        timestamp: currentTimestamp,
      });
    }
  }
  return result;
}
```

---

### 4.5 — CREATE `src/meetingBot/transcriptionService.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";

/**
 * Azure Speech Service transcription.
 * PRIMARY: Teams native transcription (via Graph API) is preferred.
 * FALLBACK: Azure Speech SDK used when Teams transcription is unavailable.
 * UTILITY: formatTranscript used by all paths to prepare text for OpenAI.
 */
export class TranscriptionService {
  constructor(private logger: ILogger) {}

  static formatTranscript(
    lines: Array<{ speakerName: string; text: string; timestamp: string }>
  ): string {
    if (!lines.length) return "";
    return lines
      .map((line) => `[${line.timestamp}] ${line.speakerName}: ${line.text}`)
      .join("\n");
  }

  validate(): boolean {
    if (!process.env.AZURE_SPEECH_KEY || !process.env.AZURE_SPEECH_REGION) {
      this.logger.warn("⚠️ AZURE_SPEECH_KEY/REGION not set — Speech SDK fallback unavailable");
      return false;
    }
    return true;
  }
}
```

---

### 4.6 — CREATE `src/meetingBot/summarizationService.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";
import { MeetingSession, MeetingSummary } from "./types";

export class SummarizationService {
  constructor(private logger: ILogger) {}

  async summarize(transcript: string, session: MeetingSession): Promise<MeetingSummary> {
    this.logger.debug(`🤖 Summarizing: ${session.title}`);

    const participantNames = session.participants
      .map((p) => p.displayName)
      .filter(Boolean)
      .join(", ");

    const durationMinutes =
      session.actualStart && session.actualEnd
        ? Math.round(
            (new Date(session.actualEnd).getTime() - new Date(session.actualStart).getTime()) /
              60000
          )
        : 0;

    const prompt = `You are an expert meeting analyst. Analyze this Microsoft Teams meeting transcript and produce a structured JSON summary.

MEETING DETAILS:
Title: ${session.title}
Participants: ${participantNames || "Unknown"}
Duration: ${durationMinutes} minutes

TRANSCRIPT:
${transcript}

TASK: Return ONLY valid JSON (no markdown, no code fences, no explanation) matching this exact structure:
{
  "executiveSummary": "2-3 sentence overview of what was accomplished",
  "keyDecisions": [
    { "decision": "specific decision", "owner": "person or null", "context": "why" }
  ],
  "actionItems": [
    { "task": "specific task", "owner": "person or null", "dueDate": "date or null", "priority": "high|medium|low" }
  ],
  "discussionTopics": [
    { "topic": "topic name", "summary": "what was discussed", "participants": ["names"] }
  ],
  "nextSteps": ["next step as string"],
  "meetingStats": {
    "durationMinutes": ${durationMinutes},
    "participantCount": ${session.participants.length},
    "totalSpeakers": 0
  }
}`;

    const response = await this.callAzureOpenAI(prompt);
    return this.parseSummary(response, session, durationMinutes);
  }

  private async callAzureOpenAI(prompt: string): Promise<string> {
    const endpoint = process.env.AOAI_ENDPOINT!.replace(/\/$/, "");
    const model = process.env.AOAI_MODEL || "gpt-4o";
    const apiKey = process.env.AOAI_API_KEY!;
    const url = `${endpoint}/openai/deployments/${model}/chat/completions?api-version=2025-04-01-preview`;

    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "api-key": apiKey },
      body: JSON.stringify({
        messages: [
          { role: "system", content: "You are a precise meeting analyst. Return only valid JSON." },
          { role: "user", content: prompt },
        ],
        temperature: 0.1,
        max_tokens: 2000,
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Azure OpenAI error ${response.status}: ${err}`);
    }

    const data = (await response.json()) as {
      choices: Array<{ message: { content: string } }>;
    };
    return data.choices[0]?.message?.content || "{}";
  }

  private parseSummary(
    raw: string,
    session: MeetingSession,
    durationMinutes: number
  ): MeetingSummary {
    try {
      const cleaned = raw.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
      const parsed = JSON.parse(cleaned) as MeetingSummary;
      return {
        executiveSummary: parsed.executiveSummary || "Summary not available.",
        keyDecisions: parsed.keyDecisions || [],
        actionItems: parsed.actionItems || [],
        discussionTopics: parsed.discussionTopics || [],
        nextSteps: parsed.nextSteps || [],
        meetingStats: {
          durationMinutes,
          participantCount: session.participants.length,
          totalSpeakers: parsed.meetingStats?.totalSpeakers || 0,
        },
      };
    } catch {
      return {
        executiveSummary: raw.slice(0, 500),
        keyDecisions: [],
        actionItems: [],
        discussionTopics: [],
        nextSteps: [],
        meetingStats: { durationMinutes, participantCount: session.participants.length, totalSpeakers: 0 },
      };
    }
  }
}
```

---

### 4.7 — CREATE `src/meetingBot/notificationService.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";
import { MeetingBotGraphClient } from "./graphClient";
import { MeetingSession, MeetingSummary } from "./types";

export class NotificationService {
  constructor(private graph: MeetingBotGraphClient, private logger: ILogger) {}

  async sendSummaryEmail(
    session: MeetingSession,
    summary: MeetingSummary,
    transcript: string
  ): Promise<void> {
    const recipientEmails = session.participants
      .map((p) => p.email)
      .filter((e): e is string => Boolean(e));

    if (recipientEmails.length === 0) {
      this.logger.warn(`⚠️ No recipient emails for: ${session.callId}`);
      return;
    }

    const senderUpn = process.env.EMAIL_SENDER_UPN;
    if (!senderUpn) {
      this.logger.error("❌ EMAIL_SENDER_UPN not configured");
      return;
    }

    const date = session.actualStart
      ? new Date(session.actualStart).toLocaleString("en-US", { dateStyle: "full", timeStyle: "short" })
      : "Unknown date";

    const htmlBody = buildEmailHtml(session, summary, transcript, date);

    await this.graph.post(`/users/${senderUpn}/sendMail`, {
      message: {
        subject: `📋 Meeting Summary: ${session.title}`,
        body: { contentType: "HTML", content: htmlBody },
        toRecipients: recipientEmails.map((email) => ({ emailAddress: { address: email } })),
        importance: "normal",
      },
      saveToSentItems: true,
    });

    this.logger.debug(`📧 Email sent to ${recipientEmails.length} recipients`);
  }

  async postSummaryToMeetingChat(
    chatId: string,
    session: MeetingSession,
    summary: MeetingSummary
  ): Promise<void> {
    const actionItemsList =
      summary.actionItems.length > 0
        ? summary.actionItems
            .map((a) => `• ${a.task}${a.owner ? ` → ${a.owner}` : ""}`)
            .join("<br>")
        : "No action items identified";

    const content = `<strong>📋 Meeting Summary: ${escHtml(session.title)}</strong><br><br>
<strong>Overview:</strong> ${escHtml(summary.executiveSummary)}<br><br>
<strong>✅ Action Items:</strong><br>${actionItemsList}<br><br>
<strong>📊 Stats:</strong> ${summary.meetingStats.durationMinutes} min · ${summary.meetingStats.participantCount} participants<br><br>
<em>Full summary has been emailed to all participants.</em>`;

    await this.graph.post(`/chats/${chatId}/messages`, {
      body: { contentType: "html", content },
    });

    this.logger.debug(`💬 Summary card posted to meeting chat: ${chatId}`);
  }
}

function buildEmailHtml(
  session: MeetingSession,
  summary: MeetingSummary,
  transcript: string,
  date: string
): string {
  const participants = session.participants
    .map((p) => `<li>${escHtml(p.displayName)}${p.email ? ` (${p.email})` : ""}</li>`)
    .join("");

  const actionItemsHtml =
    summary.actionItems.length > 0
      ? `<table style="width:100%;border-collapse:collapse;font-size:14px;">
          <thead><tr style="background:#f0f0f0;">
            <th style="padding:8px;text-align:left;">Task</th>
            <th style="padding:8px;text-align:left;">Owner</th>
            <th style="padding:8px;text-align:left;">Due</th>
            <th style="padding:8px;text-align:left;">Priority</th>
          </tr></thead><tbody>
          ${summary.actionItems
            .map(
              (a) =>
                `<tr><td style="padding:8px;border-bottom:1px solid #eee;">${escHtml(a.task)}</td>
                <td style="padding:8px;border-bottom:1px solid #eee;">${a.owner || "—"}</td>
                <td style="padding:8px;border-bottom:1px solid #eee;">${a.dueDate || "—"}</td>
                <td style="padding:8px;border-bottom:1px solid #eee;">
                  <span style="background:${priorityColor(a.priority)};color:white;padding:2px 8px;border-radius:10px;font-size:11px;">${a.priority}</span>
                </td></tr>`
            )
            .join("")}
          </tbody></table>`
      : `<p style="color:#666;">No action items identified</p>`;

  const decisionsHtml =
    summary.keyDecisions.length > 0
      ? `<ul>${summary.keyDecisions
          .map(
            (d) =>
              `<li style="margin-bottom:8px;"><strong>${escHtml(d.decision)}</strong>
              ${d.owner ? `<span style="color:#666;"> — ${d.owner}</span>` : ""}
              ${d.context ? `<br><em style="color:#888;font-size:13px;">${escHtml(d.context)}</em>` : ""}</li>`
          )
          .join("")}</ul>`
      : `<p style="color:#666;">No key decisions recorded</p>`;

  const nextStepsHtml =
    summary.nextSteps.length > 0
      ? `<ul>${summary.nextSteps.map((s) => `<li>${escHtml(s)}</li>`).join("")}</ul>`
      : `<p style="color:#666;">No next steps recorded</p>`;

  const transcriptSection = transcript
    ? `<details style="margin-top:20px;">
        <summary style="cursor:pointer;color:#0078d4;font-weight:600;font-size:14px;">📄 View Full Transcript</summary>
        <pre style="background:#f5f5f5;padding:16px;border-radius:8px;margin-top:12px;white-space:pre-wrap;font-size:11px;line-height:1.6;max-height:400px;overflow-y:auto;">${escHtml(transcript)}</pre>
      </details>`
    : "";

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:700px;margin:0 auto;color:#333;">
  <div style="background:linear-gradient(135deg,#0078d4,#106ebe);color:white;padding:24px;border-radius:8px 8px 0 0;">
    <h1 style="margin:0;font-size:22px;">📋 Meeting Summary</h1>
    <p style="margin:8px 0 0;opacity:.9;font-size:16px;">${escHtml(session.title)}</p>
    <p style="margin:4px 0 0;opacity:.8;font-size:13px;">${date} · ${summary.meetingStats.durationMinutes} min · ${summary.meetingStats.participantCount} participants</p>
  </div>
  <div style="background:white;padding:24px;border:1px solid #e0e0e0;border-top:none;">
    <h2 style="color:#0078d4;font-size:16px;margin-top:0;">🔍 Executive Summary</h2>
    <p style="line-height:1.6;background:#f8f9fa;padding:16px;border-left:4px solid #0078d4;border-radius:0 8px 8px 0;">${escHtml(summary.executiveSummary)}</p>
    <h2 style="color:#0078d4;font-size:16px;">✅ Action Items</h2>
    ${actionItemsHtml}
    <h2 style="color:#0078d4;font-size:16px;margin-top:24px;">🏛️ Key Decisions</h2>
    ${decisionsHtml}
    <h2 style="color:#0078d4;font-size:16px;">🚀 Next Steps</h2>
    ${nextStepsHtml}
    <h2 style="color:#0078d4;font-size:16px;">👥 Participants (${session.participants.length})</h2>
    <ul>${participants}</ul>
    ${transcriptSection}
  </div>
  <div style="background:#f5f5f5;padding:12px 24px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;font-size:12px;color:#888;text-align:center;">
    Auto-generated by Meeting Notes Bot · Project Missa
  </div>
</body></html>`;
}

function priorityColor(p: string): string {
  return p === "high" ? "#d13438" : p === "medium" ? "#f7630c" : "#0078d4";
}

function escHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
```

---

### 4.8 — CREATE `src/meetingBot/calendarWatcher.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";
import { CallHandler } from "./callHandler";
import { MeetingBotGraphClient } from "./graphClient";

export class CalendarWatcher {
  private watchedUsers = new Map<string, ReturnType<typeof setInterval>>();
  private scheduledMeetings = new Set<string>();

  constructor(
    private graph: MeetingBotGraphClient,
    private callHandler: CallHandler,
    private logger: ILogger
  ) {}

  startWatching(userId: string, tenantId: string): void {
    if (this.watchedUsers.has(userId)) {
      this.logger.debug(`Already watching: ${userId}`);
      return;
    }
    this.logger.debug(`📅 Starting calendar watch for: ${userId}`);

    const interval = setInterval(() => {
      this.checkUpcomingMeetings(userId, tenantId).catch((err) =>
        this.logger.error(`Calendar check error: ${err}`)
      );
    }, 5 * 60 * 1000);

    this.watchedUsers.set(userId, interval);
    this.checkUpcomingMeetings(userId, tenantId).catch((err) =>
      this.logger.error(`Initial calendar check error: ${err}`)
    );
  }

  stopWatching(userId: string): void {
    const interval = this.watchedUsers.get(userId);
    if (interval) {
      clearInterval(interval);
      this.watchedUsers.delete(userId);
    }
  }

  private async checkUpcomingMeetings(userId: string, tenantId: string): Promise<void> {
    const now = new Date();
    const soon = new Date(now.getTime() + 15 * 60 * 1000);
    const past = new Date(now.getTime() - 2 * 60 * 1000);

    const events = await this.graph.get<{
      value: Array<{
        id: string;
        subject: string;
        start: { dateTime: string };
        isOnlineMeeting: boolean;
        onlineMeeting?: { joinUrl: string };
      }>;
    }>(
      `/users/${userId}/calendarView?startDateTime=${past.toISOString()}&endDateTime=${soon.toISOString()}&$filter=isOnlineMeeting eq true&$select=id,subject,start,isOnlineMeeting,onlineMeeting&$top=10`
    );

    for (const event of events.value || []) {
      if (!event.onlineMeeting?.joinUrl) continue;
      if (this.scheduledMeetings.has(event.id)) continue;

      const meetingId = await this.callHandler.resolveMeetingId(userId, event.onlineMeeting.joinUrl);
      if (!meetingId) continue;

      if (this.callHandler.store.getByMeetingId(meetingId)) continue;

      this.scheduledMeetings.add(event.id);
      const startTime = new Date(event.start.dateTime).getTime();
      const joinTime = startTime - 60 * 1000;
      const delay = Math.max(0, joinTime - Date.now());

      setTimeout(async () => {
        try {
          await this.callHandler.joinMeeting(
            event.onlineMeeting!.joinUrl,
            meetingId,
            userId,
            tenantId,
            event.subject
          );
        } catch (err) {
          this.logger.error(`❌ Failed to join "${event.subject}": ${err}`);
          this.scheduledMeetings.delete(event.id);
        }
      }, delay);

      this.logger.debug(
        `⏰ Scheduled join for "${event.subject}" in ${Math.round(delay / 1000)}s`
      );
    }
  }
}
```

---

### 4.9 — CREATE `src/meetingBot/index.ts`

```typescript
import { ILogger } from "@microsoft/teams.common";
import { CallHandler } from "./callHandler";
import { CalendarWatcher } from "./calendarWatcher";
import { MeetingBotGraphClient } from "./graphClient";
import { MeetingStore } from "./meetingStore";
import { NotificationService } from "./notificationService";
import { SummarizationService } from "./summarizationService";
import { TranscriptionService } from "./transcriptionService";
import { MeetingSession } from "./types";

export type { MeetingSession };

export class MeetingBotOrchestrator {
  public graph: MeetingBotGraphClient;
  public store: MeetingStore;
  public callHandler: CallHandler;
  public transcription: TranscriptionService;
  public summarization: SummarizationService;
  public notifications: NotificationService;
  public calendarWatcher: CalendarWatcher;

  constructor(private logger: ILogger) {
    this.graph = new MeetingBotGraphClient(logger.child("graph"));
    this.store = new MeetingStore(logger.child("store"));
    this.callHandler = new CallHandler(this.graph, this.store, logger.child("callHandler"));
    this.transcription = new TranscriptionService(logger.child("transcription"));
    this.summarization = new SummarizationService(logger.child("summarization"));
    this.notifications = new NotificationService(this.graph, logger.child("notifications"));
    this.calendarWatcher = new CalendarWatcher(
      this.graph,
      this.callHandler,
      logger.child("calendar")
    );
  }

  // Handles POST /api/calls — call state events from Graph Communications API
  async handleCallWebhook(body: unknown): Promise<void> {
    await this.callHandler.handleCallWebhook(
      body as { value?: Array<{ resource?: string; resourceData?: { state?: string } }> }
    );
  }

  // Handles POST /api/notifications — fired when Teams transcript is ready
  async handleGraphNotification(body: unknown): Promise<void> {
    const notification = body as {
      value?: Array<{ clientState?: string }>;
    };
    for (const item of notification.value || []) {
      const callId = item.clientState;
      if (!callId) continue;
      const session = this.store.getByCallId(callId);
      if (!session) continue;
      if (session.state === "transcript_pending" || session.state === "call_ended") {
        this.logger.debug(`📄 Transcript available for: ${callId}`);
        this.store.setState(callId, "transcript_ready");
        await this.processMeetingCompletion(session);
      }
    }
  }

  async processMeetingCompletion(session: MeetingSession): Promise<void> {
    const { callId } = session;
    try {
      this.store.setState(callId, "summarizing");

      // 1. Fetch transcript from Graph API
      const transcriptLines = await this.callHandler.fetchTranscript(
        session.meetingId,
        session.organizerUserId
      );
      const transcriptText = TranscriptionService.formatTranscript(transcriptLines);

      if (!transcriptText.trim()) {
        this.logger.warn(`⚠️ Empty transcript for: ${callId}`);
        this.store.setState(callId, "completed");
        return;
      }

      // 2. Generate AI summary via Azure OpenAI
      const summary = await this.summarization.summarize(transcriptText, session);

      // 3. Get meeting chat ID
      const chatId = await this.callHandler.getMeetingChatId(
        session.meetingId,
        session.organizerUserId
      );
      if (chatId) this.store.update(callId, { chatId });

      // 4. Send summary email to all participants
      await this.notifications.sendSummaryEmail(session, summary, transcriptText);

      // 5. Post summary card to meeting chat
      if (chatId) {
        await this.notifications.postSummaryToMeetingChat(chatId, session, summary);
      }

      this.store.setState(callId, "completed");
      this.logger.debug(`✅ Meeting processing complete: ${session.title}`);
    } catch (error) {
      this.logger.error(`❌ Processing failed for ${callId}: ${error}`);
      this.store.setState(callId, "failed");
    }
  }

  watchUserCalendar(userId: string, tenantId: string): void {
    this.calendarWatcher.startWatching(userId, tenantId);
  }

  async joinMeetingByUrl(
    joinWebUrl: string,
    organizerUserId: string,
    tenantId: string,
    title: string
  ): Promise<string> {
    const meetingId = await this.callHandler.resolveMeetingId(organizerUserId, joinWebUrl);
    if (!meetingId) throw new Error(`Could not resolve meeting ID from URL`);
    return this.callHandler.joinMeeting(joinWebUrl, meetingId, organizerUserId, tenantId, title);
  }

  getStatus(): { activeSessions: number; sessions: unknown[] } {
    return {
      activeSessions: this.store.size(),
      sessions: this.store.getAll().map((s) => ({
        callId: s.callId,
        title: s.title,
        state: s.state,
        participants: s.participants.length,
        actualStart: s.actualStart,
      })),
    };
  }
}
```

---

## 🔌 STEP 5 — MODIFY `src/index.ts`

Add the meeting bot to the existing application. Make these changes to the existing file:

**1. Add import at the top (after existing imports):**
```typescript
import { MeetingBotOrchestrator } from "./meetingBot/index";
```

**2. After `const app = new App({...})` declaration, add:**
```typescript
const meetingBot = new MeetingBotOrchestrator(logger.child("meetingBot"));
```

**3. The Teams AI SDK v2 `App` class exposes the underlying Express instance. Find the correct property name by inspecting the SDK. Try these in order until one works:**
```typescript
// Option A: app.server
// Option B: app.express  
// Option C: app.cloud
// Option D: Cast as any and inspect at runtime

// Check at startup which property exposes Express routes:
const expressApp = (app as any).server || (app as any).express || (app as any).cloud || (app as any);
```

**4. Register the three new HTTP routes. Add BEFORE the `app.on("install.add", ...)` line:**

```typescript
// ─── IMPORTANT: Determine the correct Express instance property ───────────────
// Check node_modules/@microsoft/teams.apps/dist/index.js for how to access Express
// Common properties: .server, .express, .cloud, .http
// Use whichever exposes app.use() / app.post() / app.get()
const expressServer = (app as any).express ?? (app as any).server ?? (app as any).cloud;

// ─── Meeting Bot: Graph Communications call webhook ───────────────────────────
// Teams calls this for every call state change (establishing, established, terminated)
// MUST respond 200 immediately — Teams stops retrying after 15 seconds
expressServer.post("/api/calls", async (req: any, res: any) => {
  res.status(200).send(); // ACK immediately
  try {
    await meetingBot.handleCallWebhook(req.body);
  } catch (err) {
    logger.error("❌ Call webhook error:", err);
  }
});

// ─── Meeting Bot: Graph change notifications (transcript ready) ───────────────
// Graph fires this when Teams has finished processing the meeting transcript
// Also handles subscription validation (validationToken query param)
expressServer.post("/api/notifications", async (req: any, res: any) => {
  if (req.query?.validationToken) {
    // Graph subscription validation — echo back the token
    res.status(200).contentType("text/plain").send(req.query.validationToken);
    return;
  }
  res.status(200).send(); // ACK immediately
  try {
    await meetingBot.handleGraphNotification(req.body);
  } catch (err) {
    logger.error("❌ Notification webhook error:", err);
  }
});

// ─── Meeting Bot: Status/health endpoint ─────────────────────────────────────
expressServer.get("/api/meeting-bot/status", (_req: any, res: any) => {
  res.json(meetingBot.getStatus());
});
```

**5. Extend the existing `app.on("message", ...)` handler with bot commands.**

In the existing message handler, add meeting bot command handling BEFORE the existing manager/missa logic:

```typescript
app.on("message", async ({ send, activity, api }) => {
  const rawText = activity.text || "";
  const text = rawText.toLowerCase().trim();
  const botMentioned = activity.entities?.some((e: any) => e.type === "mention");

  // ─── Meeting Bot Commands (only when bot is @mentioned) ──────────────────────
  if (botMentioned) {
    // Command: "@bot meeting status" — show active meeting sessions
    if (text.includes("meeting status") || text.includes("bot status")) {
      const status = meetingBot.getStatus();
      const sessionList =
        (status.sessions as any[]).map((s) => `• ${s.title} → ${s.state}`).join("\n") ||
        "No active sessions";
      await send(`📊 **Meeting Bot Status**\nActive sessions: ${status.activeSessions}\n\n${sessionList}`);
      return;
    }

    // Command: "@bot watch me" — auto-join the user's upcoming meetings
    if (text.includes("watch me") || text.includes("watch my calendar")) {
      const userId = activity.from.id;
      const tenantId = process.env.TENANT_ID!;
      meetingBot.watchUserCalendar(userId, tenantId);
      await send(
        `📅 Got it! I'll watch your calendar and automatically join your Teams meetings as **${process.env.MEETING_BOT_DISPLAY_NAME || "Meeting Notes Bot"}**.\n\nAfter each meeting ends, I'll send a summary email with action items and transcript to all participants.`
      );
      return;
    }

    // Command: "@bot join <URL>" — manually join a specific meeting by URL
    const joinMatch = rawText.match(/join (https:\/\/teams\.microsoft\.com[^\s]+)/i);
    if (joinMatch) {
      const joinUrl = joinMatch[1];
      const userId = activity.from.id;
      const tenantId = process.env.TENANT_ID!;
      try {
        await send("⏳ Joining meeting...");
        const callId = await meetingBot.joinMeetingByUrl(
          joinUrl,
          userId,
          tenantId,
          "Teams Meeting"
        );
        await send(
          `✅ Joined meeting! (Call ID: \`${callId}\`)\n\nI'll send a summary email to all participants after the meeting ends.`
        );
      } catch (err) {
        await send(
          `❌ Could not join meeting: ${err instanceof Error ? err.message : String(err)}`
        );
      }
      return;
    }

    // Command: "@bot help" — show available commands
    if (text.includes("help")) {
      await send(
        `👋 **Meeting Notes Bot** — Here's what I can do:\n\n` +
          `📅 **@bot watch me** — Auto-join your upcoming calendar meetings\n` +
          `🔗 **@bot join <URL>** — Manually join a specific meeting\n` +
          `📊 **@bot meeting status** — Show active meeting sessions\n\n` +
          `I also work as the **Missa Agent** — ask me to summarize conversations, find action items, or search messages!`
      );
      return;
    }
  }

  // ─── Existing Missa Agent logic (unchanged below this line) ────────────
  const context = botMentioned
    ? await createMessageContext(storage, activity, api)
    : await createMessageContext(storage, activity);

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
```

---

## 📋 STEP 6 — UPDATE `appPackage/manifest.json`

Replace the existing manifest with this updated version that adds calling support:

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.24/MicrosoftTeams.schema.json",
  "manifestVersion": "1.24",
  "version": "1.0.0",
  "id": "${{TEAMS_APP_ID}}",
  "developer": {
    "name": "My App, Inc.",
    "websiteUrl": "https://www.example.com",
    "privacyUrl": "https://www.example.com/privacy",
    "termsOfUseUrl": "https://www.example.com/termofuse"
  },
  "icons": {
    "color": "color.png",
    "outline": "outline.png"
  },
  "name": {
    "short": "Project missa${{APP_NAME_SUFFIX}}",
    "full": "Project Missa - AI Meeting Note Taker"
  },
  "description": {
    "short": "AI Meeting Note Taker",
    "full": "Joins Teams meetings automatically, transcribes conversations, and sends AI-generated summaries with action items to all participants."
  },
  "accentColor": "#FFFFFF",
  "bots": [
    {
      "botId": "${{BOT_ID}}",
      "scopes": ["personal", "team", "groupChat"],
      "supportsFiles": false,
      "isNotificationOnly": false,
      "supportsCalling": true,
      "supportsVideo": false
    }
  ],
  "validDomains": [
    "${{BOT_DOMAIN}}",
    "*.botframework.com"
  ],
  "webApplicationInfo": {
    "id": "${{BOT_ID}}",
    "resource": "api://botid-${{BOT_ID}}"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "ChatMessage.Read.Chat",
          "type": "Application"
        },
        {
          "name": "OnlineMeeting.ReadBasic.Chat",
          "type": "Application"
        }
      ]
    }
  }
}
```

---

## 🔧 STEP 7 — UPDATE `infra/azure.bicep`

Add an Azure Speech Service resource. Add this block after the `sqlFirewallRule` resource:

```bicep
param speechServiceName string = '${resourceBaseName}-speech'

resource speechService 'Microsoft.CognitiveServices/accounts@2023-05-01' = {
  name: speechServiceName
  location: location
  kind: 'SpeechServices'
  sku: {
    name: 'S0'
  }
  properties: {
    customSubDomainName: speechServiceName
    publicNetworkAccess: 'Enabled'
  }
}
```

Add to the `webApp` resource `appSettings` array inside `siteConfig`:
```bicep
{
  name: 'AZURE_SPEECH_KEY'
  value: speechService.listKeys().key1
}
{
  name: 'AZURE_SPEECH_REGION'
  value: location
}
{
  name: 'MEETING_BOT_DISPLAY_NAME'
  value: 'Meeting Notes Bot'
}
```

Add to `output` block at end of file:
```bicep
output SPEECH_SERVICE_ENDPOINT string = speechService.properties.endpoint
output SPEECH_SERVICE_REGION string = location
```

---

## ✅ STEP 8 — POST-DEPLOYMENT: ENABLE CALLING WEBHOOK

After deploying (or after dev tunnel starts), do this in Azure Portal:

1. Azure Portal → Bot Services → your bot
2. Click **Channels** → **Microsoft Teams** → click the pencil (Edit)
3. Click the **Calling** tab
4. Check ✅ **Enable calling**
5. Set Webhook (for calling): `https://YOUR_BOT_DOMAIN/api/calls`
6. Click **Save**

⚠️ This CANNOT be automated via Bicep. It must be done manually.

---

## 🎛️ STEP 9 — TEAMS ADMIN CENTER: ENABLE TRANSCRIPTION

For Teams native transcription to work:
1. Go to https://admin.teams.microsoft.com
2. Click **Meetings** → **Meeting Policies**
3. Open the **Global (Org-wide default)** policy
4. Under **Recording & Transcription**, set:
   - **Transcription**: On
   - **Cloud recording**: On (optional but recommended)
5. Click **Save**

---

## 🧪 STEP 10 — PROOF OF CONCEPT TEST PLAN

### Test 1: Manual Join (Day 1 Test)
1. Start a Teams meeting in your browser
2. In a Teams chat, @mention the bot: `@Project missa join https://teams.microsoft.com/l/meetup-join/...`
3. ✅ Expected: Bot appears in meeting roster as "Meeting Notes Bot"
4. ✅ Expected: Teams shows a "Transcription" banner in the meeting
5. Have a 2-minute conversation
6. End the meeting
7. Wait 5-15 minutes
8. ✅ Expected: Summary email received by all participants
9. ✅ Expected: Summary card posted to meeting chat

### Test 2: Calendar Auto-Join
1. @mention bot: `@Project missa watch me`
2. ✅ Expected: Bot confirms it will watch your calendar
3. Create a new Teams meeting starting in 12 minutes
4. ✅ Expected: Bot joins automatically ~1 minute before start
5. Have a conversation, end meeting
6. ✅ Expected: Same email + chat card result

### Test 3: Status Check
1. During active meeting: `@Project missa meeting status`
2. ✅ Expected: Shows active session with state "in_call"

### Test 4: Health Check
```bash
curl https://YOUR_BOT_DOMAIN/api/meeting-bot/status
# Expected: { "activeSessions": 0, "sessions": [] }
```

---

## 🔍 DEBUGGING GUIDE

### Bot fails to join meeting → 403 error:
- Check ALL permissions have Admin Consent (green checkmarks) in Azure Portal
- Verify `CLIENT_ID` and `CLIENT_SECRET` in `.localConfigs` are correct
- Ensure `BOT_ENDPOINT` is publicly reachable (check dev tunnel is running)

### Bot joins but transcription not showing:
- Check Teams Admin Center → transcription policy is enabled
- The `isTranscriptionEnabled: true` PATCH may fail if organizer policy blocks it — this is non-fatal
- Teams native transcription still starts if the tenant policy allows it

### No email sent after meeting:
- Verify `EMAIL_SENDER_UPN` is a real user mailbox in your tenant
- Verify `Mail.Send` has admin consent
- Check server logs for `📧 Email sent` or `❌` errors
- If transcript notification never fires, check subscription was created (`📡 Subscribed to transcripts`)

### Transcript notification fires but transcript is empty:
- Teams can take 5-30 minutes to process transcripts after meeting ends
- The subscription may expire if meeting is very long — max 4 hours
- Check: `GET /users/{id}/onlineMeetings/{id}/transcripts` manually via Graph Explorer

### `app.cloud` / Express route registration fails:
- Open `node_modules/@microsoft/teams.apps/dist/index.js`
- Search for `express` or `server` or `http` in the App class
- Use the correct property name to register routes
- Alternative: Create a SEPARATE Express server on a different port if needed

---

## 📊 LOG MESSAGES TO MONITOR

These are the exact log messages emitted by the meeting bot. Use them to trace the pipeline:

```
📅 Starting calendar watch for user: <userId>
⏰ Scheduled join for "<title>" in Xs
📞 Joining meeting: <title>
✅ Bot joined. Call ID: <id>
📞 Call event: <id> → establishing
📞 Call event: <id> → established
✅ Call established: <title>
🎙️ Transcription enabled for: <meetingId>
📡 Subscribed to transcripts. Sub ID: <id>
👥 N participants for <callId>
📞 Call event: <id> → terminated
📴 Call ended: <title>
📄 Transcript available for: <callId>
📥 Fetching transcript for: <title>
🤖 Summarizing: <title>
📧 Email sent to N recipients
💬 Summary card posted to meeting chat: <chatId>
✅ Meeting processing complete: <title>
```

---

## 📝 IMPLEMENTATION ORDER (Execute Exactly)

```
1.  npm install (after updating package.json)
2.  Create src/meetingBot/types.ts
3.  Create src/meetingBot/meetingStore.ts
4.  Create src/meetingBot/graphClient.ts
5.  Create src/meetingBot/transcriptionService.ts
6.  Create src/meetingBot/callHandler.ts
7.  Create src/meetingBot/summarizationService.ts
8.  Create src/meetingBot/notificationService.ts
9.  Create src/meetingBot/calendarWatcher.ts
10. Create src/meetingBot/index.ts
11. Modify src/index.ts (import + routes + commands)
12. Update appPackage/manifest.json
13. Update infra/azure.bicep
14. npm run build — fix any TypeScript errors
15. Configure Azure AD permissions + admin consent
16. Add AZURE_SPEECH_KEY/REGION to .env.*.user
17. Add MEETING_BOT_DISPLAY_NAME and EMAIL_SENDER_UPN to .env.*.user
18. Press F5 / run dev:teamsfx — start dev tunnel
19. Enable calling webhook in Azure Bot Service portal
20. Enable transcription in Teams Admin Center
21. npm run build && test
```

---

*This skills.md is the single source of truth. Build every file exactly as specified. When encountering TypeScript errors, check that imports match the exact exports from the existing codebase (especially ILogger from @microsoft/teams.common). The existing capabilities (summarizer, action items, search) must remain fully functional after all changes.*
ENDOFFILE
