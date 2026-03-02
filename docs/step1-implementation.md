# Step 1 Implementation Summary

This document summarizes the implementation of the meeting bot functionality for Project Missa.

## Architecture Overview

The meeting bot system consists of two main components:

1. **Project Missa** (Main Agent) - The orchestration layer that handles Teams chat interactions
2. **Meeting-Media-Bot** (Service) - A dedicated microservice for call handling and media processing

```
┌─────────────────────────┐     ┌─────────────────────────┐     ┌─────────────────┐
│      Microsoft Teams    │     │     Project Missa       │     │ meeting-media-  │
│         Client          │────>│    (Port 3978)          │────>│     bot         │
│                         │     │ + Internal API (3979)   │     │   (Port 4000)   │
└─────────────────────────┘     └─────────────────────────┘     └─────────────────┘
                                          │                              │
                                          v                              v
                               ┌─────────────────────────┐     ┌─────────────────┐
                               │   SQLite / MSSQL        │     │  Microsoft      │
                               │   Database              │     │  Graph API      │
                               └─────────────────────────┘     └─────────────────┘
```

## Components Implemented

### 1. Database Schema (Phase B)
Location: `src/storage/`

New types in `meetingTypes.ts`:
- `MeetingRecord` - Meeting metadata and status
- `MeetingParticipantRecord` - Participant information
- `TranscriptChunkRecord` - Real-time transcript chunks
- `MeetingSummaryRecord` - Generated meeting summaries

New database methods in `IDatabase`:
- `upsertMeeting()` - Create or update meeting records
- `updateMeetingStatus()` - Update meeting lifecycle status
- `getMeeting()` - Retrieve meeting by ID
- `upsertParticipants()` - Add/update meeting participants
- `getParticipants()` - Get participants for a meeting
- `appendTranscriptChunk()` - Store transcript chunks
- `getTranscriptChunks()` - Get chunks by meeting ID
- `getTranscriptByMeetingId()` - Get full transcript with metadata
- `getMeetingsByStatus()` - Query meetings by status

### 2. Environment Configuration (Phase C)
Location: `src/utils/config.ts`, `.env.example`

New configuration for:
- Bot endpoint URL
- Meeting-media-bot service URL and shared secret
- Azure Speech Service credentials (optional)
- Database connection settings

### 3. Meeting-Media-Bot Service (Phase F)
Location: `services/meeting-media-bot/`

Files:
- `src/index.ts` - Express server with webhook endpoints
- `src/config/index.ts` - Configuration management
- `src/graph/graphClient.ts` - Graph API client with retry logic
- `src/graph/callManager.ts` - Call join/leave operations
- `src/speech/transcriber.ts` - Azure Speech SDK integration
- `src/routes/callbacks.ts` - Graph notification webhooks
- `src/routes/api.ts` - REST API for meeting control

### 4. Project-Missa Internal API (Phase G)
Location: `src/routes/meetingApi.ts`

Endpoints:
- `POST /api/meeting-transcripts/chunk` - Receive transcript chunks
- `POST /api/meeting-capture/status` - Receive status updates
- `GET /api/meeting-transcripts/:meetingId` - Get full transcript
- `GET /api/meetings/:meetingId` - Get meeting details
- `GET /api/health` - Health check

### 5. Meeting Notes Capability (Phase H)
Location: `src/capabilities/meeting-notes/`

New functions added to `meeting-notes.ts`:
- `start_meeting_capture` - Join meeting and begin transcription
- `stop_meeting_capture` - Leave meeting and stop transcription

Updated `prompt.ts` with routing for capture commands.

### 6. Client Service (Phase H)
Location: `src/services/meetingMediaBotClient.ts`

Client class for communicating with meeting-media-bot:
- `startMeetingCapture()` - Request bot to join meeting
- `stopMeetingCapture()` - Request bot to leave meeting
- `getMeetingCaptureStatus()` - Check capture status
- `checkHealth()` - Verify service availability

## Usage Flow

1. User sends: "@missa start capture https://teams.microsoft.com/l/meetup-join/..."
2. Meeting-notes capability:
   - Creates meeting record in database (status: "joining")
   - Calls MeetingMediaBotClient.startMeetingCapture()
3. Meeting-media-bot:
   - Joins Teams meeting via Graph Cloud Communications API
   - Starts Azure Speech transcription
   - Streams transcript chunks to Project-Missa internal API
4. Project-Missa internal API:
   - Stores chunks in database
   - Updates meeting status
5. User sends: "@missa stop capture"
6. Meeting-media-bot leaves meeting
7. User sends: "@missa summarize meeting"
8. Meeting-notes capability:
   - Retrieves transcript from database
   - Generates structured summary
   - Offers to email/share

## Required Permissions

### Graph API (Application)
- `Calls.JoinGroupCall.All` - Join meetings
- `Calls.AccessMedia.All` - Access audio streams
- `OnlineMeetings.Read.All` - Read meeting details
- `OnlineMeetingTranscript.Read.All` - Read transcripts (fallback)

### Azure Cognitive Services
- Azure Speech Services subscription for transcription

## Environment Variables

```env
# Bot Configuration
BOT_ENDPOINT=https://your-bot.azurewebsites.net
BOT_APP_ID=your-app-id
BOT_APP_PASSWORD=your-app-secret

# Meeting Media Bot
MEETING_MEDIA_BOT_URL=http://localhost:4000
MEETING_MEDIA_BOT_SHARED_SECRET=your-shared-secret

# Azure Speech
AZURE_SPEECH_KEY=your-speech-key
AZURE_SPEECH_REGION=eastus

# Database
SQL_CONNECTION_STRING=your-connection-string
```

## Next Steps

1. **Test locally** - Run both services and test meeting capture flow
2. **Configure Graph permissions** - Register app with required permissions
3. **Deploy meeting-media-bot** - Deploy as separate Azure App Service
4. **Set up dev tunnel** - For local development callback URLs
5. **Implement Graph transcript fallback** - For meetings where real-time capture wasn't used

## Files Changed/Created

### New Files
- `src/storage/meetingTypes.ts`
- `src/routes/meetingApi.ts`
- `src/services/meetingMediaBotClient.ts`
- `.env.example`
- `docs/manifest-step1.md`
- `docs/graph-permissions-step1.md`
- `services/meeting-media-bot/` (entire directory)

### Modified Files
- `src/storage/database.ts` - Added meeting interfaces
- `src/storage/storage.ts` - Implemented meeting methods (SQLite)
- `src/storage/mssqlStorage.ts` - Implemented meeting methods (MSSQL)
- `src/utils/config.ts` - Added loadConfig, AppConfig
- `src/utils/messageContext.ts` - Added database, userAadId
- `src/capabilities/meeting-notes/meeting-notes.ts` - Added capture functions
- `src/capabilities/meeting-notes/prompt.ts` - Updated routing logic
- `src/index.ts` - Added internal API server
- `tsconfig.json` - Excluded services folder
