# Meeting Bot Manifest Configuration - Step 1

## Overview

This document describes the Teams app manifest changes required for the meeting bot functionality in Step 1 of the implementation. These changes enable the bot to:
- Join Teams meetings as a participant
- Receive audio streams for real-time transcription
- Respond to @mentions during meetings
- Access meeting lifecycle events

---

## Manifest Configuration (v1.24)

### Schema Version

```json
"$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.24/MicrosoftTeams.schema.json",
"manifestVersion": "1.24"
```

Using schema version 1.24 which is the current stable version supporting all calling/meeting features.

---

## Bot Configuration

### Current Bot Section

```json
"bots": [
    {
        "botId": "${{BOT_ID}}",
        "scopes": [
            "personal",
            "team",
            "groupChat"
        ],
        "supportsFiles": false,
        "isNotificationOnly": false,
        "supportsCalling": true,
        "supportsVideo": false
    }
]
```

### Scope Details

| Scope | Purpose | Meeting Access |
|-------|---------|----------------|
| **personal** | 1:1 chat with users | N/A |
| **team** | Channel conversations + **meeting access** | ✅ Full meeting support |
| **groupChat** | Multi-user conversations | Limited meeting context |

**Important**: The `team` scope provides meeting access. When a bot is installed in a team, it automatically has access to that team's scheduled meetings and can receive meeting-related events.

### Calling Properties

| Property | Value | Purpose |
|----------|-------|---------|
| `supportsCalling` | `true` | **Required** - Enables bot to join calls/meetings and receive audio |
| `supportsVideo` | `false` | Video capabilities disabled (not needed for transcription) |

---

## Resource-Specific Consent (RSC) Permissions

### Current RSC Permissions

```json
"authorization": {
    "permissions": {
        "resourceSpecific": [
            {
                "name": "ChatMessage.Read.Chat",
                "type": "Application"
            }
        ]
    }
}
```

### Permission Details

| Permission | Type | Purpose |
|------------|------|---------|
| `ChatMessage.Read.Chat` | Application | Read all messages in chats where bot is installed |

**Note**: This existing permission is preserved and essential for the bot's chat listening functionality. It allows the bot to receive all messages in channels/chats without requiring explicit @mentions.

---

## Meeting Event Support

With `supportsCalling: true`, the bot can receive the following meeting events:

### Meeting Lifecycle Events

| Event Type | When Triggered | Handler |
|------------|----------------|---------|
| `application/vnd.microsoft.meetingStart` | Meeting begins | Start transcription |
| `application/vnd.microsoft.meetingEnd` | Meeting ends | Save final transcript |

### Participant Events

| Event Type | When Triggered | Handler |
|------------|----------------|---------|
| `application/vnd.microsoft.meetingParticipantJoin` | User joins meeting | Update participant list |
| `application/vnd.microsoft.meetingParticipantLeave` | User leaves meeting | Update participant list |

### Call Events (for joining via Graph)

| Event Type | Purpose |
|------------|---------|
| `Microsoft.Communication.CallStateChange` | Track call state (joining → established → terminated) |
| `Microsoft.Communication.ParticipantsUpdated` | Real-time roster updates |

---

## Validation Checklist

### Before Deployment

- [x] Schema version: v1.24 ✓
- [x] `supportsCalling: true` in bot section ✓
- [x] All required scopes: personal, team, groupChat ✓
- [x] RSC permission preserved: ChatMessage.Read.Chat ✓
- [x] Valid domain configuration ✓

### JSON Validation

```powershell
# Validate manifest is valid JSON
Get-Content "appPackage/manifest.json" | ConvertFrom-Json | Out-Null

# Or use Teams Toolkit
teamsfx validate
```

---

## Changes from Previous Version

| Aspect | Before | After | Reason |
|--------|--------|-------|--------|
| `supportsCalling` | Not present | `true` | Enable meeting join capability |
| `supportsVideo` | Not present | `false` | Explicitly disable video (not needed) |
| Scopes | Same | Same | Already had meeting-capable scopes |
| RSC Permissions | Same | Same | Preserved existing functionality |

### No Breaking Changes

- Existing chat listening via RSC permission: **Preserved** ✓
- All original scopes: **Unchanged** ✓
- Notification mode: **Still interactive** (`isNotificationOnly: false`) ✓

---

## How Meeting Access Works

### Installation Flow

1. User installs Missa app in a Team
2. Bot is registered with `team` scope
3. Bot automatically gets access to:
   - All channels in that team
   - All scheduled meetings for that team
   - Meeting chat threads

### Meeting Join Flow (Graph API)

1. User sends: `@Missa start meeting capture <joinUrl>`
2. Missa parses the join URL and extracts meeting ID
3. meeting-media-bot service calls Graph API:
   ```http
   POST /communications/calls
   {
     "callbackUri": "{BOT_ENDPOINT}/api/calls/callback",
     "requestedModalities": ["audio"],
     "mediaConfig": {
       "@odata.type": "#microsoft.graph.appHostedMediaConfig"
     },
     "chatInfo": {
       "threadId": "{meetingId}",
       "messageId": "0"
     },
     "meetingInfo": {
       "@odata.type": "#microsoft.graph.organizerMeetingInfo",
       "organizer": {
         "user": {
           "id": "{organizerId}"
         }
       }
     }
   }
   ```
4. Bot appears in meeting participant roster
5. Audio streams to bot for transcription

---

## Required Azure AD App Permissions

While not in the manifest, the Azure AD app registration must have these Graph API permissions for meeting functionality:

| Permission | Type | Purpose |
|------------|------|---------|
| `Calls.JoinGroupCall.All` | Application | Join meetings programmatically |
| `Calls.InitiateGroupCall.All` | Application | Create calls to meetings |
| `Calls.AccessMedia.All` | Application | Access audio/video streams |
| `OnlineMeetings.Read.All` | Application | Read meeting details |
| `OnlineMeetingTranscript.Read.All` | Application | Read meeting transcripts (fallback) |

See [graph-permissions-step1.md](graph-permissions-step1.md) for the complete permissions documentation.

---

## Related Files

| File | Purpose |
|------|---------|
| `appPackage/manifest.json` | Teams app manifest |
| `docs/graph-permissions-step1.md` | Graph API permissions guide |
| `docs/manifest-notes.md` | General manifest documentation |
| `.env.example` | Environment variables template |

---

## Troubleshooting

### Bot doesn't appear in meeting roster

1. Verify `supportsCalling: true` is set
2. Check Azure AD app has `Calls.JoinGroupCall.All` permission
3. Confirm admin consent is granted
4. Verify the meeting join URL is valid

### Bot can't receive audio

1. Ensure `Calls.AccessMedia.All` permission is granted
2. Check callback URL is publicly accessible
3. Verify Real-time Media Platform is configured

### Meeting events not received

1. Confirm bot is installed in the team where meeting is scheduled
2. Check webhook endpoints are registered
3. Verify callback URL responds to validation requests

---

## Next Steps

After manifest deployment:
1. Configure Azure AD app permissions ([Phase E](graph-permissions-step1.md))
2. Deploy meeting-media-bot service (Phase F)
3. Test meeting join flow (Phase H)
