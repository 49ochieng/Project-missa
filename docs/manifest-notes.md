# Teams App Manifest Configuration Notes

## Overview
This document describes the Teams app manifest configuration for **Missa** and explains the scopes, permissions, and features enabled for the bot.

---

## Manifest Changes (Task 3A)

### Date: March 2, 2026

### What Changed

Added calling/meetings bot features to the `bots` section:

```json
"supportsCalling": true,
"supportsVideo": false
```

### Why These Changes Were Made

1. **supportsCalling: true**
   - Enables the bot to participate in Teams calls and meetings
   - Allows the bot to receive meeting-related events (meeting start, end, participant join/leave)
   - Required for meeting transcript access and real-time meeting interactions
   - Supports the **meeting-notes** capability which needs to read and summarize meeting transcripts

2. **supportsVideo: false**
   - Explicitly disables video capabilities for now
   - Reduces complexity during initial development
   - Can be enabled in future if video features are needed (e.g., bot video presence in meetings)
   - Keeps the bot lightweight and focused on text/audio processing

---

## Current Manifest Configuration

### Schema Version
- **Version**: `1.24` 
- **Schema URL**: `https://developer.microsoft.com/en-us/json-schemas/teams/v1.24/MicrosoftTeams.schema.json`
- **Status**: Valid and current (as of March 2026)

### Bot Scopes

The bot supports three scopes:

1. **personal** - One-on-one chat with the bot
   - Users can @mention Missa in personal chats
   - Supports all capabilities: action items, search, summarizer, meeting notes
   - Example: Direct message to @Missa asking for help

2. **team** - Team channels and meetings
   - Bot can be added to team channels
   - **Includes meeting scope** - bot can participate in Teams meetings
   - Receives meeting events when supportsCalling is true
   - Example: @Missa in a team channel or during a team meeting

3. **groupChat** - Group conversations
   - Bot can be added to group chats (multi-person non-team conversations)
   - Example: Group chat with 3-5 people where @Missa is added

### Bot Features

```json
{
  "supportsFiles": false,
  "isNotificationOnly": false,
  "supportsCalling": true,
  "supportsVideo": false
}
```

- **supportsFiles**: Disabled - bot does not handle file uploads/downloads directly
- **isNotificationOnly**: False - bot is interactive (responds to messages)
- **supportsCalling**: **Enabled** - bot can join calls and meetings (Task 3A)
- **supportsVideo**: **Disabled** - no video stream support (Task 3A)

### Resource-Specific Consent (RSC) Permissions

The bot uses RSC permissions to listen to messages without being @mentioned:

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

- **ChatMessage.Read.Chat**: Application permission to read chat messages
- Allows the bot to receive all messages in chats/channels where it's installed
- **Critical for existing functionality**: DO NOT REMOVE - the bot's message handler relies on this
- Used for context-aware features (e.g., summarizing conversation history)

---

## Meeting Scope Support

### How Meetings Work with This Configuration

1. **Installation**: 
   - The bot can be added to team channels (via "team" scope)
   - When added to a channel, it automatically has access to that team's meetings

2. **Meeting Events**:
   - With `supportsCalling: true`, the bot receives:
     - `application/vnd.microsoft.meetingStart`
     - `application/vnd.microsoft.meetingEnd`
     - `application/vnd.microsoft.meetingParticipantJoin`
     - `application/vnd.microsoft.meetingParticipantLeave`
     - And other meeting-related events

3. **Meeting Transcript Access**:
   - The **meeting-notes** capability uses Microsoft Graph API to:
     - `GET /users/{userId}/onlineMeetings/{meetingId}/transcripts`
     - `GET /users/{userId}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content`
   - Requires Azure AD app permissions (see environment configuration)

### Example Meeting Use Cases

- **@Missa summarize this meeting** - Reads transcript and generates summary
- **@Missa send meeting notes to the team** - Emails summary with action items
- **@Missa what were the action items from today's standup?** - Extracts action items from meeting transcript

---

## Validation

### Schema Validation
The manifest follows the official Microsoft Teams app manifest schema v1.24.

To validate the manifest:
```bash
# Teams Toolkit validates automatically on deploy
teamsfx validate

# Or manually validate against schema
npx @microsoft/teams-manifest-validator appPackage/manifest.json
```

### Required Environment Variables

The manifest uses template variables that are replaced during deployment:

- `${{TEAMS_APP_ID}}` - Unique Teams app ID
- `${{BOT_ID}}` - Bot registration ID (Azure Bot Service)
- `${{BOT_DOMAIN}}` - Bot endpoint domain (e.g., dev tunnel or Azure hostname)
- `${{APP_NAME_SUFFIX}}` - Environment suffix (e.g., "(local)", "(dev)", "(prod)")

These are populated from `.env.local`, `.env.sandbox`, etc.

---

## Breaking Changes - NONE ✅

### What Was Preserved

1. **RSC Permissions**: `ChatMessage.Read.Chat` permission retained
   - Existing chat listener functionality still works
   - Bot can still receive messages without @mention

2. **Existing Scopes**: All three scopes unchanged (personal, team, groupChat)
   - No removal or modification of existing installation contexts
   - Backward compatible with existing installations

3. **File Support**: Kept as `false`
   - No change to file handling behavior

4. **Notification Mode**: Kept as `false`
   - Bot remains interactive

### What Was Added (Non-Breaking)

- `supportsCalling: true` - NEW feature, does not break existing functionality
- `supportsVideo: false` - NEW flag, explicitly disables video (no impact on existing features)

---

## Future Considerations

### Potential Future Enhancements

1. **supportsVideo: true**
   - Enable if we want bot video presence in meetings
   - Requires additional infrastructure (video processing, streaming)
   - Use case: Visual indicators, recordings with video

2. **supportsFiles: true**
   - Enable if we want to handle file uploads
   - Use case: "Analyze this PDF", "Summarize this document"

3. **Additional RSC Permissions**
   - `TeamMember.Read.Group` - Read team member info
   - `ChannelMessage.Read.Group` - Read channel messages (broader than ChatMessage.Read.Chat)
   - `OnlineMeetingTranscript.Read.Chat` - Future RSC for meeting transcripts (if/when available)

4. **Meeting Stage Extension**
   - Add `configurableTabs` with `meetingDetailsTab` or `meetingStageTab`
   - Use case: In-meeting UI panel showing live meeting notes
   - Requires frontend implementation (React app)

---

## References

- [Teams App Manifest Schema v1.24](https://learn.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema)
- [Bots in Teams Meetings](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/meeting-events)
- [RSC Permissions](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/resource-specific-consent)
- [Meeting Transcripts API](https://learn.microsoft.com/en-us/graph/api/resources/calltranscript)

---

## Summary

✅ **Task 3A Complete**
- Personal scope: ✅ Supported (existing)
- Group chat: ✅ Supported (existing)
- Meeting scope: ✅ Supported (via "team" scope)
- supportsCalling: ✅ Enabled (new)
- supportsVideo: ✅ Explicitly disabled (new)
- Valid schema: ✅ v1.24
- RSC permissions: ✅ Preserved
- No breaking changes: ✅ Confirmed

The manifest is now configured to support all required scopes and calling/meeting features while maintaining backward compatibility with existing functionality.
