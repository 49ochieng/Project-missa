# Microsoft Graph API Permissions - Step 1

## Overview

This document details the Microsoft Graph API permissions required for the Missa meeting bot to:
1. Join Teams meetings as a visible participant
2. Access meeting audio for real-time transcription
3. Read meeting details and participants
4. Fetch meeting transcripts (fallback)
5. Send notifications (email and Teams messages)

---

## Required Application Permissions

### 1. Meeting Join and Calling Bot

| Permission | Type | Purpose | Required |
|------------|------|---------|----------|
| `Calls.JoinGroupCall.All` | Application | Join Teams meetings programmatically | **Yes** |
| `Calls.InitiateGroupCall.All` | Application | Create and initiate calls to meetings | **Yes** |
| `Calls.AccessMedia.All` | Application | Access audio streams from meetings | **Yes** |

#### How They're Used

**Calls.JoinGroupCall.All**
- Allow the bot to join an existing meeting via join URL
- Required for making the POST request to `/communications/calls`

**Calls.InitiateGroupCall.All**
- Required when initiating a call to a meeting that hasn't started
- Used alongside JoinGroupCall for full meeting coverage

**Calls.AccessMedia.All**
- **Critical**: Required to receive audio streams
- Without this, the bot joins but cannot receive media
- Enables the Real-time Media Platform to stream audio to the bot

---

### 2. Meeting Details and Participants

| Permission | Type | Purpose | Required |
|------------|------|---------|----------|
| `OnlineMeetings.Read.All` | Application | Read meeting details, participants, attendance | **Yes** |
| `Calendars.Read` | Application | Read calendar events for attendee emails | Optional |

#### How They're Used

**OnlineMeetings.Read.All**
- Fetch meeting metadata (title, organizer, start/end time)
- Get participant list with AAD IDs for speaker identification
- Required endpoints:
  - `GET /users/{userId}/onlineMeetings/{meetingId}`
  - `GET /communications/callRecords/{id}/sessions`

**Calendars.Read** (Optional)
- Enrich participant data with email addresses
- Only needed if you want attendee emails not available in meeting roster
- May be omitted in Step 1

---

### 3. Transcript Fallback Reading

| Permission | Type | Purpose | Required |
|------------|------|---------|----------|
| `OnlineMeetingTranscript.Read.All` | Application | Read meeting transcripts from Graph | **Yes** |
| `CallTranscripts.Read.All` | Application | Read call transcripts (some tenants) | Conditional |

#### How They're Used

**OnlineMeetingTranscript.Read.All**
- Primary fallback when real-time speech transcription fails
- Required endpoints:
  - `GET /users/{userId}/onlineMeetings/{meetingId}/transcripts`
  - `GET /users/{userId}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content`

**CallTranscripts.Read.All**
- Alternative permission in some tenant configurations
- May be required depending on how transcripts are stored
- Add if `OnlineMeetingTranscript.Read.All` doesn't return transcripts

---

### 4. Notifications (Future - Step 2+)

| Permission | Type | Purpose | Required |
|------------|------|---------|----------|
| `Mail.Send` | Application | Send meeting summary emails | Step 2 |
| `Chat.Create` | Application | Create chats for summary distribution | Step 2 |
| `ChatMessage.Send` | Application | Post summaries to meeting chat | Step 2 |

These are not required for Step 1 but documented for planning.

---

## Application Access Policy

### Important Notice

> **Online meetings application access policy may be required** to scope which users/meetings the app can access.

#### What This Means

By default, application permissions grant access to **all** meetings in the tenant. Microsoft provides an access policy mechanism to restrict this:

```powershell
# Create an application access policy
New-CsApplicationAccessPolicy -Identity "MissaMeetingBotPolicy" -AppIds "your-app-id" -Description "Policy for Missa meeting bot"

# Grant policy to specific users
Grant-CsApplicationAccessPolicy -PolicyName "MissaMeetingBotPolicy" -Identity "user@domain.com"
```

#### When to Configure

- **Development/Testing**: Not required (full tenant access is fine)
- **Production**: Consider restricting to specific organizers/groups
- **Enterprise**: Often required by IT security policies

#### Reference
[Configure application access policy](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)

---

## Azure AD App Registration Setup

### Step-by-Step Configuration

1. **Navigate to Azure Portal**
   - Go to Azure Active Directory → App registrations
   - Select your bot's app registration (or create new)

2. **Add API Permissions**
   - Click "API permissions" → "Add a permission" → "Microsoft Graph"
   - Select "Application permissions"
   - Add each permission listed above

3. **Grant Admin Consent**
   - Click "Grant admin consent for {tenant}"
   - This requires Global Administrator or Privileged Role Administrator

4. **Verify Permissions**
   - Ensure all permissions show "Granted for {tenant}" status

### Permission Summary Table

| Category | Permission | Status |
|----------|------------|--------|
| Calling | `Calls.JoinGroupCall.All` | ⬜ Add |
| Calling | `Calls.InitiateGroupCall.All` | ⬜ Add |
| Calling | `Calls.AccessMedia.All` | ⬜ Add |
| Meetings | `OnlineMeetings.Read.All` | ⬜ Add |
| Transcripts | `OnlineMeetingTranscript.Read.All` | ⬜ Add |
| Transcripts | `CallTranscripts.Read.All` | ⬜ Add (if needed) |
| Calendar | `Calendars.Read` | ⬜ Optional |

---

## Client Credentials Authentication

### Environment Variables Required

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-app-id
SECRET_AZURE_CLIENT_SECRET=your-client-secret
```

### Token Acquisition

```typescript
// Example using @azure/identity
import { ClientSecretCredential } from "@azure/identity";

const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_CLIENT_SECRET
);

// Get token for Microsoft Graph
const token = await credential.getToken("https://graph.microsoft.com/.default");
```

### Security Notes

- **Never log tokens or secrets**
- Use environment variables, not hardcoded values
- Rotate client secrets regularly
- Consider using Azure Key Vault in production

---

## API Endpoints Used

### Joining a Meeting

```http
POST https://graph.microsoft.com/v1.0/communications/calls
Authorization: Bearer {access_token}
Content-Type: application/json

{
  "@odata.type": "#microsoft.graph.call",
  "callbackUri": "https://your-bot-endpoint/api/calls/callback",
  "requestedModalities": ["audio"],
  "mediaConfig": {
    "@odata.type": "#microsoft.graph.appHostedMediaConfig",
    "blob": "<media-blob-configuration>"
  },
  "chatInfo": {
    "@odata.type": "#microsoft.graph.chatInfo",
    "threadId": "19:meeting_xxx@thread.v2",
    "messageId": "0"
  },
  "meetingInfo": {
    "@odata.type": "#microsoft.graph.organizerMeetingInfo",
    "organizer": {
      "@odata.type": "#microsoft.graph.identitySet",
      "user": {
        "@odata.type": "#microsoft.graph.identity",
        "id": "organizer-aad-id",
        "tenantId": "tenant-id"
      }
    }
  },
  "tenantId": "tenant-id"
}
```

### Reading Meeting Details

```http
GET https://graph.microsoft.com/v1.0/users/{userId}/onlineMeetings/{meetingId}
Authorization: Bearer {access_token}
```

### Listing Transcripts

```http
GET https://graph.microsoft.com/v1.0/users/{userId}/onlineMeetings/{meetingId}/transcripts
Authorization: Bearer {access_token}
```

### Fetching Transcript Content

```http
GET https://graph.microsoft.com/v1.0/users/{userId}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content
Authorization: Bearer {access_token}
Accept: text/vtt
```

---

## Troubleshooting

### Common Permission Errors

| Error Code | Message | Solution |
|------------|---------|----------|
| `403` | Forbidden | Admin consent not granted |
| `401` | Unauthorized | Token expired or invalid |
| `404` | Resource not found | Meeting ID incorrect or no access |
| `AccessDenied` | Application access policy | Configure access policy for user |

### Checking Effective Permissions

```http
GET https://graph.microsoft.com/v1.0/me/oauth2PermissionGrants
Authorization: Bearer {access_token}
```

### Verifying Consent

1. Go to Azure Portal → Enterprise Applications
2. Find your app
3. Check "Permissions" tab for granted permissions

---

## References

- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Cloud Communications Overview](https://learn.microsoft.com/en-us/graph/cloud-communications-concept-overview)
- [Calling API Overview](https://learn.microsoft.com/en-us/graph/api/resources/communications-api-overview)
- [Online Meetings API](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting)
- [Meeting Transcripts API](https://learn.microsoft.com/en-us/graph/api/resources/calltranscript)
- [Application Access Policy](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)

---

## Checklist

### Before Step 1 Testing

- [ ] All permissions added to Azure AD app registration
- [ ] Admin consent granted for all permissions
- [ ] Environment variables configured in `.env.local`
- [ ] Client credentials authentication tested
- [ ] Access policy configured (if required by IT)

### Verification Commands

```powershell
# Test Graph API connectivity
$token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers @{Authorization="Bearer $token"}
```
