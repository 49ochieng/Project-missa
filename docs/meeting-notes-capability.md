# Meeting Notes Capability

The Meeting Notes capability enables the Missa bot to manage meeting transcripts, generate structured summaries, and distribute meeting notes to participants via email and Teams chat.

## Features

### 1. Read Transcript
Fetch and store meeting transcripts from Microsoft Teams meetings using Microsoft Graph API.

### 2. Summarize Meeting
Generate a comprehensive, structured JSON summary from meeting transcripts or conversation history, including:
- **Title & DateTime**: Meeting details
- **Participants**: List of attendees
- **Decisions**: Key decisions made
- **Action Items**: Tasks with owners and due dates
- **Risks**: Identified concerns or blockers
- **Open Questions**: Unresolved topics for follow-up
- **Short Summary**: Executive overview (2-3 sentences)
- **Detailed Summary**: Comprehensive meeting notes

### 3. Send Summary
Distribute meeting summaries to participants via:
- Email (Microsoft Graph API)
- Teams chat post (optional)

## Usage in Teams Chat

All commands support natural language - the bot will understand variations of these patterns.

### Reading a Transcript

```
@Missa read the transcript for meeting abc123
@Missa fetch meeting transcript from [meeting link]
@Missa get the notes from yesterday's standup
```

**Response Example:**
```
✅ Retrieved transcript for "Q1 Planning Meeting" on March 2, 2026
   - Duration: 45 minutes
   - Participants: Alice, Bob, Carol
   - Transcript stored successfully.
   
Would you like me to summarize this meeting?
```

### Summarizing a Meeting

```
@Missa summarize the meeting
@Missa create meeting summary
@Missa analyze the last meeting transcript
@Missa generate notes from the discussion
```

**Response Example:**
```
📋 Meeting Summary: Q1 Planning Meeting
📅 Date: 2026-03-02T10:00:00Z
👥 Participants: Alice, Bob, Carol

✅ Decisions
- Finalized Q1 budget allocation: 60% engineering, 40% marketing
- Agreed to weekly progress check-ins every Monday at 2 PM

📌 Action Items
- Alice: Complete proposal draft (Due: 2026-03-10)
- Bob: Review budget numbers with finance (Due: ASAP)
- Carol: Schedule stakeholder presentation (Due: 2026-03-15)

⚠️ Risks
- Tight timeline for Q1 deliverables
- Budget constraints may limit hiring plans

❓ Open Questions
- Need final approval from executive team?
- Should we engage external consultants?

📝 Summary
The team reviewed Q1 priorities and finalized budget allocation with a focus on engineering 
investments. Key decisions were made on resource distribution and weekly cadence for tracking.

Would you like me to send this summary to all participants?
```

### Sending a Summary

```
@Missa send the summary to all participants
@Missa email the meeting notes
@Missa share this summary with the team
@Missa distribute notes and post to chat
```

**Response Example:**
```
✅ Meeting summary sent successfully!
   - Email sent to: alice@example.com, bob@example.com, carol@example.com
   - Posted to Teams chat: Yes
   
Summary delivered to 3 participants.
```

## Complete Workflow Example

```
User: @Missa read the transcript for meeting abc123

Bot: ✅ Retrieved transcript for "Sprint Planning" on March 2, 2026
     - Duration: 60 minutes
     - Participants: Dev Team (5 members)
     - Transcript stored successfully.
     
     Would you like me to summarize this meeting?

User: Yes, please summarize it

Bot: [Generates structured summary with all sections]
     
     Would you like me to send this summary to all participants?

User: Yes, send via email and post to this chat

Bot: ✅ Meeting summary sent successfully!
     - Email sent to: 5 participants
     - Posted to Teams chat: Yes
```

## Advanced Usage

### Summarize Current Conversation
If you don't have a specific meeting transcript, you can summarize the current conversation:

```
@Missa summarize our discussion from the past hour
@Missa create meeting notes from today's chat
```

The bot will analyze the conversation history and generate a structured summary.

### Custom Recipient Lists
Specify custom recipients instead of all participants:

```
@Missa send the summary to alice@example.com and bob@example.com
@Missa email notes to the stakeholders only
```

### Email Only (No Teams Post)
By default, summaries are sent via email only. To also post to Teams:

```
@Missa send summary and post to this chat
@Missa share via email and Teams
```

## Technical Details

### Structured Output Schema
The capability generates JSON-formatted summaries conforming to this schema:

```json
{
  "title": "Meeting Topic",
  "dateTime": "2026-03-02T10:00:00Z",
  "participants": ["Alice", "Bob"],
  "decisions": ["Decision 1", "Decision 2"],
  "actionItems": [
    {
      "owner": "Alice",
      "task": "Complete proposal",
      "due": "2026-03-10"
    }
  ],
  "risks": ["Risk 1"],
  "openQuestions": ["Question 1"],
  "shortSummary": "Executive summary...",
  "detailedSummary": "Comprehensive summary..."
}
```

### Database Storage
The capability stores:
- **Meeting Transcripts**: Full transcript text with metadata
- **Meeting Summaries**: Structured summaries for future reference
- **Send History**: Track when and to whom summaries were sent

### Microsoft Graph API Integration

The capability uses Microsoft Graph API for:
- **Read Transcript**: `GET /communications/onlineMeetings/{meetingId}/transcripts`
  - Requires: `OnlineMeetingTranscript.Read.All` permission
- **Send Email**: `POST /users/{userId}/sendMail`
  - Requires: `Mail.Send` permission
- **Post to Teams**: `POST /teams/{teamId}/channels/{channelId}/messages`
  - Requires: `ChannelMessage.Send` permission

> **Note**: The current implementation includes stub functions with TODO markers for Graph API integration. These need to be completed with actual API calls and proper authentication (see [meeting-notes.ts](../src/capabilities/meeting-notes/meeting-notes.ts) for implementation details).

## Future Enhancements

Planned features:
- [ ] Automatic meeting detection (listen for meeting end events)
- [ ] Multi-language transcript support
- [ ] Custom summary templates
- [ ] Integration with task management systems (Planner, DevOps)
- [ ] Recurring meeting pattern detection
- [ ] Meeting analytics and insights

## Troubleshooting

### "Failed to fetch meeting transcript"
- Verify the meeting ID is correct
- Ensure the bot has `OnlineMeetingTranscript.Read.All` permission
- Check that the meeting has ended and transcript is available

### "Failed to send meeting summary"
- Verify recipient email addresses are valid
- Ensure the bot has `Mail.Send` permission
- Check Azure AD app registration configuration

### "Summary not structured correctly"
- The AI may struggle with very short or informal conversations
- Provide more context or use a longer transcript
- Try re-running the summarization with different phrasing

## See Also

- [Action Items Capability](./capabilities-audit.md#action-items) - Extract action items from conversations
- [Summarizer Capability](./capabilities-audit.md#summarizer) - General conversation summarization
- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/api/overview)
- [Teams AI Library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/teams-conversation-ai-overview)
