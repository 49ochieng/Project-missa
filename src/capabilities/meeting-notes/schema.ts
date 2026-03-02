/**
 * Schema definitions for meeting notes capability
 */

export interface MeetingSummary {
  title: string;
  dateTime: string;
  participants: string[];
  decisions: string[];
  actionItems: ActionItem[];
  risks: string[];
  openQuestions: string[];
  shortSummary: string;
  detailedSummary: string;
}

export interface ActionItem {
  owner: string;
  task: string;
  due: string;
}

export interface MeetingTranscript {
  meetingId: string;
  title: string;
  startDateTime: string;
  endDateTime: string;
  organizer: string;
  participants: string[];
  transcriptText: string;
  retrievedAt: string;
}

export interface MeetingEmailRequest {
  recipients: string[];
  subject: string;
  summary: MeetingSummary;
  includeTeamsPost: boolean;
  conversationId?: string;
}

/**
 * JSON schema for structured meeting summary output
 */
export const MEETING_SUMMARY_SCHEMA = {
  type: "object",
  properties: {
    title: {
      type: "string",
      description: "Meeting title or main topic",
    },
    dateTime: {
      type: "string",
      description: "ISO 8601 datetime of the meeting",
    },
    participants: {
      type: "array",
      items: { type: "string" },
      description: "List of participant names",
    },
    decisions: {
      type: "array",
      items: { type: "string" },
      description: "Key decisions made during the meeting",
    },
    actionItems: {
      type: "array",
      items: {
        type: "object",
        properties: {
          owner: { type: "string", description: "Person responsible" },
          task: { type: "string", description: "Task description" },
          due: { type: "string", description: "Due date (ISO 8601 or 'ASAP')" },
        },
        required: ["owner", "task", "due"],
      },
      description: "Action items with owners and due dates",
    },
    risks: {
      type: "array",
      items: { type: "string" },
      description: "Identified risks or concerns",
    },
    openQuestions: {
      type: "array",
      items: { type: "string" },
      description: "Unresolved questions or topics for follow-up",
    },
    shortSummary: {
      type: "string",
      description: "One-paragraph executive summary (2-3 sentences)",
    },
    detailedSummary: {
      type: "string",
      description: "Comprehensive summary with key discussion points",
    },
  },
  required: [
    "title",
    "dateTime",
    "participants",
    "decisions",
    "actionItems",
    "risks",
    "openQuestions",
    "shortSummary",
    "detailedSummary",
  ],
};
