/**
 * Meeting-related database types for transcript storage and meeting tracking
 */

/**
 * Meeting status enum representing lifecycle states
 */
export type MeetingStatus = 
  | "joining"      // Bot is attempting to join the meeting
  | "recording"    // Bot has joined and is actively recording/transcribing
  | "ended"        // Meeting has ended normally
  | "failed"       // Bot failed to join or was disconnected
  | "cancelled";   // Meeting capture was cancelled by user

/**
 * Transcript chunk source indicator
 */
export type TranscriptSource = "speech" | "graphTranscript";

/**
 * Meeting record stored in database
 */
export interface MeetingRecord {
  meetingId: string;
  joinUrl: string;
  organizerAadId: string;
  organizerDisplayName?: string;
  organizerEmail?: string;
  startedAt: string;         // ISO 8601 timestamp
  endedAt?: string;          // ISO 8601 timestamp, null if still active
  status: MeetingStatus;
  title?: string;
  conversationId?: string;   // Teams conversation where capture was requested
  requestedByAadId?: string; // User who requested the capture
  createdAt: string;         // ISO 8601 timestamp
  updatedAt: string;         // ISO 8601 timestamp
}

/**
 * Meeting participant record
 */
export interface MeetingParticipantRecord {
  id?: number;
  meetingId: string;
  participantAadId: string;
  displayName: string;
  email?: string;
  joinedAt?: string;         // ISO 8601 timestamp
  leftAt?: string;           // ISO 8601 timestamp
}

/**
 * Transcript chunk record for real-time transcription storage
 */
export interface TranscriptChunkRecord {
  id?: number;
  meetingId: string;
  timestampUtc: string;      // ISO 8601 timestamp of when this chunk was spoken
  speaker: string;           // Speaker name or identifier
  speakerAadId?: string;     // Azure AD object ID of speaker if resolved
  text: string;              // Transcribed text content
  confidence: number;        // Speech recognition confidence (0.0 - 1.0)
  source: TranscriptSource;  // Where this transcript came from
  sequenceNumber?: number;   // Order within the meeting
  createdAt: string;         // ISO 8601 timestamp
}

/**
 * Meeting summary record for storing generated summaries
 */
export interface MeetingSummaryRecord {
  id?: number;
  meetingId: string;
  title: string;
  summaryJson: string;       // Full structured summary as JSON
  shortSummary: string;      // Executive summary
  generatedAt: string;       // ISO 8601 timestamp
  generatedByAadId?: string; // Who requested the summary
}

/**
 * Input for creating a new meeting record
 */
export interface CreateMeetingInput {
  meetingId: string;
  joinUrl: string;
  organizerAadId: string;
  organizerDisplayName?: string;
  organizerEmail?: string;
  title?: string;
  conversationId?: string;
  requestedByAadId?: string;
}

/**
 * Input for updating meeting status
 */
export interface UpdateMeetingStatusInput {
  meetingId: string;
  status: MeetingStatus;
  endedAt?: string;
}

/**
 * Input for appending a transcript chunk
 */
export interface AppendTranscriptChunkInput {
  meetingId: string;
  timestampUtc: string;
  speaker: string;
  speakerAadId?: string;
  text: string;
  confidence: number;
  source: TranscriptSource;
}

/**
 * Combined transcript for a meeting
 */
export interface MeetingTranscriptResult {
  meetingId: string;
  title?: string;
  startedAt?: string;
  endedAt?: string;
  participants: MeetingParticipantRecord[];
  chunks: TranscriptChunkRecord[];
  totalChunks: number;
}
