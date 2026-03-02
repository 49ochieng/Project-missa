import { MessageRecord } from "./types";
import {
  MeetingRecord,
  MeetingParticipantRecord,
  TranscriptChunkRecord,
  MeetingTranscriptResult,
  CreateMeetingInput,
  UpdateMeetingStatusInput,
  AppendTranscriptChunkInput,
} from "./meetingTypes";

/**
 * Abstract database interface that both SQLite and MSSQL implementations follow
 */
export interface IDatabase {
  initialize(): Promise<void>;
  clearAll(): void | Promise<void>;
  get(conversationId: string): MessageRecord[] | Promise<MessageRecord[]>;
  getMessagesByTimeRange(
    conversationId: string,
    startTime: string,
    endTime: string
  ): MessageRecord[] | Promise<MessageRecord[]>;
  getRecentMessages(
    conversationId: string,
    limit?: number
  ): MessageRecord[] | Promise<MessageRecord[]>;
  clearConversation(conversationId: string): void | Promise<void>;
  addMessages(messages: MessageRecord[]): void | Promise<void>;
  countMessages(conversationId: string): number | Promise<number>;
  clearAllMessages(): void | Promise<void>;
  getFilteredMessages(
    conversationId: string,
    keywords: string[],
    startTime: string,
    endTime: string,
    participants?: string[],
    maxResults?: number
  ): MessageRecord[] | Promise<MessageRecord[]>;
  recordFeedback(
    replyToId: string,
    reaction: "like" | "dislike" | string,
    feedbackJson?: unknown
  ): boolean | Promise<boolean>;
  close(): void | Promise<void>;

  // ============================================================================
  // MEETING STORAGE METHODS
  // ============================================================================

  /**
   * Create or update a meeting record
   */
  upsertMeeting(meeting: CreateMeetingInput): Promise<MeetingRecord>;

  /**
   * Update meeting status (joining, recording, ended, failed)
   */
  updateMeetingStatus(input: UpdateMeetingStatusInput): Promise<boolean>;

  /**
   * Get a meeting by ID
   */
  getMeeting(meetingId: string): Promise<MeetingRecord | null>;

  /**
   * Upsert participants for a meeting
   */
  upsertParticipants(meetingId: string, participants: MeetingParticipantRecord[]): Promise<void>;

  /**
   * Get participants for a meeting
   */
  getParticipants(meetingId: string): Promise<MeetingParticipantRecord[]>;

  /**
   * Append a transcript chunk for real-time transcription
   */
  appendTranscriptChunk(input: AppendTranscriptChunkInput): Promise<TranscriptChunkRecord>;

  /**
   * Get all transcript chunks for a meeting
   */
  getTranscriptChunks(meetingId: string): Promise<TranscriptChunkRecord[]>;

  /**
   * Get full transcript with meeting info and participants
   */
  getTranscriptByMeetingId(meetingId: string): Promise<MeetingTranscriptResult | null>;

  /**
   * Get meetings by status (e.g., all active recordings)
   */
  getMeetingsByStatus(status: string): Promise<MeetingRecord[]>;
}
