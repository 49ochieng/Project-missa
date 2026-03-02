import { ILogger } from "@microsoft/teams.common";
import Database from "better-sqlite3";
import path from "node:path";
import { IDatabase } from "./database";
import {
  MeetingRecord,
  MeetingParticipantRecord,
  TranscriptChunkRecord,
  MeetingTranscriptResult,
  CreateMeetingInput,
  UpdateMeetingStatusInput,
  AppendTranscriptChunkInput,
} from "./meetingTypes";
import { MessageRecord } from "./types";

export class SqliteKVStore implements IDatabase {
  private db: Database.Database;

  constructor(private logger: ILogger, dbPath?: string) {
    // Use environment variable if set, otherwise use provided dbPath, otherwise use default relative to project root
    const resolvedDbPath = process.env.CONVERSATIONS_DB_PATH
      ? path.resolve(process.env.CONVERSATIONS_DB_PATH)
      : dbPath
      ? dbPath
      : path.resolve(__dirname, "../../src/storage/conversations.db");
    this.db = new Database(resolvedDbPath);
    this.initializeDatabase();
  }

  async initialize(): Promise<void> {
    // SQLite initialization is done in constructor, so this is a no-op for compatibility
    return Promise.resolve();
  }
  private initializeDatabase(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS conversations (
        conversation_id TEXT NOT NULL,
        role TEXT NOT NULL,
        name TEXT NOT NULL,
        content TEXT NOT NULL,
        activity_id TEXT NOT NULL,
        timestamp TEXT NOT NULL,
        blob TEXT NOT NULL
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_conversation_id ON conversations(conversation_id);
    `);
    this.db.exec(`
    CREATE TABLE IF NOT EXISTS feedback (
    reply_to_id  TEXT    NOT NULL,                -- the Teams message ID you replied to
    reaction     TEXT    NOT NULL CHECK (reaction IN ('like','dislike')),
    feedback     TEXT,                           -- JSON or plain text
    created_at   TEXT    NOT NULL DEFAULT (CURRENT_TIMESTAMP)
  );
    `);

    // ============================================================================
    // MEETING TABLES
    // ============================================================================
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meetings (
        meeting_id TEXT PRIMARY KEY,
        join_url TEXT NOT NULL,
        organizer_aad_id TEXT NOT NULL,
        organizer_display_name TEXT,
        organizer_email TEXT,
        started_at TEXT,
        ended_at TEXT,
        status TEXT NOT NULL DEFAULT 'joining',
        title TEXT,
        conversation_id TEXT,
        requested_by_aad_id TEXT,
        created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
        updated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_meetings_status ON meetings(status);
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_meetings_conversation ON meetings(conversation_id);
    `);

    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meeting_participants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        meeting_id TEXT NOT NULL,
        participant_aad_id TEXT NOT NULL,
        display_name TEXT NOT NULL,
        email TEXT,
        joined_at TEXT,
        left_at TEXT,
        FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id),
        UNIQUE(meeting_id, participant_aad_id)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_participants_meeting ON meeting_participants(meeting_id);
    `);

    this.db.exec(`
      CREATE TABLE IF NOT EXISTS transcript_chunks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        meeting_id TEXT NOT NULL,
        timestamp_utc TEXT NOT NULL,
        speaker TEXT NOT NULL,
        speaker_aad_id TEXT,
        text TEXT NOT NULL,
        confidence REAL NOT NULL DEFAULT 0.0,
        source TEXT NOT NULL CHECK (source IN ('speech', 'graphTranscript')),
        sequence_number INTEGER,
        created_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
        FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_chunks_meeting ON transcript_chunks(meeting_id);
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_chunks_timestamp ON transcript_chunks(timestamp_utc);
    `);

    this.db.exec(`
      CREATE TABLE IF NOT EXISTS meeting_summaries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        meeting_id TEXT NOT NULL,
        title TEXT NOT NULL,
        summary_json TEXT NOT NULL,
        short_summary TEXT NOT NULL,
        generated_at TEXT NOT NULL DEFAULT (CURRENT_TIMESTAMP),
        generated_by_aad_id TEXT,
        FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
      )
    `);
    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_summaries_meeting ON meeting_summaries(meeting_id);
    `);
  }

  clearAll(): void {
    this.db.exec("DELETE FROM conversations; VACUUM;");
    this.logger.debug("🧹 Cleared all conversations from SQLite store.");
  }

  get(conversationId: string): MessageRecord[] {
    const stmt = this.db.prepare<unknown[], { blob: string }>(
      "SELECT blob FROM conversations WHERE conversation_id = ? ORDER BY timestamp ASC"
    );
    return stmt.all(conversationId).map((row) => JSON.parse(row.blob) as MessageRecord);
  }

  getMessagesByTimeRange(
    conversationId: string,
    startTime: string,
    endTime: string
  ): MessageRecord[] {
    const stmt = this.db.prepare<unknown[], { blob: string }>(
      "SELECT blob FROM conversations WHERE conversation_id = ? AND timestamp >= ? AND timestamp <= ? ORDER BY timestamp ASC"
    );
    return stmt
      .all([conversationId, startTime, endTime])
      .map((row) => JSON.parse(row.blob) as MessageRecord);
  }

  getRecentMessages(conversationId: string, limit = 10): MessageRecord[] {
    const messages = this.get(conversationId);
    return messages.slice(-limit);
  }

  clearConversation(conversationId: string): void {
    const stmt = this.db.prepare("DELETE FROM conversations WHERE conversation_id = ?");
    stmt.run(conversationId);
  }

  addMessages(messages: MessageRecord[]): void {
    const stmt = this.db.prepare(
      "INSERT INTO conversations (conversation_id, role, name, content, activity_id, timestamp, blob) VALUES (?, ?, ?, ?, ?, ?, ?)"
    );
    for (const message of messages) {
      stmt.run(
        message.conversation_id,
        message.role,
        message.name,
        message.content,
        message.activity_id,
        message.timestamp,
        JSON.stringify(message)
      );
    }
  }

  countMessages(conversationId: string): number {
    const stmt = this.db.prepare(
      "SELECT COUNT(*) as count FROM conversations WHERE conversation_id = ?"
    );
    const result = stmt.get(conversationId) as { count: number };
    return result.count;
  }

  // Clear all messages for debugging (optional utility method)
  clearAllMessages(): void {
    try {
      const stmt = this.db.prepare("DELETE FROM conversations");
      const result = stmt.run();
      this.logger.debug(
        `🧹 Cleared all conversations from database. Deleted ${result.changes} records.`
      );
    } catch (error) {
      this.logger.error("❌ Error clearing all conversations:", error);
    }
  }

  getFilteredMessages(
    conversationId: string,
    keywords: string[],
    startTime: string,
    endTime: string,
    participants?: string[],
    maxResults?: number
  ): MessageRecord[] {
    const keywordClauses = keywords.map(() => `content LIKE ?`).join(" OR ");
    const participantClauses = participants?.map(() => `name LIKE ?`).join(" OR ");

    // Base where clauses
    const whereClauses = [
      `conversation_id = ?`,
      `timestamp >= ?`,
      `timestamp <= ?`,
      `(${keywordClauses})`,
    ];

    // Values for the prepared statement
    const values: (string | number)[] = [
      conversationId,
      startTime,
      endTime,
      ...keywords.map((k) => `%${k.toLowerCase()}%`),
    ];

    // Add participant filters if present
    if (participants && participants.length > 0) {
      whereClauses.push(`(${participantClauses})`);
      values.push(...participants.map((p) => `%${p.toLowerCase()}%`));
    }

    const limit = maxResults && typeof maxResults === "number" ? maxResults : 5;
    values.push(limit);

    const query = `
  SELECT blob FROM conversations
  WHERE ${whereClauses.join(" AND ")}
  ORDER BY timestamp DESC
  LIMIT ?
`;

    const stmt = this.db.prepare(query);
    const rows = stmt.all(...values) as Array<{ blob: string }>;
    return rows.map((row) => JSON.parse(row.blob) as MessageRecord);
  }
  // ===== FEEDBACK MANAGEMENT =====

  // Initialize feedback record for a message with optional delegated capability
  // Insert one row per submission
  recordFeedback(
    replyToId: string,
    reaction: "like" | "dislike" | string,
    feedbackJson?: unknown
  ): boolean {
    try {
      const stmt = this.db.prepare(`
      INSERT INTO feedback (reply_to_id, reaction, feedback)
      VALUES (?, ?, ?)
    `);
      const result = stmt.run(
        replyToId,
        reaction,
        feedbackJson ? JSON.stringify(feedbackJson) : null
      );
      return result.changes > 0;
    } catch (err) {
      this.logger.error(`❌ recordFeedback error:`, err);
      return false;
    }
  }

  // ============================================================================
  // MEETING STORAGE METHODS
  // ============================================================================

  async upsertMeeting(input: CreateMeetingInput): Promise<MeetingRecord> {
    const now = new Date().toISOString();
    
    try {
      const stmt = this.db.prepare(`
        INSERT INTO meetings (
          meeting_id, join_url, organizer_aad_id, organizer_display_name, organizer_email,
          started_at, status, title, conversation_id, requested_by_aad_id, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, 'joining', ?, ?, ?, ?, ?)
        ON CONFLICT(meeting_id) DO UPDATE SET
          join_url = excluded.join_url,
          organizer_aad_id = excluded.organizer_aad_id,
          organizer_display_name = excluded.organizer_display_name,
          organizer_email = excluded.organizer_email,
          title = excluded.title,
          conversation_id = excluded.conversation_id,
          requested_by_aad_id = excluded.requested_by_aad_id,
          updated_at = excluded.updated_at
      `);

      stmt.run(
        input.meetingId,
        input.joinUrl,
        input.organizerAadId,
        input.organizerDisplayName || null,
        input.organizerEmail || null,
        now,
        input.title || null,
        input.conversationId || null,
        input.requestedByAadId || null,
        now,
        now
      );

      const meeting = await this.getMeeting(input.meetingId);
      if (!meeting) {
        throw new Error(`Failed to retrieve meeting after upsert: ${input.meetingId}`);
      }

      this.logger.debug(`✅ Upserted meeting: ${input.meetingId}`);
      return meeting;
    } catch (err) {
      this.logger.error(`❌ upsertMeeting error:`, err);
      throw err;
    }
  }

  async updateMeetingStatus(input: UpdateMeetingStatusInput): Promise<boolean> {
    const now = new Date().toISOString();

    try {
      let stmt;
      if (input.endedAt) {
        stmt = this.db.prepare(`
          UPDATE meetings SET status = ?, ended_at = ?, updated_at = ?
          WHERE meeting_id = ?
        `);
        stmt.run(input.status, input.endedAt, now, input.meetingId);
      } else {
        stmt = this.db.prepare(`
          UPDATE meetings SET status = ?, updated_at = ?
          WHERE meeting_id = ?
        `);
        stmt.run(input.status, now, input.meetingId);
      }

      this.logger.debug(`✅ Updated meeting ${input.meetingId} status to: ${input.status}`);
      return true;
    } catch (err) {
      this.logger.error(`❌ updateMeetingStatus error:`, err);
      return false;
    }
  }

  async getMeeting(meetingId: string): Promise<MeetingRecord | null> {
    try {
      const stmt = this.db.prepare<unknown[], {
        meeting_id: string;
        join_url: string;
        organizer_aad_id: string;
        organizer_display_name: string | null;
        organizer_email: string | null;
        started_at: string;
        ended_at: string | null;
        status: string;
        title: string | null;
        conversation_id: string | null;
        requested_by_aad_id: string | null;
        created_at: string;
        updated_at: string;
      }>(`SELECT * FROM meetings WHERE meeting_id = ?`);

      const row = stmt.get(meetingId);
      if (!row) return null;

      return {
        meetingId: row.meeting_id,
        joinUrl: row.join_url,
        organizerAadId: row.organizer_aad_id,
        organizerDisplayName: row.organizer_display_name || undefined,
        organizerEmail: row.organizer_email || undefined,
        startedAt: row.started_at,
        endedAt: row.ended_at || undefined,
        status: row.status as MeetingRecord["status"],
        title: row.title || undefined,
        conversationId: row.conversation_id || undefined,
        requestedByAadId: row.requested_by_aad_id || undefined,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
      };
    } catch (err) {
      this.logger.error(`❌ getMeeting error:`, err);
      return null;
    }
  }

  async upsertParticipants(meetingId: string, participants: MeetingParticipantRecord[]): Promise<void> {
    try {
      const stmt = this.db.prepare(`
        INSERT INTO meeting_participants (
          meeting_id, participant_aad_id, display_name, email, joined_at, left_at
        ) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(meeting_id, participant_aad_id) DO UPDATE SET
          display_name = excluded.display_name,
          email = excluded.email,
          joined_at = COALESCE(excluded.joined_at, meeting_participants.joined_at),
          left_at = excluded.left_at
      `);

      for (const participant of participants) {
        stmt.run(
          meetingId,
          participant.participantAadId,
          participant.displayName,
          participant.email || null,
          participant.joinedAt || null,
          participant.leftAt || null
        );
      }

      this.logger.debug(`✅ Upserted ${participants.length} participants for meeting: ${meetingId}`);
    } catch (err) {
      this.logger.error(`❌ upsertParticipants error:`, err);
      throw err;
    }
  }

  async getParticipants(meetingId: string): Promise<MeetingParticipantRecord[]> {
    try {
      const stmt = this.db.prepare<unknown[], {
        id: number;
        meeting_id: string;
        participant_aad_id: string;
        display_name: string;
        email: string | null;
        joined_at: string | null;
        left_at: string | null;
      }>(`SELECT * FROM meeting_participants WHERE meeting_id = ?`);

      const rows = stmt.all(meetingId);
      return rows.map(row => ({
        id: row.id,
        meetingId: row.meeting_id,
        participantAadId: row.participant_aad_id,
        displayName: row.display_name,
        email: row.email || undefined,
        joinedAt: row.joined_at || undefined,
        leftAt: row.left_at || undefined,
      }));
    } catch (err) {
      this.logger.error(`❌ getParticipants error:`, err);
      return [];
    }
  }

  async appendTranscriptChunk(input: AppendTranscriptChunkInput): Promise<TranscriptChunkRecord> {
    const now = new Date().toISOString();

    try {
      // Get the next sequence number
      const seqStmt = this.db.prepare<unknown[], { max_seq: number | null }>(
        `SELECT MAX(sequence_number) as max_seq FROM transcript_chunks WHERE meeting_id = ?`
      );
      const seqResult = seqStmt.get(input.meetingId);
      const sequenceNumber = (seqResult?.max_seq ?? -1) + 1;

      const stmt = this.db.prepare(`
        INSERT INTO transcript_chunks (
          meeting_id, timestamp_utc, speaker, speaker_aad_id, text, confidence, source, sequence_number, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);

      const result = stmt.run(
        input.meetingId,
        input.timestampUtc,
        input.speaker,
        input.speakerAadId || null,
        input.text,
        input.confidence,
        input.source,
        sequenceNumber,
        now
      );

      this.logger.debug(`✅ Appended transcript chunk for meeting: ${input.meetingId}, seq: ${sequenceNumber}`);

      return {
        id: Number(result.lastInsertRowid),
        meetingId: input.meetingId,
        timestampUtc: input.timestampUtc,
        speaker: input.speaker,
        speakerAadId: input.speakerAadId,
        text: input.text,
        confidence: input.confidence,
        source: input.source,
        sequenceNumber,
        createdAt: now,
      };
    } catch (err) {
      this.logger.error(`❌ appendTranscriptChunk error:`, err);
      throw err;
    }
  }

  async getTranscriptChunks(meetingId: string): Promise<TranscriptChunkRecord[]> {
    try {
      const stmt = this.db.prepare<unknown[], {
        id: number;
        meeting_id: string;
        timestamp_utc: string;
        speaker: string;
        speaker_aad_id: string | null;
        text: string;
        confidence: number;
        source: string;
        sequence_number: number | null;
        created_at: string;
      }>(`SELECT * FROM transcript_chunks WHERE meeting_id = ? ORDER BY sequence_number ASC, timestamp_utc ASC`);

      const rows = stmt.all(meetingId);
      return rows.map(row => ({
        id: row.id,
        meetingId: row.meeting_id,
        timestampUtc: row.timestamp_utc,
        speaker: row.speaker,
        speakerAadId: row.speaker_aad_id || undefined,
        text: row.text,
        confidence: row.confidence,
        source: row.source as TranscriptChunkRecord["source"],
        sequenceNumber: row.sequence_number || undefined,
        createdAt: row.created_at,
      }));
    } catch (err) {
      this.logger.error(`❌ getTranscriptChunks error:`, err);
      return [];
    }
  }

  async getTranscriptByMeetingId(meetingId: string): Promise<MeetingTranscriptResult | null> {
    try {
      const meeting = await this.getMeeting(meetingId);
      if (!meeting) {
        this.logger.warn(`Meeting not found: ${meetingId}`);
        return null;
      }

      const participants = await this.getParticipants(meetingId);
      const chunks = await this.getTranscriptChunks(meetingId);

      return {
        meetingId: meeting.meetingId,
        title: meeting.title,
        startedAt: meeting.startedAt,
        endedAt: meeting.endedAt,
        participants,
        chunks,
        totalChunks: chunks.length,
      };
    } catch (err) {
      this.logger.error(`❌ getTranscriptByMeetingId error:`, err);
      return null;
    }
  }

  async getMeetingsByStatus(status: string): Promise<MeetingRecord[]> {
    try {
      const stmt = this.db.prepare<unknown[], {
        meeting_id: string;
        join_url: string;
        organizer_aad_id: string;
        organizer_display_name: string | null;
        organizer_email: string | null;
        started_at: string;
        ended_at: string | null;
        status: string;
        title: string | null;
        conversation_id: string | null;
        requested_by_aad_id: string | null;
        created_at: string;
        updated_at: string;
      }>(`SELECT * FROM meetings WHERE status = ?`);

      const rows = stmt.all(status);
      return rows.map(row => ({
        meetingId: row.meeting_id,
        joinUrl: row.join_url,
        organizerAadId: row.organizer_aad_id,
        organizerDisplayName: row.organizer_display_name || undefined,
        organizerEmail: row.organizer_email || undefined,
        startedAt: row.started_at,
        endedAt: row.ended_at || undefined,
        status: row.status as MeetingRecord["status"],
        title: row.title || undefined,
        conversationId: row.conversation_id || undefined,
        requestedByAadId: row.requested_by_aad_id || undefined,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
      }));
    } catch (err) {
      this.logger.error(`❌ getMeetingsByStatus error:`, err);
      return [];
    }
  }

  close(): void {
    if (this.db) {
      this.db.close();
      this.logger.debug("🔌 Closed SQLite database connection");
    }
  }
}
