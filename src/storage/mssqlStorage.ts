import { ILogger } from "@microsoft/teams.common";
import * as mssql from "mssql";
import { DatabaseConfig } from "../utils/config";
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

export class MssqlKVStore implements IDatabase {
  private pool: mssql.ConnectionPool | null = null;
  private isInitialized = false;

  constructor(private logger: ILogger, private config: DatabaseConfig) {}

  async initialize(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Use connection string directly or individual config properties
      let sqlConfig: mssql.config;

      if (this.config.connectionString) {
        // When using connection string, pass it directly to ConnectionPool
        this.pool = new mssql.ConnectionPool(this.config.connectionString);
      } else {
        // Use individual config properties
        sqlConfig = {
          server: this.config.server!,
          database: this.config.database!,
          user: this.config.username!,
          password: this.config.password!,
          options: {
            encrypt: true,
            trustServerCertificate: false,
          },
        };
        this.pool = new mssql.ConnectionPool(sqlConfig);
      }
      await this.pool.connect();
      await this.initializeDatabase();
      this.isInitialized = true;
      this.logger.debug("✅ Connected to MSSQL database");
    } catch (error) {
      this.logger.error("❌ Error connecting to MSSQL database:", error);
      throw error;
    }
  }

  private async initializeDatabase(): Promise<void> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      // Create conversations table
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='conversations' AND xtype='U')
        BEGIN
          CREATE TABLE conversations (
            id INT IDENTITY(1,1) PRIMARY KEY,
            conversation_id NVARCHAR(255) NOT NULL,
            role NVARCHAR(50) NOT NULL,
            name NVARCHAR(255) NOT NULL,
            content NVARCHAR(MAX) NOT NULL,
            activity_id NVARCHAR(255) NOT NULL,
            timestamp NVARCHAR(50) NOT NULL,
            blob NVARCHAR(MAX) NOT NULL
          )
        END
      `);

      // Create index on conversation_id
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_conversation_id' AND object_id = OBJECT_ID('conversations'))
        BEGIN
          CREATE INDEX idx_conversation_id ON conversations(conversation_id)
        END
      `);

      // Create feedback table
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='feedback' AND xtype='U')
        BEGIN
          CREATE TABLE feedback (
            id INT IDENTITY(1,1) PRIMARY KEY,
            reply_to_id NVARCHAR(255) NOT NULL,
            reaction NVARCHAR(50) NOT NULL CHECK (reaction IN ('like','dislike')),
            feedback NVARCHAR(MAX),
            created_at DATETIME NOT NULL DEFAULT GETDATE()
          )
        END
      `);

      // ============================================================================
      // MEETING TABLES
      // ============================================================================

      // Create meetings table
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='meetings' AND xtype='U')
        BEGIN
          CREATE TABLE meetings (
            meeting_id NVARCHAR(255) PRIMARY KEY,
            join_url NVARCHAR(2048) NOT NULL,
            organizer_aad_id NVARCHAR(255) NOT NULL,
            organizer_display_name NVARCHAR(255),
            organizer_email NVARCHAR(255),
            started_at NVARCHAR(50),
            ended_at NVARCHAR(50),
            status NVARCHAR(50) NOT NULL DEFAULT 'joining',
            title NVARCHAR(500),
            conversation_id NVARCHAR(255),
            requested_by_aad_id NVARCHAR(255),
            created_at DATETIME NOT NULL DEFAULT GETDATE(),
            updated_at DATETIME NOT NULL DEFAULT GETDATE()
          )
        END
      `);

      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_meetings_status' AND object_id = OBJECT_ID('meetings'))
        BEGIN
          CREATE INDEX idx_meetings_status ON meetings(status)
        END
      `);

      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_meetings_conversation' AND object_id = OBJECT_ID('meetings'))
        BEGIN
          CREATE INDEX idx_meetings_conversation ON meetings(conversation_id)
        END
      `);

      // Create meeting_participants table
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='meeting_participants' AND xtype='U')
        BEGIN
          CREATE TABLE meeting_participants (
            id INT IDENTITY(1,1) PRIMARY KEY,
            meeting_id NVARCHAR(255) NOT NULL,
            participant_aad_id NVARCHAR(255) NOT NULL,
            display_name NVARCHAR(255) NOT NULL,
            email NVARCHAR(255),
            joined_at NVARCHAR(50),
            left_at NVARCHAR(50),
            CONSTRAINT FK_participants_meeting FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id),
            CONSTRAINT UQ_participant UNIQUE (meeting_id, participant_aad_id)
          )
        END
      `);

      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_participants_meeting' AND object_id = OBJECT_ID('meeting_participants'))
        BEGIN
          CREATE INDEX idx_participants_meeting ON meeting_participants(meeting_id)
        END
      `);

      // Create transcript_chunks table
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='transcript_chunks' AND xtype='U')
        BEGIN
          CREATE TABLE transcript_chunks (
            id INT IDENTITY(1,1) PRIMARY KEY,
            meeting_id NVARCHAR(255) NOT NULL,
            timestamp_utc NVARCHAR(50) NOT NULL,
            speaker NVARCHAR(255) NOT NULL,
            speaker_aad_id NVARCHAR(255),
            text NVARCHAR(MAX) NOT NULL,
            confidence FLOAT NOT NULL DEFAULT 0.0,
            source NVARCHAR(50) NOT NULL CHECK (source IN ('speech', 'graphTranscript')),
            sequence_number INT,
            created_at DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT FK_chunks_meeting FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
          )
        END
      `);

      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_chunks_meeting' AND object_id = OBJECT_ID('transcript_chunks'))
        BEGIN
          CREATE INDEX idx_chunks_meeting ON transcript_chunks(meeting_id)
        END
      `);

      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_chunks_timestamp' AND object_id = OBJECT_ID('transcript_chunks'))
        BEGIN
          CREATE INDEX idx_chunks_timestamp ON transcript_chunks(timestamp_utc)
        END
      `);

      // Create meeting_summaries table
      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='meeting_summaries' AND xtype='U')
        BEGIN
          CREATE TABLE meeting_summaries (
            id INT IDENTITY(1,1) PRIMARY KEY,
            meeting_id NVARCHAR(255) NOT NULL,
            title NVARCHAR(500) NOT NULL,
            summary_json NVARCHAR(MAX) NOT NULL,
            short_summary NVARCHAR(MAX) NOT NULL,
            generated_at DATETIME NOT NULL DEFAULT GETDATE(),
            generated_by_aad_id NVARCHAR(255),
            CONSTRAINT FK_summaries_meeting FOREIGN KEY (meeting_id) REFERENCES meetings(meeting_id)
          )
        END
      `);

      await this.pool.request().query(`
        IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_summaries_meeting' AND object_id = OBJECT_ID('meeting_summaries'))
        BEGIN
          CREATE INDEX idx_summaries_meeting ON meeting_summaries(meeting_id)
        END
      `);

      this.logger.debug("✅ Database tables initialized");
    } catch (error) {
      this.logger.error("❌ Error initializing database tables:", error);
      throw error;
    }
  }

  async clearAll(): Promise<void> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      await this.pool.request().query("DELETE FROM conversations");
      this.logger.debug("🧹 Cleared all conversations from MSSQL store.");
    } catch (error) {
      this.logger.error("❌ Error clearing all conversations:", error);
      throw error;
    }
  }

  async get(conversationId: string): Promise<MessageRecord[]> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("conversationId", mssql.NVarChar, conversationId)
        .query(
          "SELECT blob FROM conversations WHERE conversation_id = @conversationId ORDER BY timestamp ASC"
        );

      return result.recordset.map((row) => JSON.parse(row.blob) as MessageRecord);
    } catch (error) {
      this.logger.error("❌ Error getting messages:", error);
      return [];
    }
  }

  async getMessagesByTimeRange(
    conversationId: string,
    startTime: string,
    endTime: string
  ): Promise<MessageRecord[]> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("conversationId", mssql.NVarChar, conversationId)
        .input("startTime", mssql.NVarChar, startTime)
        .input("endTime", mssql.NVarChar, endTime).query(`
          SELECT blob FROM conversations 
          WHERE conversation_id = @conversationId 
            AND timestamp >= @startTime 
            AND timestamp <= @endTime 
          ORDER BY timestamp ASC
        `);

      return result.recordset.map((row) => JSON.parse(row.blob) as MessageRecord);
    } catch (error) {
      this.logger.error("❌ Error getting messages by time range:", error);
      return [];
    }
  }

  async getRecentMessages(conversationId: string, limit = 10): Promise<MessageRecord[]> {
    const messages = await this.get(conversationId);
    return messages.slice(-limit);
  }

  async clearConversation(conversationId: string): Promise<void> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      await this.pool
        .request()
        .input("conversationId", mssql.NVarChar, conversationId)
        .query("DELETE FROM conversations WHERE conversation_id = @conversationId");
    } catch (error) {
      this.logger.error("❌ Error clearing conversation:", error);
      throw error;
    }
  }

  async addMessages(messages: MessageRecord[]): Promise<void> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const transaction = new mssql.Transaction(this.pool);
      await transaction.begin();

      try {
        for (const message of messages) {
          await transaction
            .request()
            .input("conversationId", mssql.NVarChar, message.conversation_id)
            .input("role", mssql.NVarChar, message.role)
            .input("name", mssql.NVarChar, message.name)
            .input("content", mssql.NVarChar, message.content)
            .input("activityId", mssql.NVarChar, message.activity_id)
            .input("timestamp", mssql.NVarChar, message.timestamp)
            .input("blob", mssql.NVarChar, JSON.stringify(message)).query(`
              INSERT INTO conversations (conversation_id, role, name, content, activity_id, timestamp, blob)
              VALUES (@conversationId, @role, @name, @content, @activityId, @timestamp, @blob)
            `);
        }

        await transaction.commit();
      } catch (error) {
        await transaction.rollback();
        throw error;
      }
    } catch (error) {
      this.logger.error("❌ Error adding messages:", error);
      throw error;
    }
  }

  async countMessages(conversationId: string): Promise<number> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("conversationId", mssql.NVarChar, conversationId)
        .query(
          "SELECT COUNT(*) as count FROM conversations WHERE conversation_id = @conversationId"
        );

      return result.recordset[0].count;
    } catch (error) {
      this.logger.error("❌ Error counting messages:", error);
      return 0;
    }
  }

  async clearAllMessages(): Promise<void> {
    await this.clearAll();
  }

  async getFilteredMessages(
    conversationId: string,
    keywords: string[],
    startTime: string,
    endTime: string,
    participants?: string[],
    maxResults = 5
  ): Promise<MessageRecord[]> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const request = this.pool.request();

      // Build dynamic query
      let whereClause =
        "conversation_id = @conversationId AND timestamp >= @startTime AND timestamp <= @endTime";
      request.input("conversationId", mssql.NVarChar, conversationId);
      request.input("startTime", mssql.NVarChar, startTime);
      request.input("endTime", mssql.NVarChar, endTime);
      request.input("maxResults", mssql.Int, maxResults);

      // Add keyword filters
      if (keywords.length > 0) {
        const keywordConditions = keywords
          .map((_, index) => {
            request.input(`keyword${index}`, mssql.NVarChar, `%${keywords[index].toLowerCase()}%`);
            return `content LIKE @keyword${index}`;
          })
          .join(" OR ");
        whereClause += ` AND (${keywordConditions})`;
      }

      // Add participant filters
      if (participants && participants.length > 0) {
        const participantConditions = participants
          .map((_, index) => {
            request.input(
              `participant${index}`,
              mssql.NVarChar,
              `%${participants[index].toLowerCase()}%`
            );
            return `name LIKE @participant${index}`;
          })
          .join(" OR ");
        whereClause += ` AND (${participantConditions})`;
      }

      const query = `
        SELECT TOP (@maxResults) blob FROM conversations
        WHERE ${whereClause}
        ORDER BY timestamp DESC
      `;

      const result = await request.query(query);
      return result.recordset.map((row) => JSON.parse(row.blob) as MessageRecord);
    } catch (error) {
      this.logger.error("❌ Error getting filtered messages:", error);
      return [];
    }
  }

  async recordFeedback(
    replyToId: string,
    reaction: "like" | "dislike" | string,
    feedbackJson?: unknown
  ): Promise<boolean> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      await this.pool
        .request()
        .input("replyToId", mssql.NVarChar, replyToId)
        .input("reaction", mssql.NVarChar, reaction)
        .input("feedback", mssql.NVarChar, feedbackJson ? JSON.stringify(feedbackJson) : null)
        .query(`
          INSERT INTO feedback (reply_to_id, reaction, feedback)
          VALUES (@replyToId, @reaction, @feedback)
        `);

      return true;
    } catch (error) {
      this.logger.error("❌ Error recording feedback:", error);
      return false;
    }
  }

  // ============================================================================
  // MEETING STORAGE METHODS
  // ============================================================================

  async upsertMeeting(input: CreateMeetingInput): Promise<MeetingRecord> {
    if (!this.pool) throw new Error("Database not connected");

    const now = new Date().toISOString();

    try {
      await this.pool
        .request()
        .input("meetingId", mssql.NVarChar, input.meetingId)
        .input("joinUrl", mssql.NVarChar, input.joinUrl)
        .input("organizerAadId", mssql.NVarChar, input.organizerAadId)
        .input("organizerDisplayName", mssql.NVarChar, input.organizerDisplayName || null)
        .input("organizerEmail", mssql.NVarChar, input.organizerEmail || null)
        .input("startedAt", mssql.NVarChar, now)
        .input("title", mssql.NVarChar, input.title || null)
        .input("conversationId", mssql.NVarChar, input.conversationId || null)
        .input("requestedByAadId", mssql.NVarChar, input.requestedByAadId || null)
        .input("now", mssql.DateTime, new Date())
        .query(`
          MERGE meetings AS target
          USING (SELECT @meetingId AS meeting_id) AS source
          ON (target.meeting_id = source.meeting_id)
          WHEN MATCHED THEN
            UPDATE SET
              join_url = @joinUrl,
              organizer_aad_id = @organizerAadId,
              organizer_display_name = @organizerDisplayName,
              organizer_email = @organizerEmail,
              title = @title,
              conversation_id = @conversationId,
              requested_by_aad_id = @requestedByAadId,
              updated_at = @now
          WHEN NOT MATCHED THEN
            INSERT (meeting_id, join_url, organizer_aad_id, organizer_display_name, organizer_email, started_at, status, title, conversation_id, requested_by_aad_id, created_at, updated_at)
            VALUES (@meetingId, @joinUrl, @organizerAadId, @organizerDisplayName, @organizerEmail, @startedAt, 'joining', @title, @conversationId, @requestedByAadId, @now, @now);
        `);

      const meeting = await this.getMeeting(input.meetingId);
      if (!meeting) {
        throw new Error(`Failed to retrieve meeting after upsert: ${input.meetingId}`);
      }

      this.logger.debug(`✅ Upserted meeting: ${input.meetingId}`);
      return meeting;
    } catch (error) {
      this.logger.error("❌ upsertMeeting error:", error);
      throw error;
    }
  }

  async updateMeetingStatus(input: UpdateMeetingStatusInput): Promise<boolean> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const request = this.pool
        .request()
        .input("meetingId", mssql.NVarChar, input.meetingId)
        .input("status", mssql.NVarChar, input.status)
        .input("now", mssql.DateTime, new Date());

      if (input.endedAt) {
        request.input("endedAt", mssql.NVarChar, input.endedAt);
        await request.query(`
          UPDATE meetings 
          SET status = @status, ended_at = @endedAt, updated_at = @now
          WHERE meeting_id = @meetingId
        `);
      } else {
        await request.query(`
          UPDATE meetings 
          SET status = @status, updated_at = @now
          WHERE meeting_id = @meetingId
        `);
      }

      this.logger.debug(`✅ Updated meeting ${input.meetingId} status to: ${input.status}`);
      return true;
    } catch (error) {
      this.logger.error("❌ updateMeetingStatus error:", error);
      return false;
    }
  }

  async getMeeting(meetingId: string): Promise<MeetingRecord | null> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("meetingId", mssql.NVarChar, meetingId)
        .query(`SELECT * FROM meetings WHERE meeting_id = @meetingId`);

      if (result.recordset.length === 0) return null;

      const row = result.recordset[0];
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
        createdAt: row.created_at?.toISOString() || new Date().toISOString(),
        updatedAt: row.updated_at?.toISOString() || new Date().toISOString(),
      };
    } catch (error) {
      this.logger.error("❌ getMeeting error:", error);
      return null;
    }
  }

  async upsertParticipants(meetingId: string, participants: MeetingParticipantRecord[]): Promise<void> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      for (const participant of participants) {
        await this.pool
          .request()
          .input("meetingId", mssql.NVarChar, meetingId)
          .input("participantAadId", mssql.NVarChar, participant.participantAadId)
          .input("displayName", mssql.NVarChar, participant.displayName)
          .input("email", mssql.NVarChar, participant.email || null)
          .input("joinedAt", mssql.NVarChar, participant.joinedAt || null)
          .input("leftAt", mssql.NVarChar, participant.leftAt || null)
          .query(`
            MERGE meeting_participants AS target
            USING (SELECT @meetingId AS meeting_id, @participantAadId AS participant_aad_id) AS source
            ON (target.meeting_id = source.meeting_id AND target.participant_aad_id = source.participant_aad_id)
            WHEN MATCHED THEN
              UPDATE SET
                display_name = @displayName,
                email = @email,
                joined_at = COALESCE(@joinedAt, target.joined_at),
                left_at = @leftAt
            WHEN NOT MATCHED THEN
              INSERT (meeting_id, participant_aad_id, display_name, email, joined_at, left_at)
              VALUES (@meetingId, @participantAadId, @displayName, @email, @joinedAt, @leftAt);
          `);
      }

      this.logger.debug(`✅ Upserted ${participants.length} participants for meeting: ${meetingId}`);
    } catch (error) {
      this.logger.error("❌ upsertParticipants error:", error);
      throw error;
    }
  }

  async getParticipants(meetingId: string): Promise<MeetingParticipantRecord[]> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("meetingId", mssql.NVarChar, meetingId)
        .query(`SELECT * FROM meeting_participants WHERE meeting_id = @meetingId`);

      return result.recordset.map(row => ({
        id: row.id,
        meetingId: row.meeting_id,
        participantAadId: row.participant_aad_id,
        displayName: row.display_name,
        email: row.email || undefined,
        joinedAt: row.joined_at || undefined,
        leftAt: row.left_at || undefined,
      }));
    } catch (error) {
      this.logger.error("❌ getParticipants error:", error);
      return [];
    }
  }

  async appendTranscriptChunk(input: AppendTranscriptChunkInput): Promise<TranscriptChunkRecord> {
    if (!this.pool) throw new Error("Database not connected");

    const now = new Date().toISOString();

    try {
      // Get next sequence number
      const seqResult = await this.pool
        .request()
        .input("meetingId", mssql.NVarChar, input.meetingId)
        .query(`SELECT ISNULL(MAX(sequence_number), -1) as max_seq FROM transcript_chunks WHERE meeting_id = @meetingId`);

      const sequenceNumber = (seqResult.recordset[0].max_seq ?? -1) + 1;

      const insertResult = await this.pool
        .request()
        .input("meetingId", mssql.NVarChar, input.meetingId)
        .input("timestampUtc", mssql.NVarChar, input.timestampUtc)
        .input("speaker", mssql.NVarChar, input.speaker)
        .input("speakerAadId", mssql.NVarChar, input.speakerAadId || null)
        .input("text", mssql.NVarChar, input.text)
        .input("confidence", mssql.Float, input.confidence)
        .input("source", mssql.NVarChar, input.source)
        .input("sequenceNumber", mssql.Int, sequenceNumber)
        .query(`
          INSERT INTO transcript_chunks (meeting_id, timestamp_utc, speaker, speaker_aad_id, text, confidence, source, sequence_number)
          OUTPUT INSERTED.id
          VALUES (@meetingId, @timestampUtc, @speaker, @speakerAadId, @text, @confidence, @source, @sequenceNumber)
        `);

      const insertedId = insertResult.recordset[0].id;

      this.logger.debug(`✅ Appended transcript chunk for meeting: ${input.meetingId}, seq: ${sequenceNumber}`);

      return {
        id: insertedId,
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
    } catch (error) {
      this.logger.error("❌ appendTranscriptChunk error:", error);
      throw error;
    }
  }

  async getTranscriptChunks(meetingId: string): Promise<TranscriptChunkRecord[]> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("meetingId", mssql.NVarChar, meetingId)
        .query(`SELECT * FROM transcript_chunks WHERE meeting_id = @meetingId ORDER BY sequence_number ASC, timestamp_utc ASC`);

      return result.recordset.map(row => ({
        id: row.id,
        meetingId: row.meeting_id,
        timestampUtc: row.timestamp_utc,
        speaker: row.speaker,
        speakerAadId: row.speaker_aad_id || undefined,
        text: row.text,
        confidence: row.confidence,
        source: row.source as TranscriptChunkRecord["source"],
        sequenceNumber: row.sequence_number || undefined,
        createdAt: row.created_at?.toISOString() || new Date().toISOString(),
      }));
    } catch (error) {
      this.logger.error("❌ getTranscriptChunks error:", error);
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
    } catch (error) {
      this.logger.error("❌ getTranscriptByMeetingId error:", error);
      return null;
    }
  }

  async getMeetingsByStatus(status: string): Promise<MeetingRecord[]> {
    if (!this.pool) throw new Error("Database not connected");

    try {
      const result = await this.pool
        .request()
        .input("status", mssql.NVarChar, status)
        .query(`SELECT * FROM meetings WHERE status = @status`);

      return result.recordset.map(row => ({
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
        createdAt: row.created_at?.toISOString() || new Date().toISOString(),
        updatedAt: row.updated_at?.toISOString() || new Date().toISOString(),
      }));
    } catch (error) {
      this.logger.error("❌ getMeetingsByStatus error:", error);
      return [];
    }
  }

  async close(): Promise<void> {
    if (this.pool) {
      await this.pool.close();
      this.pool = null;
      this.isInitialized = false;
      this.logger.debug("🔌 Closed MSSQL database connection");
    }
  }
}
