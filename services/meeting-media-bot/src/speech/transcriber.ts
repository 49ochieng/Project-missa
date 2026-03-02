/**
 * Azure Speech SDK integration for real-time transcription
 * Handles meeting audio transcription with speaker diarization
 */

import * as sdk from "microsoft-cognitiveservices-speech-sdk";
import { getConfig } from "../config";

export interface TranscriptionChunk {
  text: string;
  speakerId?: string;
  timestamp: Date;
  offsetMs: number;
  durationMs: number;
  isFinal: boolean;
}

export interface TranscriberEvents {
  onTranscriptionChunk: (chunk: TranscriptionChunk) => void;
  onError: (error: Error) => void;
  onSessionStarted: () => void;
  onSessionStopped: () => void;
}

/**
 * Meeting transcriber wrapper
 */
export class MeetingTranscriber {
  private conversationTranscriber: sdk.ConversationTranscriber | null = null;
  private audioConfig: sdk.AudioConfig | null = null;
  private speechConfig: sdk.SpeechConfig | null = null;
  private isRunning = false;
  private callId: string;
  private events: TranscriberEvents;

  constructor(callId: string, events: TranscriberEvents) {
    this.callId = callId;
    this.events = events;
  }

  /**
   * Initialize the transcriber with push stream for audio input
   * Returns the push stream to feed audio data into
   */
  initializePushStream(): sdk.PushAudioInputStream {
    const config = getConfig();

    console.log(`[MeetingTranscriber] Initializing for call ${this.callId}`);

    // Create speech config
    this.speechConfig = sdk.SpeechConfig.fromSubscription(
      config.speechKey,
      config.speechRegion
    );

    // Configure for conversation transcription
    this.speechConfig.speechRecognitionLanguage = "en-US";
    this.speechConfig.setProperty(
      sdk.PropertyId.SpeechServiceConnection_LanguageIdMode,
      "Continuous"
    );

    // Create push stream for receiving audio
    const pushStream = sdk.AudioInputStream.createPushStream(
      sdk.AudioStreamFormat.getWaveFormatPCM(16000, 16, 1) // 16kHz, 16-bit, mono
    );

    this.audioConfig = sdk.AudioConfig.fromStreamInput(pushStream);

    // Create conversation transcriber
    this.conversationTranscriber = new sdk.ConversationTranscriber(
      this.speechConfig,
      this.audioConfig
    );

    // Set up event handlers
    this.setupEventHandlers();

    console.log(`[MeetingTranscriber] Push stream created for call ${this.callId}`);
    return pushStream;
  }

  /**
   * Initialize with microphone input (for testing)
   */
  initializeMicrophone(): void {
    const config = getConfig();

    console.log(`[MeetingTranscriber] Initializing with microphone for call ${this.callId}`);

    this.speechConfig = sdk.SpeechConfig.fromSubscription(
      config.speechKey,
      config.speechRegion
    );

    this.speechConfig.speechRecognitionLanguage = "en-US";

    this.audioConfig = sdk.AudioConfig.fromDefaultMicrophoneInput();

    this.conversationTranscriber = new sdk.ConversationTranscriber(
      this.speechConfig,
      this.audioConfig
    );

    this.setupEventHandlers();
  }

  /**
   * Set up event handlers for transcription events
   */
  private setupEventHandlers(): void {
    if (!this.conversationTranscriber) return;

    // Transcribing event - intermediate results
    this.conversationTranscriber.transcribing = (
      _sender,
      event
    ) => {
      if (event.result.text) {
        const chunk: TranscriptionChunk = {
          text: event.result.text,
          speakerId: event.result.speakerId || undefined,
          timestamp: new Date(),
          offsetMs: event.result.offset / 10000, // Convert from 100ns to ms
          durationMs: event.result.duration / 10000,
          isFinal: false,
        };

        console.log(
          `[MeetingTranscriber] Intermediate: [${chunk.speakerId || "Unknown"}] ${chunk.text}`
        );
      }
    };

    // Transcribed event - final results
    this.conversationTranscriber.transcribed = (
      _sender,
      event
    ) => {
      if (event.result.reason === sdk.ResultReason.RecognizedSpeech && event.result.text) {
        const chunk: TranscriptionChunk = {
          text: event.result.text,
          speakerId: event.result.speakerId || undefined,
          timestamp: new Date(),
          offsetMs: event.result.offset / 10000,
          durationMs: event.result.duration / 10000,
          isFinal: true,
        };

        console.log(
          `[MeetingTranscriber] Final: [${chunk.speakerId || "Unknown"}] ${chunk.text}`
        );

        this.events.onTranscriptionChunk(chunk);
      } else if (event.result.reason === sdk.ResultReason.NoMatch) {
        console.log(`[MeetingTranscriber] No speech recognized`);
      }
    };

    // Session started
    this.conversationTranscriber.sessionStarted = () => {
      console.log(`[MeetingTranscriber] Session started for call ${this.callId}`);
      this.isRunning = true;
      this.events.onSessionStarted();
    };

    // Session stopped
    this.conversationTranscriber.sessionStopped = () => {
      console.log(`[MeetingTranscriber] Session stopped for call ${this.callId}`);
      this.isRunning = false;
      this.events.onSessionStopped();
    };

    // Canceled
    this.conversationTranscriber.canceled = (
      _sender,
      event
    ) => {
      console.error(
        `[MeetingTranscriber] Canceled: ${event.reason}, Error: ${event.errorDetails}`
      );

      if (event.reason === sdk.CancellationReason.Error) {
        this.events.onError(new Error(event.errorDetails || "Transcription canceled"));
      }

      this.isRunning = false;
    };
  }

  /**
   * Start transcription
   */
  async start(): Promise<void> {
    if (!this.conversationTranscriber) {
      throw new Error("Transcriber not initialized. Call initializePushStream() first.");
    }

    if (this.isRunning) {
      console.log(`[MeetingTranscriber] Already running for call ${this.callId}`);
      return;
    }

    console.log(`[MeetingTranscriber] Starting transcription for call ${this.callId}`);

    return new Promise((resolve, reject) => {
      this.conversationTranscriber!.startTranscribingAsync(
        () => {
          console.log(`[MeetingTranscriber] Transcription started for call ${this.callId}`);
          resolve();
        },
        (error: string) => {
          console.error(`[MeetingTranscriber] Failed to start: ${error}`);
          reject(new Error(error));
        }
      );
    });
  }

  /**
   * Stop transcription
   */
  async stop(): Promise<void> {
    if (!this.conversationTranscriber || !this.isRunning) {
      console.log(`[MeetingTranscriber] Not running for call ${this.callId}`);
      return;
    }

    console.log(`[MeetingTranscriber] Stopping transcription for call ${this.callId}`);

    return new Promise((resolve, reject) => {
      this.conversationTranscriber!.stopTranscribingAsync(
        () => {
          console.log(`[MeetingTranscriber] Transcription stopped for call ${this.callId}`);
          this.isRunning = false;
          resolve();
        },
        (error: string) => {
          console.error(`[MeetingTranscriber] Failed to stop: ${error}`);
          reject(new Error(error));
        }
      );
    });
  }

  /**
   * Clean up resources
   */
  dispose(): void {
    console.log(`[MeetingTranscriber] Disposing resources for call ${this.callId}`);

    if (this.conversationTranscriber) {
      this.conversationTranscriber.close();
      this.conversationTranscriber = null;
    }

    if (this.audioConfig) {
      this.audioConfig.close();
      this.audioConfig = null;
    }

    if (this.speechConfig) {
      this.speechConfig.close();
      this.speechConfig = null;
    }

    this.isRunning = false;
  }

  /**
   * Check if transcriber is running
   */
  getIsRunning(): boolean {
    return this.isRunning;
  }
}

/**
 * Active transcribers by call ID
 */
const activeTranscribers = new Map<string, MeetingTranscriber>();

/**
 * Create and register a transcriber for a call
 */
export function createTranscriber(
  callId: string,
  events: TranscriberEvents
): MeetingTranscriber {
  // Clean up existing transcriber if any
  const existing = activeTranscribers.get(callId);
  if (existing) {
    existing.dispose();
  }

  const transcriber = new MeetingTranscriber(callId, events);
  activeTranscribers.set(callId, transcriber);

  return transcriber;
}

/**
 * Get transcriber by call ID
 */
export function getTranscriber(callId: string): MeetingTranscriber | undefined {
  return activeTranscribers.get(callId);
}

/**
 * Remove and dispose transcriber
 */
export function removeTranscriber(callId: string): void {
  const transcriber = activeTranscribers.get(callId);
  if (transcriber) {
    transcriber.dispose();
    activeTranscribers.delete(callId);
  }
}
