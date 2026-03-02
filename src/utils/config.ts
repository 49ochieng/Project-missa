import { ILogger } from "@microsoft/teams.common";

// Configuration for AI models used by different capabilities
export interface ModelConfig {
  model: string;
  apiKey: string;
  endpoint: string;
  apiVersion: string;
}

// Database configuration
export interface DatabaseConfig {
  type: "sqlite" | "mssql";
  connectionString?: string;
  server?: string;
  database?: string;
  username?: string;
  password?: string;
  sqlitePath?: string;
}

// Database configuration
export const DATABASE_CONFIG: DatabaseConfig = {
  type: process.env.RUNNING_ON_AZURE === "1" ? "mssql" : "sqlite",
  connectionString: process.env.SQL_CONNECTION_STRING,
  server: process.env.SQL_SERVER,
  database: process.env.SQL_DATABASE,
  username: process.env.SQL_USERNAME,
  password: process.env.SQL_PASSWORD,
  sqlitePath: process.env.CONVERSATIONS_DB_PATH,
};

// Model configurations for different capabilities
export const AI_MODELS = {
  // Manager Capability - Uses lighter, faster model for routing decisions
  MANAGER: {
    model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o-mini",
    apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Summarizer Capability - Uses more capable model for complex analysis
  SUMMARIZER: {
    model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Action Items Capability - Uses capable model for analysis and task management
  ACTION_ITEMS: {
    model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Search Capability - Uses capable model for semantic search and deep linking
  SEARCH: {
    model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Meeting Notes Capability - Uses capable model for transcript analysis and structured summaries
  MEETING_NOTES: {
    model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Default model configuration (fallback)
  DEFAULT: {
    model: process.env.AOAI_MODEL || process.env.AZURE_OPENAI_DEPLOYMENT_NAME || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,
};

// Helper function to get model config for a specific capability
export function getModelConfig(capabilityType: string): ModelConfig {
  switch (capabilityType.toLowerCase()) {
    case "manager":
      return AI_MODELS.MANAGER;
    case "summarizer":
      return AI_MODELS.SUMMARIZER;
    case "actionitems":
      return AI_MODELS.ACTION_ITEMS;
    case "search":
      return AI_MODELS.SEARCH;
    case "meetingnotes":
      return AI_MODELS.MEETING_NOTES;
    default:
      return AI_MODELS.DEFAULT;
  }
}

// Environment validation
export function validateEnvironment(logger: ILogger): void {
  const hasAoaiKey = process.env.AOAI_API_KEY || process.env.SECRET_AZURE_OPENAI_API_KEY;
  const hasAoaiEndpoint = process.env.AOAI_ENDPOINT || process.env.AZURE_OPENAI_ENDPOINT;

  if (!hasAoaiKey || !hasAoaiEndpoint) {
    throw new Error(`Missing required environment variables: ${[
      !hasAoaiKey && "AOAI_API_KEY / SECRET_AZURE_OPENAI_API_KEY",
      !hasAoaiEndpoint && "AOAI_ENDPOINT / AZURE_OPENAI_ENDPOINT",
    ].filter(Boolean).join(", ")}`);
  }

  // Validate database configuration
  if (DATABASE_CONFIG.type === "mssql") {
    const sqlRequiredVars = ["SQL_CONNECTION_STRING"];
    const sqlMissing = sqlRequiredVars.filter((envVar) => !process.env[envVar]);
    if (sqlMissing.length > 0) {
      logger.warn(
        `SQL Server configuration incomplete. Missing: ${sqlMissing.join(
          ", "
        )}. Falling back to SQLite.`
      );
      DATABASE_CONFIG.type = "sqlite";
    } else {
      logger.debug("✅ SQL Server configuration validated");
    }
  }

  logger.debug(`📦 Using database: ${DATABASE_CONFIG.type}`);
  logger.debug("✅ Environment validation passed");
}

// Model configuration logging
export function logModelConfigs(logger: ILogger): void {
  logger.debug("🔧 AI Model Configuration:");
  logger.debug(`  Manager Capability: ${AI_MODELS.MANAGER.model}`);
  logger.debug(`  Summarizer Capability: ${AI_MODELS.SUMMARIZER.model}`);
  logger.debug(`  Action Items Capability: ${AI_MODELS.ACTION_ITEMS.model}`);
  logger.debug(`  Search Capability: ${AI_MODELS.SEARCH.model}`);
  logger.debug(`  Meeting Notes Capability: ${AI_MODELS.MEETING_NOTES.model}`);
  logger.debug(`  Default Model: ${AI_MODELS.DEFAULT.model}`);
}

// Application configuration for runtime
export interface AppConfig {
  // Bot configuration
  botEndpoint: string;
  
  // Meeting media bot communication
  meetingMediaBotUrl: string;
  meetingMediaBotSharedSecret: string;
  
  // Database
  databaseType: "sqlite" | "mssql";
  
  // Azure Speech (optional, for fallback)
  speechKey?: string;
  speechRegion?: string;
}

let appConfigInstance: AppConfig | null = null;

/**
 * Load and validate application configuration
 */
export function loadConfig(logger?: ILogger): AppConfig {
  if (appConfigInstance) {
    return appConfigInstance;
  }

  const config: AppConfig = {
    botEndpoint: process.env.BOT_ENDPOINT || `http://localhost:${process.env.PORT || 3978}`,
    meetingMediaBotUrl: process.env.MEETING_MEDIA_BOT_URL || "http://localhost:4000",
    meetingMediaBotSharedSecret: process.env.MEETING_MEDIA_BOT_SHARED_SECRET || "dev-secret",
    databaseType: DATABASE_CONFIG.type,
    speechKey: process.env.AZURE_SPEECH_KEY,
    speechRegion: process.env.AZURE_SPEECH_REGION,
  };

  logger?.debug("📋 App Configuration loaded:");
  logger?.debug(`  Bot Endpoint: ${config.botEndpoint}`);
  logger?.debug(`  Meeting Media Bot URL: ${config.meetingMediaBotUrl}`);
  logger?.debug(`  Database Type: ${config.databaseType}`);

  appConfigInstance = config;
  return config;
}

/**
 * Get the current app config (must call loadConfig first)
 */
export function getAppConfig(): AppConfig {
  if (!appConfigInstance) {
    throw new Error("App config not loaded. Call loadConfig() first.");
  }
  return appConfigInstance;
}
