/**
 * Configuration for meeting-media-bot service
 * Validates required environment variables on startup
 * Never logs secrets
 */

export interface MeetingMediaBotConfig {
  // Azure Authentication
  azureTenantId: string;
  azureClientId: string;
  azureClientSecret: string;

  // Bot Configuration
  botAppId: string;
  botAppPassword: string;
  botEndpoint: string;

  // Microsoft Graph
  graphBaseUrl: string;

  // Azure Speech Service
  speechKey: string;
  speechRegion: string;

  // Project-missa service
  projectMissaUrl: string;
  sharedSecret: string;

  // Server
  port: number;

  // Auto-join (optional)
  autoJoinUserEmail: string | null;
  autoJoinUserObjectId: string | null;
  calendarPollIntervalMs: number;
}

interface ConfigSchema {
  envVars: string[];
  required: boolean;
  isSecret: boolean;
  defaultValue?: string;
}

const configSchema: Record<keyof MeetingMediaBotConfig, ConfigSchema> = {
  azureTenantId: {
    envVars: ["AZURE_TENANT_ID"],
    required: true,
    isSecret: false,
  },
  azureClientId: {
    envVars: ["AZURE_CLIENT_ID", "BOT_APP_ID"],
    required: true,
    isSecret: false,
  },
  azureClientSecret: {
    envVars: ["AZURE_CLIENT_SECRET", "SECRET_AZURE_CLIENT_SECRET"],
    required: true,
    isSecret: true,
  },
  botAppId: {
    envVars: ["BOT_APP_ID", "AZURE_CLIENT_ID"],
    required: true,
    isSecret: false,
  },
  botAppPassword: {
    envVars: ["BOT_APP_PASSWORD", "SECRET_BOT_APP_PASSWORD"],
    required: false, // Optional for local development; required for Azure deployment
    isSecret: true,
    defaultValue: "",
  },
  botEndpoint: {
    envVars: ["BOT_ENDPOINT", "PUBLIC_BASE_URL"],
    required: false,
    isSecret: false,
  },
  graphBaseUrl: {
    envVars: ["GRAPH_BASE_URL"],
    required: false,
    isSecret: false,
    defaultValue: "https://graph.microsoft.com",
  },
  speechKey: {
    envVars: ["AZURE_SPEECH_KEY", "SECRET_AZURE_SPEECH_KEY"],
    required: true,
    isSecret: true,
  },
  speechRegion: {
    envVars: ["AZURE_SPEECH_REGION"],
    required: true,
    isSecret: false,
  },
  projectMissaUrl: {
    envVars: ["PROJECT_MISSA_URL"],
    required: true,
    isSecret: false,
  },
  sharedSecret: {
    envVars: ["SHARED_SECRET", "SECRET_MEETING_MEDIA_BOT_SHARED_SECRET"],
    required: true,
    isSecret: true,
  },
  port: {
    envVars: ["PORT"],
    required: false,
    isSecret: false,
    defaultValue: "4000",
  },
  autoJoinUserEmail: {
    envVars: ["AUTO_JOIN_USER_EMAIL"],
    required: false,
    isSecret: false,
  },
  autoJoinUserObjectId: {
    envVars: ["AUTO_JOIN_USER_OBJECT_ID"],
    required: false,
    isSecret: false,
  },
  calendarPollIntervalMs: {
    envVars: ["CALENDAR_POLL_INTERVAL_MS"],
    required: false,
    isSecret: false,
    defaultValue: "60000",
  },
};

function resolveEnvVar(envVars: string[], defaultValue?: string): string | undefined {
  for (const envVar of envVars) {
    const value = process.env[envVar];
    if (value && value.trim() !== "") {
      return value.trim();
    }
  }
  return defaultValue;
}

function maskSecret(value: string): string {
  if (value.length <= 4) return "****";
  return `${value.substring(0, 4)}${"*".repeat(Math.min(value.length - 4, 20))}`;
}

export function loadConfig(): MeetingMediaBotConfig {
  const errors: string[] = [];
  const config: Partial<MeetingMediaBotConfig> = {};

  for (const [key, schema] of Object.entries(configSchema)) {
    const configKey = key as keyof MeetingMediaBotConfig;
    const value = resolveEnvVar(schema.envVars, schema.defaultValue);

    if (!value || value === "") {
      if (schema.required) {
        errors.push(`Missing required: ${configKey} (set: ${schema.envVars.join(" or ")})`);
      } else if (configKey === "autoJoinUserEmail" || configKey === "autoJoinUserObjectId") {
        // Nullable optional fields default to null
        (config as Record<string, unknown>)[configKey] = null;
      } else if (configKey === "calendarPollIntervalMs") {
        (config as Record<string, unknown>)[configKey] = 60000;
      }
    } else {
      // Handle numeric fields
      if (configKey === "port" || configKey === "calendarPollIntervalMs") {
        (config as Record<string, unknown>)[configKey] = parseInt(value, 10);
      } else {
        (config as Record<string, unknown>)[configKey] = value;
      }

      // Log (mask secrets)
      const displayValue = schema.isSecret ? maskSecret(value) : value;
      console.log(`✓ ${configKey}: ${displayValue}`);
    }
  }

  // Auto-derive BOT_ENDPOINT from Azure's WEBSITE_HOSTNAME if not explicitly set
  if (!config.botEndpoint && process.env.WEBSITE_HOSTNAME) {
    config.botEndpoint = `https://${process.env.WEBSITE_HOSTNAME}`;
    console.log(`✓ botEndpoint: ${config.botEndpoint} (auto-derived from WEBSITE_HOSTNAME)`);
  } else if (!config.botEndpoint) {
    errors.push("Missing required: botEndpoint (set: BOT_ENDPOINT or deploy to Azure App Service)");
  }

  if (errors.length > 0) {
    console.error("\n❌ Configuration errors:\n" + errors.map(e => `  - ${e}`).join("\n"));
    process.exit(1);
  }

  console.log("✓ Configuration loaded successfully\n");
  return config as MeetingMediaBotConfig;
}

let cachedConfig: MeetingMediaBotConfig | null = null;

export function getConfig(): MeetingMediaBotConfig {
  if (!cachedConfig) {
    cachedConfig = loadConfig();
  }
  return cachedConfig;
}
