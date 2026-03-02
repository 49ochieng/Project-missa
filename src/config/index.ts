/**
 * Configuration module for environment variable validation
 * Validates required configuration on startup and provides type-safe access
 * SECURITY: Never logs secret values
 */

export interface AppConfig {
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
  publicBaseUrl: string;

  // Azure OpenAI
  azureOpenAIEndpoint: string;
  azureOpenAIApiKey: string;
  azureOpenAIDeployment: string;

  // Azure Speech Service
  azureSpeechKey: string;
  azureSpeechRegion: string;

  // Database
  sqlServer: string;
  sqlDatabase: string;
  sqlUsername: string;
  sqlPassword: string;
  databaseConnectionString: string;
  sqlitePath: string;

  // Meeting Media Bot
  meetingMediaBotUrl: string;
  meetingMediaBotSharedSecret: string;
}

/**
 * Configuration schema defining required environment variables
 * Secret variables are marked with isSecret flag to prevent logging
 */
const configSchema: Record<
  keyof AppConfig,
  { envVars: string[]; required: boolean; isSecret: boolean }
> = {
  // Azure Authentication
  azureTenantId: {
    envVars: ["AZURE_TENANT_ID", "TEAMS_APP_TENANT_ID", "BOT_TENANT_ID"],
    required: true,
    isSecret: false,
  },
  azureClientId: {
    envVars: ["AZURE_CLIENT_ID"],
    required: true,
    isSecret: false,
  },
  azureClientSecret: {
    envVars: ["SECRET_AZURE_CLIENT_SECRET", "AZURE_CLIENT_SECRET"],
    required: true,
    isSecret: true,
  },

  // Bot Configuration
  botAppId: {
    envVars: ["BOT_APP_ID", "BOT_ID"],
    required: false, // Populated during debug session for local development
    isSecret: false,
  },
  botAppPassword: {
    envVars: ["SECRET_BOT_APP_PASSWORD", "BOT_APP_PASSWORD", "SECRET_BOT_PASSWORD"],
    required: false, // Optional for local development with tunneling
    isSecret: true,
  },
  botEndpoint: {
    envVars: ["BOT_ENDPOINT", "PUBLIC_BASE_URL"],
    required: false, // Used for callback URLs
    isSecret: false,
  },

  // Microsoft Graph
  graphBaseUrl: {
    envVars: ["GRAPH_BASE_URL"],
    required: true,
    isSecret: false,
  },
  publicBaseUrl: {
    envVars: ["PUBLIC_BASE_URL", "BOT_ENDPOINT"],
    required: false, // Optional for local development
    isSecret: false,
  },

  // Azure OpenAI
  azureOpenAIEndpoint: {
    envVars: ["AZURE_OPENAI_ENDPOINT", "AOAI_ENDPOINT"],
    required: true,
    isSecret: false,
  },
  azureOpenAIApiKey: {
    envVars: ["SECRET_AZURE_OPENAI_API_KEY", "AZURE_OPENAI_API_KEY", "AOAI_API_KEY"],
    required: true,
    isSecret: true,
  },
  azureOpenAIDeployment: {
    envVars: ["AZURE_OPENAI_DEPLOYMENT_NAME", "AZURE_OPENAI_DEPLOYMENT", "AOAI_MODEL"],
    required: true,
    isSecret: false,
  },

  // Azure Speech Service
  azureSpeechKey: {
    envVars: ["SECRET_AZURE_SPEECH_KEY", "AZURE_SPEECH_KEY"],
    required: true,
    isSecret: true,
  },
  azureSpeechRegion: {
    envVars: ["AZURE_SPEECH_REGION"],
    required: true,
    isSecret: false,
  },

  // Database
  sqlServer: {
    envVars: ["SQL_SERVER", "SQL_SERVER_FQDN"],
    required: false, // Not required if using SQLite or connection string
    isSecret: false,
  },
  sqlDatabase: {
    envVars: ["SQL_DATABASE", "SQL_DATABASE_NAME"],
    required: false, // Not required if using SQLite or connection string
    isSecret: false,
  },
  sqlUsername: {
    envVars: ["SQL_USERNAME"],
    required: false, // Not required if using SQLite or connection string
    isSecret: false,
  },
  sqlPassword: {
    envVars: ["SQL_PASSWORD", "SQL_ADMIN_PASSWORD"],
    required: false, // Not required if using SQLite or connection string
    isSecret: true,
  },
  databaseConnectionString: {
    envVars: ["DATABASE_CONNECTION_STRING", "SQL_CONNECTION_STRING"],
    required: false, // Optional - can use individual SQL vars or SQLite
    isSecret: true,
  },
  sqlitePath: {
    envVars: ["SQLITE_PATH", "CONVERSATIONS_DB_PATH"],
    required: false, // Optional - defaults to ./src/storage/conversations.db
    isSecret: false,
  },

  // Meeting Media Bot (internal service)
  meetingMediaBotUrl: {
    envVars: ["MEETING_MEDIA_BOT_URL"],
    required: false, // Only required when using meeting capture
    isSecret: false,
  },
  meetingMediaBotSharedSecret: {
    envVars: ["SECRET_MEETING_MEDIA_BOT_SHARED_SECRET", "MEETING_MEDIA_BOT_SHARED_SECRET"],
    required: false, // Only required when using meeting capture
    isSecret: true,
  },
};

/**
 * Resolves a configuration value from multiple possible environment variable names
 * Returns the first non-empty value found, or undefined if none exist
 */
function resolveEnvVar(envVars: string[]): string | undefined {
  for (const envVar of envVars) {
    const value = process.env[envVar];
    if (value && value.trim() !== "") {
      return value.trim();
    }
  }
  return undefined;
}

/**
 * Masks a secret value for safe logging
 * Shows first 4 characters only to help identify which secret is missing
 */
function maskSecret(value: string): string {
  if (value.length <= 4) {
    return "****";
  }
  return `${value.substring(0, 4)}${"*".repeat(value.length - 4)}`;
}

/**
 * Validates and loads application configuration from environment variables
 * Fails fast with descriptive error messages if required variables are missing
 * 
 * @throws {Error} If any required configuration is missing or invalid
 * @returns {AppConfig} Validated configuration object
 */
export function loadConfig(): AppConfig {
  const errors: string[] = [];
  const warnings: string[] = [];
  const config: Partial<AppConfig> = {};

  // Validate each configuration field
  for (const [key, schema] of Object.entries(configSchema)) {
    const configKey = key as keyof AppConfig;
    const value = resolveEnvVar(schema.envVars);

    if (!value || value === "") {
      if (schema.required) {
        const envVarsList = schema.envVars.join(" or ");
        errors.push(
          `Missing required configuration: ${configKey}\n` +
            `  Set one of: ${envVarsList}`
        );
      } else {
        const envVarsList = schema.envVars.join(" or ");
        warnings.push(
          `Optional configuration not set: ${configKey} (${envVarsList})`
        );
      }
    } else {
      config[configKey] = value;
      
      // Log successful load (mask secrets)
      const displayValue = schema.isSecret ? maskSecret(value) : value;
      const usedVar = schema.envVars.find(v => process.env[v] === value) || schema.envVars[0];
      console.log(`✓ Loaded ${configKey} from ${usedVar}: ${displayValue}`);
    }
  }

  // Log warnings for optional missing configs
  if (warnings.length > 0) {
    console.warn("\n⚠️  Optional configuration warnings:");
    warnings.forEach((warning) => console.warn(`  ${warning}`));
    console.warn("");
  }

  // Fail fast if any required configuration is missing
  if (errors.length > 0) {
    const errorMessage =
      "\n❌ Configuration validation failed!\n\n" +
      errors.map((e) => `  ${e}`).join("\n\n") +
      "\n\nPlease set the required environment variables in your .env.*.user file and restart.\n";
    
    throw new Error(errorMessage);
  }

  console.log("✓ All required configuration loaded successfully\n");
  
  return config as AppConfig;
}

/**
 * Cached configuration instance
 * Loaded once on first access for performance
 */
let cachedConfig: AppConfig | null = null;

/**
 * Gets the application configuration (cached after first load)
 * @returns {AppConfig} Application configuration
 */
export function getConfig(): AppConfig {
  if (!cachedConfig) {
    cachedConfig = loadConfig();
  }
  return cachedConfig;
}

/**
 * Clears the cached configuration (useful for testing)
 */
export function clearConfigCache(): void {
  cachedConfig = null;
}
