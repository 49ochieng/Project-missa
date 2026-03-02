/**
 * Microsoft Graph Client with retry logic and token management
 * Never logs tokens or secrets
 */

import { ClientSecretCredential } from "@azure/identity";
import { getConfig } from "../config";

interface GraphResponse<T = unknown> {
  success: boolean;
  data?: T;
  error?: string;
  statusCode?: number;
}

interface RetryConfig {
  maxRetries: number;
  baseDelayMs: number;
  maxDelayMs: number;
}

const DEFAULT_RETRY_CONFIG: RetryConfig = {
  maxRetries: 3,
  baseDelayMs: 1000,
  maxDelayMs: 30000,
};

export class GraphClient {
  private credential: ClientSecretCredential;
  private baseUrl: string;
  private accessToken: string | null = null;
  private tokenExpiry: Date | null = null;

  constructor() {
    const config = getConfig();
    this.credential = new ClientSecretCredential(
      config.azureTenantId,
      config.azureClientId,
      config.azureClientSecret
    );
    this.baseUrl = config.graphBaseUrl || "https://graph.microsoft.com";
  }

  /**
   * Get access token, refreshing if expired
   * Never logs the actual token
   */
  async getAccessToken(): Promise<string> {
    // Check if we have a valid cached token (with 5 min buffer)
    if (this.accessToken && this.tokenExpiry) {
      const bufferMs = 5 * 60 * 1000;
      if (new Date().getTime() < this.tokenExpiry.getTime() - bufferMs) {
        console.log("[GraphClient] Using cached access token");
        return this.accessToken;
      }
    }

    // Get new token
    console.log("[GraphClient] Acquiring new access token...");
    const tokenResponse = await this.credential.getToken(
      "https://graph.microsoft.com/.default"
    );

    this.accessToken = tokenResponse.token;
    this.tokenExpiry = tokenResponse.expiresOnTimestamp
      ? new Date(tokenResponse.expiresOnTimestamp)
      : new Date(Date.now() + 3600 * 1000); // Default 1 hour

    console.log("[GraphClient] Access token acquired (expires: " + this.tokenExpiry.toISOString() + ")");
    return this.accessToken;
  }

  /**
   * Make a Graph API request with automatic retry for 429 and 5xx errors
   */
  async request<T = unknown>(
    method: "GET" | "POST" | "PATCH" | "DELETE",
    path: string,
    body?: unknown,
    retryConfig: RetryConfig = DEFAULT_RETRY_CONFIG
  ): Promise<GraphResponse<T>> {
    const url = path.startsWith("http") ? path : `${this.baseUrl}${path}`;

    for (let attempt = 0; attempt <= retryConfig.maxRetries; attempt++) {
      try {
        const token = await this.getAccessToken();

        const headers: Record<string, string> = {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        };

        const options: RequestInit = {
          method,
          headers,
        };

        if (body && (method === "POST" || method === "PATCH")) {
          options.body = JSON.stringify(body);
        }

        // Log request (without sensitive data)
        console.log(`[GraphClient] ${method} ${path} (attempt ${attempt + 1})`);

        const response = await fetch(url, options);

        // Handle rate limiting (429)
        if (response.status === 429) {
          const retryAfter = response.headers.get("Retry-After");
          const delayMs = retryAfter
            ? parseInt(retryAfter, 10) * 1000
            : this.calculateBackoff(attempt, retryConfig);

          console.warn(`[GraphClient] Rate limited (429). Retry after ${delayMs}ms`);
          await this.delay(delayMs);
          continue;
        }

        // Handle server errors (5xx)
        if (response.status >= 500 && response.status < 600) {
          const delayMs = this.calculateBackoff(attempt, retryConfig);
          console.warn(`[GraphClient] Server error (${response.status}). Retry after ${delayMs}ms`);
          await this.delay(delayMs);
          continue;
        }

        // Parse response
        const contentType = response.headers.get("Content-Type") || "";
        let data: T | undefined;

        if (contentType.includes("application/json")) {
          data = (await response.json()) as T;
        } else if (response.status !== 204) {
          // For non-JSON responses (like transcripts)
          data = (await response.text()) as unknown as T;
        }

        if (!response.ok) {
          const errorMessage = this.extractErrorMessage(data, response.status);
          console.error(`[GraphClient] Error: ${errorMessage}`);
          return {
            success: false,
            error: errorMessage,
            statusCode: response.status,
          };
        }

        console.log(`[GraphClient] Success: ${method} ${path}`);
        return {
          success: true,
          data,
          statusCode: response.status,
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error";
        console.error(`[GraphClient] Request failed: ${errorMessage}`);

        if (attempt === retryConfig.maxRetries) {
          return {
            success: false,
            error: errorMessage,
          };
        }

        const delayMs = this.calculateBackoff(attempt, retryConfig);
        console.log(`[GraphClient] Retrying after ${delayMs}ms...`);
        await this.delay(delayMs);
      }
    }

    return {
      success: false,
      error: "Max retries exceeded",
    };
  }

  /**
   * Calculate exponential backoff delay
   */
  private calculateBackoff(attempt: number, config: RetryConfig): number {
    const delay = config.baseDelayMs * Math.pow(2, attempt);
    const jitter = Math.random() * 1000;
    return Math.min(delay + jitter, config.maxDelayMs);
  }

  /**
   * Delay helper
   */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * Extract error message from response
   */
  private extractErrorMessage(data: unknown, statusCode: number): string {
    if (data && typeof data === "object") {
      const errorObj = data as Record<string, unknown>;
      if (errorObj.error && typeof errorObj.error === "object") {
        const graphError = errorObj.error as Record<string, unknown>;
        return graphError.message as string || `HTTP ${statusCode}`;
      }
      if (errorObj.message) {
        return errorObj.message as string;
      }
    }
    return `HTTP ${statusCode}`;
  }
}

// Singleton instance
let graphClientInstance: GraphClient | null = null;

export function getGraphClient(): GraphClient {
  if (!graphClientInstance) {
    graphClientInstance = new GraphClient();
  }
  return graphClientInstance;
}
