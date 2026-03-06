/**
 * Meeting Media Bot - Entry Point
 * Express server for handling Teams meeting media and transcription
 */

import express, { Request, Response, NextFunction } from "express";
import { loadConfig, getConfig } from "./config";
import { startCalendarPoller, stopCalendarPoller } from "./calendar/calendarPoller";
import callbackRoutes from "./routes/callbacks";
import apiRoutes from "./routes/api";
import acsCallbackRoutes from "./routes/acsCallbacks";
import { getGraphClient } from "./graph/graphClient";
import { initializeAcs, isAcsConfigured, shutdownAcs } from "./acs/acsTranscriber";

// Load configuration first
loadConfig();
const config = getConfig();

const app = express();

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Request logging
app.use((req: Request, _res: Response, next: NextFunction) => {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] ${req.method} ${req.path}`);
  next();
});

// Mount routes
app.use("/api/calls", callbackRoutes);
app.use("/api/acs", acsCallbackRoutes);
app.use("/api", apiRoutes);

// Root health check
app.get("/", (_req: Request, res: Response) => {
  res.json({
    name: "meeting-media-bot",
    version: "1.0.0",
    status: "running",
    timestamp: new Date().toISOString(),
  });
});

// 404 handler
app.use((_req: Request, res: Response) => {
  res.status(404).json({ error: "Not found" });
});

// Global error handler
app.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  console.error("[Server] Unhandled error:", err);
  res.status(500).json({
    error: "Internal server error",
    message: process.env.NODE_ENV === "development" ? err.message : undefined,
  });
});

// Start calendar auto-join poller if configured
if (config.autoJoinUserEmail || config.autoJoinUserObjectId) {
  startCalendarPoller();
} else {
  console.log("[Calendar] Auto-join not configured (set AUTO_JOIN_USER_EMAIL or AUTO_JOIN_USER_OBJECT_ID to enable)");
}

// Start server — use process.env.PORT directly (iisnode sets it to a named pipe on Azure)
const PORT = process.env.PORT || config.port;

app.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║            Meeting Media Bot - Started                     ║
╠════════════════════════════════════════════════════════════╣
║  Port: ${PORT.toString().padEnd(51)}║
║  Callback URL: ${config.botEndpoint}/api/calls/callback
║  Health: http://localhost:${PORT}/api/health
╚════════════════════════════════════════════════════════════╝

Available endpoints:
  POST /api/meetings/join         - Join a meeting and start capture
  POST /api/meetings/leave        - Leave a meeting
  GET  /api/meetings/:id/status   - Get meeting capture status
  GET  /api/meetings/active       - List active captures
  POST /api/calls/callback        - Graph notification webhook
  GET  /api/health                - Health check
`);

  // Initialize ACS transcription if configured (optional)
  if (isAcsConfigured()) {
    initializeAcs(8081);
    console.log("[Startup] ACS transcription fallback enabled");
  } else {
    console.log("[Startup] ACS transcription not configured (optional — set ACS_CONNECTION_STRING to enable)");
  }

  // Non-blocking: verify Graph API credentials on startup
  getGraphClient().testConnection().then(result => {
    if (result.success) {
      console.log("[Startup] Graph API connection verified");
    } else {
      console.error(`[Startup] Graph API connection FAILED: ${result.error}`);
      console.error("[Startup] Meeting join will fail until credentials are fixed");
    }
  });
});

// Graceful shutdown
process.on("SIGTERM", () => {
  console.log("[Server] SIGTERM received, shutting down gracefully...");
  stopCalendarPoller();
  shutdownAcs();
  process.exit(0);
});

process.on("SIGINT", () => {
  console.log("[Server] SIGINT received, shutting down gracefully...");
  stopCalendarPoller();
  shutdownAcs();
  process.exit(0);
});

export default app;
