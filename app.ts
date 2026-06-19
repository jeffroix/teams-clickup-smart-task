import express from "express";
import { BotFrameworkAdapter } from "botbuilder";
import { buildErrorCard } from "./cards/confirmationCard.js";
import { ClickUpClient } from "./services/clickupClient.js";
import { OpenAiTaskExtractor } from "./services/openaiTaskExtractor.js";
import { SmartClickUpMessageExtension } from "./teams/messageExtension.js";
import { AppConfig, Logger } from "./types.js";

function parsePort(value: string | undefined): number {
  const port = Number(value ?? "3978");
  return Number.isFinite(port) && port > 0 ? port : 3978;
}

function parseBoolean(value: string | undefined, fallback: boolean): boolean {
  if (!value) {
    return fallback;
  }
  return ["1", "true", "yes", "on"].includes(value.trim().toLowerCase());
}

function parseAssigneeMap(value: string | undefined): Record<string, number> {
  if (!value || value.trim().length === 0) {
    return {};
  }
  const parsed = JSON.parse(value) as Record<string, unknown>;
  return Object.fromEntries(
    Object.entries(parsed)
      .filter(([, userId]) => Number.isFinite(Number(userId)))
      .map(([hint, userId]) => [hint.trim().toLowerCase(), Number(userId)]),
  );
}

export function validateEnv(env: NodeJS.ProcessEnv): AppConfig {
  const required = [
    "OPENAI_API_KEY",
    "CLICKUP_API_TOKEN",
    "CLICKUP_LIST_ID",
    "MICROSOFT_APP_ID",
    "MICROSOFT_APP_PASSWORD",
  ] as const;
  const missing = required.filter((name) => !env[name] || env[name]?.trim().length === 0);
  if (missing.length > 0) {
    throw new Error(`Missing required env vars: ${missing.join(", ")}`);
  }
  return {
    openAiApiKey: env.OPENAI_API_KEY!.trim(),
    openAiModel: env.OPENAI_MODEL?.trim() || "gpt-4.1-mini",
    clickupApiToken: env.CLICKUP_API_TOKEN!.trim(),
    clickupListId: env.CLICKUP_LIST_ID!.trim(),
    clickupAssigneeMap: parseAssigneeMap(env.CLICKUP_ASSIGNEE_MAP),
    microsoftAppId: env.MICROSOFT_APP_ID!.trim(),
    microsoftAppPassword: env.MICROSOFT_APP_PASSWORD!.trim(),
    microsoftTenantId: env.MICROSOFT_TENANT_ID?.trim() || null,
    teamsThreadFetchEnabled: parseBoolean(env.TEAMS_THREAD_FETCH_ENABLED, true),
    port: parsePort(env.PORT),
  };
}

function createLogger(): Logger {
  function write(level: "info" | "warn" | "error", event: string, data?: Record<string, unknown>) {
    console[level](JSON.stringify({
      level,
      event,
      ...data,
    }));
  }
  return {
    info(event, data) {
      write("info", event, data);
    },
    warn(event, data) {
      write("warn", event, data);
    },
    error(event, data) {
      write("error", event, data);
    },
  };
}

export function createApp(config: AppConfig) {
  const logger = createLogger();
  const extractor = new OpenAiTaskExtractor(config, logger);
  const clickup = new ClickUpClient(config, logger);
  const bot = new SmartClickUpMessageExtension({
    extractor,
    clickup,
    logger,
    appId: config.microsoftAppId,
    appPassword: config.microsoftAppPassword,
    tenantId: config.microsoftTenantId,
    threadFetchEnabled: config.teamsThreadFetchEnabled,
  });

  const adapter = new BotFrameworkAdapter({
    appId: config.microsoftAppId,
    appPassword: config.microsoftAppPassword,
  });

  adapter.onTurnError = async (context, error) => {
    logger.error("bot_turn_error", {
      message: error.message,
    });
    await context.sendActivity({
      attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildErrorCard("The app hit an unexpected error. Check server logs and try again."),
      }],
    });
  };

  const app = express();
  app.use(express.json({ limit: "1mb" }));

  app.get("/health", (_req, res) => {
    res.json({
      ok: true,
      service: "teams-clickup-smart-task",
      thread_fetch_enabled: config.teamsThreadFetchEnabled,
      has_tenant_id: Boolean(config.microsoftTenantId),
    });
  });

  app.post("/api/messages", async (req, res) => {
    logger.info("bot_activity_received", {
      type: req.body?.type,
      name: req.body?.name,
    });
    await adapter.processActivity(req, res, async (turnContext) => {
      await bot.run(turnContext);
    });
  });

  app.use((error: Error, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
    logger.error("http_request_failed", { message: error.message });
    res.status(500).json({ ok: false, error: "Internal server error." });
  });

  return { app, logger };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const config = validateEnv(process.env);
  const { app, logger } = createApp(config);
  app.listen(config.port, () => {
    logger.info("server_started", {
      port: config.port,
    });
  });
}
