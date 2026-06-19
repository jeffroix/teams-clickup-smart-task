import {
  AttachmentLayout,
  CardFactory,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  TeamsActivityHandler,
  TurnContext,
} from "botbuilder";
import { buildConfirmationCard, buildErrorCard } from "../cards/confirmationCard.js";
import { buildTaskPreviewCard } from "../cards/taskPreviewCard.js";
import {
  FetchTaskPreviewData,
  MessageExtensionDependencies,
  TaskProposal,
  TeamsMessageContext,
} from "../types.js";

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" ? value as Record<string, unknown> : {};
}

function asString(value: unknown): string | null {
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

function flattenHtml(value: string): string {
  return value.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

function parseQuestions(value: unknown): string[] {
  if (typeof value !== "string") {
    return [];
  }
  return value
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
}

function createMessageLink(teamId: string | null, channelId: string | null, messageId: string | null): string | null {
  if (!teamId || !channelId || !messageId) {
    return null;
  }
  return `https://teams.microsoft.com/l/message/${channelId}/${messageId}?groupId=${teamId}&tenantId=`;
}

async function getGraphToken(deps: MessageExtensionDependencies): Promise<string> {
  if (!deps.tenantId) {
    throw new Error("MICROSOFT_TENANT_ID is required for thread fetch.");
  }
  const body = new URLSearchParams({
    client_id: deps.appId,
    client_secret: deps.appPassword,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });
  const response = await fetch(`https://login.microsoftonline.com/${deps.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!response.ok) {
    throw new Error(`Graph token request failed with status ${response.status}.`);
  }
  const payload = await response.json() as { access_token?: string };
  if (!payload.access_token) {
    throw new Error("Graph token was missing.");
  }
  return payload.access_token;
}

async function graphJson<T>(token: string, path: string): Promise<T> {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: {
      authorization: `Bearer ${token}`,
      "content-type": "application/json",
    },
  });
  if (!response.ok) {
    throw new Error(`Graph request failed with status ${response.status}.`);
  }
  return await response.json() as T;
}

function extractMessagePayload(activityValue: Record<string, unknown>): Record<string, unknown> {
  const messagePayload = activityValue.messagePayload;
  if (messagePayload && typeof messagePayload === "object") {
    return messagePayload as Record<string, unknown>;
  }
  const commandContext = asRecord(activityValue.commandContext);
  return asRecord(commandContext.messagePayload);
}

async function fetchThreadMessages(
  deps: MessageExtensionDependencies,
  context: TeamsMessageContext,
): Promise<string[]> {
  if (!deps.threadFetchEnabled) {
    return [];
  }
  if (!deps.tenantId) {
    deps.logger.warn("graph_fetch_skipped", { reason: "missing_tenant_id" });
    return [];
  }
  try {
    const token = await getGraphToken(deps);
    if (context.teamId && context.channelId && context.messageId) {
      const root = await graphJson<{ body?: { content?: string } }>(
        token,
        `/teams/${encodeURIComponent(context.teamId)}/channels/${encodeURIComponent(context.channelId)}/messages/${encodeURIComponent(context.messageId)}`,
      );
      const replies = await graphJson<{ value?: Array<{ body?: { content?: string } }> }>(
        token,
        `/teams/${encodeURIComponent(context.teamId)}/channels/${encodeURIComponent(context.channelId)}/messages/${encodeURIComponent(context.messageId)}/replies`,
      );
      return [
        flattenHtml(root.body?.content ?? context.selectedText),
        ...(replies.value ?? []).map((item) => flattenHtml(item.body?.content ?? "")).filter(Boolean),
      ];
    }
    if (context.chatId && context.messageId) {
      const message = await graphJson<{ body?: { content?: string } }>(
        token,
        `/chats/${encodeURIComponent(context.chatId)}/messages/${encodeURIComponent(context.messageId)}`,
      );
      return [flattenHtml(message.body?.content ?? context.selectedText)];
    }
  } catch (error) {
    deps.logger.warn("graph_fetch_failed", {
      message_id: context.messageId,
      error: error instanceof Error ? error.message : "unknown",
    });
  }
  return [];
}

function normalizeContext(activityValue: Record<string, unknown>, activity: Record<string, unknown>): TeamsMessageContext {
  const messagePayload = extractMessagePayload(activityValue);
  const from = asRecord(messagePayload.from);
  const fromUser = asRecord(from.user);
  const body = asRecord(messagePayload.body);
  const channelData = asRecord(activity.channelData);
  const team = asRecord(channelData.team);
  const channel = asRecord(channelData.channel);
  return {
    messageId: asString(messagePayload.id) ?? asString(messagePayload.replyToId),
    selectedText: flattenHtml(asString(body.content) ?? asString(messagePayload.body?.toString?.()) ?? asString(activityValue.state) ?? ""),
    senderName: asString(fromUser.displayName) ?? asString(from.name),
    senderId: asString(fromUser.id) ?? asString(from.id),
    conversationId: asString(asRecord(activity.conversation).id),
    chatId: asString(channelData.chatId),
    teamId: asString(team.id),
    channelId: asString(channel.id),
    serviceUrl: asString(activity.serviceUrl),
    tenantId: asString(asRecord(channelData.tenant).id),
    messageLink: asString(messagePayload.linkToMessage) ?? createMessageLink(asString(team.id), asString(channel.id), asString(messagePayload.id)),
    threadMessages: [],
    rawContext: {
      activityValue,
      activity,
    },
  };
}

function normalizeSubmittedProposal(data: Record<string, unknown>): TaskProposal {
  return {
    title: String(data.title ?? "").trim(),
    description: String(data.description ?? "").trim(),
    assignee_hint: asString(data.assignee_hint),
    due_date: asString(data.due_date),
    priority: asString(data.priority) as TaskProposal["priority"],
    source_summary: String(data.source_summary ?? "").trim(),
    open_questions: parseQuestions(data.open_questions),
  };
}

function composeTaskModule(card: Record<string, unknown>): MessagingExtensionActionResponse {
  return {
    task: {
      type: "continue",
      value: {
        title: "Create smart ClickUp task",
        height: 540,
        width: 520,
        card: CardFactory.adaptiveCard(card),
      },
    },
  };
}

function composeResult(card: Record<string, unknown>): MessagingExtensionActionResponse {
  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list" as AttachmentLayout,
      attachments: [CardFactory.adaptiveCard(card)],
    },
  };
}

export class SmartClickUpMessageExtension extends TeamsActivityHandler {
  constructor(private readonly deps: MessageExtensionDependencies) {
    super();
  }

  async buildPreviewData(activityValue: Record<string, unknown>, activity: Record<string, unknown>): Promise<FetchTaskPreviewData> {
    const context = normalizeContext(activityValue, activity);
    context.threadMessages = await fetchThreadMessages(this.deps, context);
    if (!context.selectedText) {
      throw new Error("Teams did not provide selected message text.");
    }
    const proposal = await this.deps.extractor.extractTaskProposal(context);
    return { proposal, context };
  }

  override async handleTeamsMessagingExtensionFetchTask(
    context: TurnContext,
    action: MessagingExtensionAction,
  ): Promise<MessagingExtensionActionResponse> {
    this.deps.logger.info("teams_fetch_task_started");
    try {
      const preview = await this.buildPreviewData(asRecord(action.data), context.activity as unknown as Record<string, unknown>);
      this.deps.logger.info("teams_fetch_task_ready", {
        has_thread_context: preview.context.threadMessages.length > 0,
      });
      return composeTaskModule(buildTaskPreviewCard(preview));
    } catch (error) {
      this.deps.logger.error("teams_fetch_task_failed", {
        error: error instanceof Error ? error.message : "unknown",
      });
      return composeTaskModule(buildErrorCard(error instanceof Error ? error.message : "Unable to load preview."));
    }
  }

  override async handleTeamsMessagingExtensionSubmitAction(
    context: TurnContext,
    action: MessagingExtensionAction,
  ): Promise<MessagingExtensionActionResponse> {
    this.deps.logger.info("teams_submit_started");
    try {
      const submitted = asRecord(action.data);
      const proposal = normalizeSubmittedProposal(submitted);
      const rawContext = asRecord(submitted.context);
      const taskContext = {
        ...(rawContext as unknown as TeamsMessageContext),
        threadMessages: Array.isArray(rawContext.threadMessages)
          ? rawContext.threadMessages.filter((item): item is string => typeof item === "string")
          : [],
      };
      const task = await this.deps.clickup.createTaskFromProposal(proposal, taskContext);
      const confirmation = buildConfirmationCard(task, proposal);
      try {
        await context.sendActivity({
          attachments: [CardFactory.adaptiveCard(confirmation)],
        });
      } catch {
        // ponytail: submit responses already surface the confirmation card even if proactive send fails.
      }
      return composeResult(confirmation);
    } catch (error) {
      this.deps.logger.error("teams_submit_failed", {
        error: error instanceof Error ? error.message : "unknown",
      });
      return composeResult(buildErrorCard(error instanceof Error ? error.message : "Unable to create ClickUp task."));
    }
  }
}
