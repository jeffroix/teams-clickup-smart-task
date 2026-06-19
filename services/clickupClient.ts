import { AppConfig, ClickUpTaskResult, Logger, TaskProposal, TeamsMessageContext } from "../types.js";

const PRIORITY_MAP: Record<NonNullable<TaskProposal["priority"]>, number> = {
  urgent: 1,
  high: 2,
  normal: 3,
  low: 4,
};

function normalizeHint(value: string): string {
  return value.trim().toLowerCase();
}

export function parseDueDateToTimestamp(value: string | null): number | undefined {
  if (!value) {
    return undefined;
  }
  const timestamp = Date.parse(value);
  return Number.isFinite(timestamp) ? timestamp : undefined;
}

function renderThreadMessages(messages: string[]): string {
  if (messages.length === 0) {
    return "_No additional thread context was available._";
  }
  return messages.map((message, index) => `${index + 1}. ${message}`).join("\n");
}

export function buildClickUpMarkdownDescription(proposal: TaskProposal, context: TeamsMessageContext): string {
  const questions = proposal.open_questions.length > 0
    ? proposal.open_questions.map((question) => `- ${question}`).join("\n")
    : "- None";
  const sourceLink = context.messageLink ?? "Unavailable";
  return [
    "## Task summary",
    proposal.description,
    "",
    "## Source summary",
    proposal.source_summary,
    "",
    "## Assignee hint",
    proposal.assignee_hint ?? "None",
    "",
    "## Open questions",
    questions,
    "",
    "## Teams source",
    `- Message link: ${sourceLink}`,
    `- Conversation ID: ${context.conversationId ?? "Unavailable"}`,
    `- Message ID: ${context.messageId ?? "Unavailable"}`,
    `- Sender: ${context.senderName ?? "Unknown"}`,
    "",
    "## Original selected Teams message",
    context.selectedText,
    "",
    "## Thread context",
    renderThreadMessages(context.threadMessages),
  ].join("\n");
}

export function buildClickUpPayload(
  proposal: TaskProposal,
  context: TeamsMessageContext,
  assigneeMap: Record<string, number>,
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    name: proposal.title,
    markdown_content: buildClickUpMarkdownDescription(proposal, context),
    description: buildClickUpMarkdownDescription(proposal, context),
  };
  const dueDate = parseDueDateToTimestamp(proposal.due_date);
  if (dueDate) {
    payload.due_date = dueDate;
  }
  if (proposal.priority) {
    payload.priority = PRIORITY_MAP[proposal.priority];
  }
  if (proposal.assignee_hint) {
    const mappedUser = assigneeMap[normalizeHint(proposal.assignee_hint)];
    if (mappedUser) {
      payload.assignees = [mappedUser];
    }
  }
  return payload;
}

export class ClickUpClient {
  constructor(private readonly config: AppConfig, private readonly logger: Logger) {}

  async createTaskFromProposal(proposal: TaskProposal, context: TeamsMessageContext): Promise<ClickUpTaskResult> {
    const payload = buildClickUpPayload(proposal, context, this.config.clickupAssigneeMap);
    const response = await fetch(`https://api.clickup.com/api/v2/list/${this.config.clickupListId}/task`, {
      method: "POST",
      headers: {
        authorization: this.config.clickupApiToken,
        "content-type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      this.logger.error("clickup_create_failed", { status: response.status });
      throw new Error("ClickUp could not create the task.");
    }

    const created = await response.json() as { id?: string; name?: string; url?: string };
    const task = {
      id: created.id ?? "unknown",
      name: created.name ?? proposal.title,
      url: created.url ?? `https://app.clickup.com/t/${created.id ?? ""}`,
    };
    this.logger.info("clickup_create_succeeded", { task_id: task.id });
    return task;
  }
}
