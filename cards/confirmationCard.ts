import { ClickUpTaskResult, TaskProposal } from "../types.js";

export function buildConfirmationCard(task: ClickUpTaskResult, proposal: TaskProposal): Record<string, unknown> {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "ClickUp task created",
        weight: "Bolder",
        size: "Medium",
      },
      {
        type: "TextBlock",
        text: proposal.title,
        wrap: true,
      },
      {
        type: "FactSet",
        facts: [
          { title: "Task ID", value: task.id },
          { title: "Priority", value: proposal.priority ?? "none" },
          { title: "Due", value: proposal.due_date ?? "not set" },
        ],
      },
      {
        type: "TextBlock",
        text: `[Open in ClickUp](${task.url})`,
        wrap: true,
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "Open task",
        url: task.url,
      },
    ],
  };
}

export function buildErrorCard(message: string): Record<string, unknown> {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Unable to create ClickUp task",
        weight: "Bolder",
        color: "Attention",
      },
      {
        type: "TextBlock",
        text: message,
        wrap: true,
      },
    ],
  };
}
