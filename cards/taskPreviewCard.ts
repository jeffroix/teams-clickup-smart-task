import { FetchTaskPreviewData } from "../types.js";

function questionsText(questions: string[]): string {
  return questions.join("\n");
}

export function buildTaskPreviewCard(data: FetchTaskPreviewData): Record<string, unknown> {
  const { proposal, context } = data;
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Create smart ClickUp task",
        weight: "Bolder",
        size: "Medium",
      },
      {
        type: "TextBlock",
        text: "Review and edit before the task is created.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "Input.Text",
        id: "title",
        label: "Title",
        value: proposal.title,
        isRequired: true,
        errorMessage: "Title is required.",
      },
      {
        type: "Input.Text",
        id: "description",
        label: "Description",
        value: proposal.description,
        isMultiline: true,
      },
      {
        type: "Input.Text",
        id: "assignee_hint",
        label: "Assignee hint",
        value: proposal.assignee_hint ?? "",
      },
      {
        type: "Input.Text",
        id: "due_date",
        label: "Due date (ISO or natural text kept as note)",
        value: proposal.due_date ?? "",
      },
      {
        type: "Input.ChoiceSet",
        id: "priority",
        label: "Priority",
        style: "compact",
        choices: [
          { title: "None", value: "" },
          { title: "Urgent", value: "urgent" },
          { title: "High", value: "high" },
          { title: "Normal", value: "normal" },
          { title: "Low", value: "low" },
        ],
        value: proposal.priority ?? "",
      },
      {
        type: "Input.Text",
        id: "source_summary",
        label: "Source summary",
        value: proposal.source_summary,
        isMultiline: true,
      },
      {
        type: "Input.Text",
        id: "open_questions",
        label: "Open questions (one per line)",
        value: questionsText(proposal.open_questions),
        isMultiline: true,
      },
      {
        type: "TextBlock",
        text: `Selected message: ${context.selectedText}`,
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Create task",
        data: {
          action: "createClickUpTask",
          context,
        },
      },
    ],
  };
}
