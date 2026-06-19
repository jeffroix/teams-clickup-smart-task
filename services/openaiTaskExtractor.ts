import { AppConfig, Logger, TaskProposal, TeamsMessageContext } from "../types.js";

const TASK_PROPOSAL_SCHEMA = {
  name: "task_proposal",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    required: [
      "title",
      "description",
      "assignee_hint",
      "due_date",
      "priority",
      "source_summary",
      "open_questions",
    ],
    properties: {
      title: { type: "string" },
      description: { type: "string" },
      assignee_hint: { type: ["string", "null"] },
      due_date: { type: ["string", "null"] },
      priority: {
        type: ["string", "null"],
        enum: ["urgent", "high", "normal", "low", null],
      },
      source_summary: { type: "string" },
      open_questions: {
        type: "array",
        items: { type: "string" },
      },
    },
  },
} as const;

function asNullableString(value: unknown): string | null {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : null;
}

function asString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`Missing or invalid string field: ${field}`);
  }
  return value.trim();
}

function asStringArray(value: unknown, field: string): string[] {
  if (!Array.isArray(value) || value.some((item) => typeof item !== "string")) {
    throw new Error(`Missing or invalid string[] field: ${field}`);
  }
  return value.map((item) => item.trim()).filter(Boolean);
}

export function parseTaskProposal(raw: unknown): TaskProposal {
  if (!raw || typeof raw !== "object") {
    throw new Error("OpenAI response was not an object.");
  }
  const record = raw as Record<string, unknown>;
  const priority = record.priority;
  if (priority !== null && priority !== undefined && !["urgent", "high", "normal", "low"].includes(String(priority))) {
    throw new Error("Priority value was invalid.");
  }
  return {
    title: asString(record.title, "title"),
    description: asString(record.description, "description"),
    assignee_hint: asNullableString(record.assignee_hint),
    due_date: asNullableString(record.due_date),
    priority: priority == null ? null : String(priority) as TaskProposal["priority"],
    source_summary: asString(record.source_summary, "source_summary"),
    open_questions: asStringArray(record.open_questions, "open_questions"),
  };
}

function extractOutputText(payload: Record<string, unknown>): string | null {
  if (typeof payload.output_text === "string" && payload.output_text.trim().length > 0) {
    return payload.output_text;
  }
  const output = payload.output;
  if (!Array.isArray(output)) {
    return null;
  }
  for (const item of output) {
    if (!item || typeof item !== "object") {
      continue;
    }
    const content = (item as { content?: unknown }).content;
    if (!Array.isArray(content)) {
      continue;
    }
    for (const part of content) {
      if (!part || typeof part !== "object") {
        continue;
      }
      const text = (part as { text?: unknown }).text;
      if (typeof text === "string" && text.trim().length > 0) {
        return text;
      }
    }
  }
  return null;
}

function buildPrompt(context: TeamsMessageContext): string {
  return [
    "You extract a ClickUp task proposal from a Microsoft Teams message and thread context.",
    "Return only the requested JSON fields.",
    "Do not omit fields. Use null when unknown.",
    "Keep the original Teams content out of the generated description except where needed for context because the app separately preserves the raw message.",
    "",
    `Selected message:\n${context.selectedText}`,
    "",
    `Sender: ${context.senderName ?? "unknown"}`,
    `Message link: ${context.messageLink ?? "unavailable"}`,
    "",
    `Thread context:\n${context.threadMessages.join("\n---\n") || "No extra thread context available."}`,
  ].join("\n");
}

export class OpenAiTaskExtractor {
  constructor(private readonly config: AppConfig, private readonly logger: Logger) {}

  async extractTaskProposal(context: TeamsMessageContext): Promise<TaskProposal> {
    const response = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        authorization: `Bearer ${this.config.openAiApiKey}`,
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: this.config.openAiModel,
        input: buildPrompt(context),
        text: {
          format: {
            type: "json_schema",
            ...TASK_PROPOSAL_SCHEMA,
          },
        },
      }),
    });

    if (!response.ok) {
      this.logger.error("openai_request_failed", { status: response.status });
      throw new Error("OpenAI could not build a task proposal right now.");
    }

    const payload = await response.json() as Record<string, unknown>;
    const outputText = extractOutputText(payload);
    if (!outputText) {
      this.logger.error("openai_missing_output");
      throw new Error("OpenAI returned an empty task proposal.");
    }

    let parsedJson: unknown;
    try {
      parsedJson = JSON.parse(outputText);
    } catch {
      this.logger.error("openai_invalid_json");
      throw new Error("OpenAI returned invalid task JSON.");
    }

    const proposal = parseTaskProposal(parsedJson);
    this.logger.info("openai_extraction_succeeded", {
      title_length: proposal.title.length,
      open_question_count: proposal.open_questions.length,
    });
    return proposal;
  }
}
