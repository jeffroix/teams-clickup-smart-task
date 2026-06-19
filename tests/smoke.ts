import fs from "node:fs";
import path from "node:path";
import { SmartClickUpMessageExtension } from "../teams/messageExtension.js";
import { Logger, TaskProposal, TeamsMessageContext } from "../types.js";

const fixturePath = path.join(process.cwd(), "tests", "fixtures", "fetchTask.invoke.json");

const logger: Logger = {
  info() {},
  warn() {},
  error() {},
};

const proposal: TaskProposal = {
  title: "Draft regulatory summary",
  description: "Write the regulatory summary and share for review.",
  assignee_hint: "Pat",
  due_date: "2026-06-26",
  priority: "high",
  source_summary: "The thread asked for a regulatory summary by Friday.",
  open_questions: ["Who is the final reviewer?"],
};

async function main() {
  const activity = JSON.parse(fs.readFileSync(fixturePath, "utf8")) as Record<string, unknown>;
  const bot = new SmartClickUpMessageExtension({
    extractor: {
      async extractTaskProposal(_context: TeamsMessageContext) {
        return proposal;
      },
    },
    clickup: {
      async createTaskFromProposal() {
        return { id: "task-1", name: proposal.title, url: "https://app.clickup.com/t/task-1" };
      },
    },
    logger,
    appId: "app-id",
    appPassword: "app-password",
    tenantId: null,
    threadFetchEnabled: false,
  });

  const preview = await bot.buildPreviewData(
    (activity.value ?? {}) as Record<string, unknown>,
    activity,
  );
  console.log(JSON.stringify({
    title: preview.proposal.title,
    selectedText: preview.context.selectedText,
    hasThreadMessages: preview.context.threadMessages.length > 0,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
