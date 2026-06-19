import test from "node:test";
import assert from "node:assert/strict";
import { validateEnv } from "../app.js";
import {
  buildClickUpPayload,
  parseDueDateToTimestamp,
} from "../services/clickupClient.js";
import { TaskProposal, TeamsMessageContext } from "../types.js";

const proposal: TaskProposal = {
  title: "Create launch brief",
  description: "Write the launch brief and collect feedback.",
  assignee_hint: "Alice",
  due_date: "2026-06-30",
  priority: "urgent",
  source_summary: "A launch brief was requested in the thread.",
  open_questions: ["Which launch date should we anchor on?"],
};

const context: TeamsMessageContext = {
  messageId: "msg-1",
  selectedText: "Can someone turn this into a ClickUp task today?",
  senderName: "Jeff",
  senderId: "user-1",
  conversationId: "conv-1",
  chatId: "chat-1",
  teamId: null,
  channelId: null,
  serviceUrl: "https://smba.trafficmanager.net/emea/",
  tenantId: "tenant-1",
  messageLink: "https://teams.microsoft.com/l/message/chat-1/msg-1",
  threadMessages: ["Please include launch milestones."],
  rawContext: {},
};

test("buildClickUpPayload maps fields and assignee hints", () => {
  const payload = buildClickUpPayload(proposal, context, { alice: 1234 });
  assert.equal(payload.name, "Create launch brief");
  assert.equal(payload.priority, 1);
  assert.deepEqual(payload.assignees, [1234]);
  assert.equal(typeof payload.markdown_content, "string");
  assert.match(String(payload.markdown_content), /Original selected Teams message/);
});

test("buildClickUpPayload leaves task unassigned when no mapping exists", () => {
  const payload = buildClickUpPayload(proposal, context, {});
  assert.equal(payload.assignees, undefined);
});

test("parseDueDateToTimestamp handles valid invalid and null values", () => {
  assert.equal(typeof parseDueDateToTimestamp("2026-06-30"), "number");
  assert.equal(parseDueDateToTimestamp("not-a-date"), undefined);
  assert.equal(parseDueDateToTimestamp(null), undefined);
});

test("validateEnv rejects missing required variables", () => {
  assert.throws(() => validateEnv({}), /Missing required env vars/);
});
