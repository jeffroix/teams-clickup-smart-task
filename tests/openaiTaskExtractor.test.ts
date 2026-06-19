import test from "node:test";
import assert from "node:assert/strict";
import { parseTaskProposal } from "../services/openaiTaskExtractor.js";

test("parseTaskProposal accepts valid structured output", () => {
  const proposal = parseTaskProposal({
    title: "Prepare investor update",
    description: "Draft the investor update and circulate for comments.",
    assignee_hint: "Alice",
    due_date: "2026-07-01",
    priority: "high",
    source_summary: "The team asked for an investor update by next week.",
    open_questions: ["Which metrics need to be included?"],
  });

  assert.equal(proposal.title, "Prepare investor update");
  assert.equal(proposal.priority, "high");
  assert.deepEqual(proposal.open_questions, ["Which metrics need to be included?"]);
});

test("parseTaskProposal rejects malformed payloads", () => {
  assert.throws(() => {
    parseTaskProposal({
      title: "Missing description",
      assignee_hint: null,
      due_date: null,
      priority: "urgent",
      source_summary: "x",
      open_questions: [],
    });
  }, /description/);
});

test("parseTaskProposal rejects invalid priority", () => {
  assert.throws(() => {
    parseTaskProposal({
      title: "Bad priority",
      description: "Task",
      assignee_hint: null,
      due_date: null,
      priority: "p0",
      source_summary: "x",
      open_questions: [],
    });
  }, /Priority/);
});
