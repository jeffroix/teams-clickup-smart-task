import test from "node:test";
import assert from "node:assert/strict";
import { validateEnv } from "../app.js";

test("validateEnv reports the names of missing required variables", () => {
  let error: unknown;
  try {
    validateEnv({
      OPENAI_API_KEY: "sentinel-openai",
      CLICKUP_API_TOKEN: "sentinel-clickup",
      MICROSOFT_APP_ID: "sentinel-app-id",
    });
  } catch (caught) {
    error = caught;
  }

  assert.ok(error instanceof Error);
  assert.match(error.message, /CLICKUP_LIST_ID/);
  assert.match(error.message, /MICROSOFT_APP_PASSWORD/);
  assert.doesNotMatch(error.message, /sentinel-openai|sentinel-clickup|sentinel-app-id/);
});

test("validateEnv applies safe defaults when optional values are omitted", () => {
  const config = validateEnv({
    OPENAI_API_KEY: "sentinel-openai",
    CLICKUP_API_TOKEN: "sentinel-clickup",
    CLICKUP_LIST_ID: "sentinel-list-id",
    MICROSOFT_APP_ID: "sentinel-app-id",
    MICROSOFT_APP_PASSWORD: "sentinel-app-password",
  });

  assert.equal(config.openAiModel, "gpt-4.1-mini");
  assert.deepEqual(config.clickupAssigneeMap, {});
  assert.equal(config.microsoftTenantId, null);
  assert.equal(config.teamsThreadFetchEnabled, true);
  assert.equal(config.port, 3978);
});
