# Teams ClickUp Smart Task

Bot-based Microsoft Teams message extension that turns a Teams message or thread into a reviewed ClickUp task proposal before creating the task.

## What it does

1. A user selects a Teams message and runs `Create smart ClickUp task`.
2. The app reads the selected message and available Teams metadata.
3. It attempts to fetch extra thread context from Microsoft Graph when configured.
4. It sends normalized context to OpenAI using strict structured outputs JSON.
5. It shows an editable Adaptive Card preview.
6. It creates the ClickUp task only after the user confirms.
7. It returns a confirmation card with the ClickUp link.

## Project layout

- `app.ts`: Express bootstrap, env validation, bot adapter, health route, logging.
- `teams/messageExtension.ts`: Teams action handling, thread fetch, confirmation flow.
- `services/openaiTaskExtractor.ts`: strict OpenAI extraction.
- `services/clickupClient.ts`: ClickUp payload builder and task creation.
- `cards/taskPreviewCard.ts`: editable preview card.
- `cards/confirmationCard.ts`: success and error cards.
- `tests/*.test.ts`: parser and payload builder tests.

## Setup

```bash
cd /home/jeffroix/teams-clickup-smart-task
npm install
cp .env.example .env
```

Fill in `.env`:

- `OPENAI_API_KEY`
- `CLICKUP_API_TOKEN`
- `CLICKUP_LIST_ID`
- `MICROSOFT_APP_ID`
- `MICROSOFT_APP_PASSWORD`

Optional but recommended:

- `MICROSOFT_TENANT_ID`
- `CLICKUP_ASSIGNEE_MAP`
- `TEAMS_THREAD_FETCH_ENABLED`

`CLICKUP_ASSIGNEE_MAP` is JSON:

```json
{"alice":123456,"bob":789012}
```

## Microsoft Teams and Azure setup

1. Create an Azure Bot or Bot Framework registration.
2. Set the bot messaging endpoint to `https://<public-domain>/api/messages`.
3. Save the Microsoft App ID and client secret into `.env`.
4. Create a public HTTPS tunnel for local work, for example with dev tunnels or ngrok.
5. Sideload a Teams app package built from [`manifest/manifest.template.json`](/home/jeffroix/teams-clickup-smart-task/manifest/manifest.template.json).

The manifest needs:

- real `MICROSOFT_APP_ID`
- your public domain in `validDomains`
- `color.png` and `outline.png` icons added before packaging

## Microsoft Graph thread fetch

Full thread fetch is best-effort. The app always uses the selected message text first, then tries Graph for more context.

Configure:

- `MICROSOFT_TENANT_ID=<your tenant id>`
- `TEAMS_THREAD_FETCH_ENABLED=true`

Graph permissions remain a manual tenant step. For channel messages, you typically need channel message read access. For chats, you need chat message read access. Grant the minimum permissions your tenant allows, then admin-consent them before testing. If Graph is unavailable or consent is missing, the app falls back to the selected Teams message.

## ClickUp setup

1. Create a personal ClickUp API token.
2. Find the target List ID.
3. Put both values in `.env`.
4. If you want assignment, add `CLICKUP_ASSIGNEE_MAP`.

The app maps:

- proposal `title` to ClickUp `name`
- proposal `description` to `markdown_content`
- proposal `due_date` to ClickUp `due_date` when parseable
- proposal `priority` to ClickUp numeric priority

It also preserves:

- the original Teams message text
- thread excerpts when available
- the Teams message link or fallback IDs

## OpenAI setup

The extractor uses the OpenAI Responses API with strict JSON schema output. You only need `OPENAI_API_KEY` unless you want to change the default model:

```bash
OPENAI_MODEL=gpt-4.1-mini
```

## Run

```bash
npm run typecheck
npm test
npm run start
```

Health check:

```bash
curl http://127.0.0.1:3978/health
```

## Local smoke check

A saved sample Teams invoke payload is included:

```bash
npm run smoke
```

That verifies the fetch-task preview path can normalize Teams input and build a task proposal preview payload.

## Manual validation

1. Start the app behind a public HTTPS tunnel.
2. Update the bot endpoint and manifest domain.
3. Sideload the Teams app package.
4. Invoke `Create smart ClickUp task` on a real Teams message.
5. Confirm the preview card is editable.
6. Submit and verify the ClickUp task contains the Teams source block and link.

## Notes

- The confirmation card is returned through the message extension response and the bot also attempts to post it into the conversation.
- If OpenAI, Graph, or ClickUp fails, the user gets a clear error card instead of silent failure.
