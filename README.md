## AI Email Smart Briefing — Project Documentation

### Overview
AI Email Smart Briefing is a Gmail Add-on that periodically summarizes unread emails using Gemini, ranks senders by importance, and sends a concise briefing to a configured recipient email. It also supports a self-serve “Request Full Body” workflow via email to fetch the full content of a specific message or a search query result.

- **Add-on name**: AI Email Smart Briefing
- **Entry point**: `onHomepage`
- **Runtime**: Apps Script V8
- **Scopes**: See Manifest section

### Architecture
- **UI layer (Gmail Add-on)**: Homepage card with controls to start/stop the service, set recipient email, and choose frequency.
- **Trigger management**: Creates and deletes time-based triggers for the briefing function.
- **Core logic**: 
  - `forwardAllEmails` composes and sends periodic AI briefings.
  - `processEmailRequest` processes email-based requests to fetch and forward full message bodies.
  - `getGeminiSummary` and `getSenderRankingFromGemini` call Gemini APIs for summarization and ranking.
- **Configuration**: Uses Apps Script `UserProperties` for per-user runtime settings, and `ScriptProperties` for the `GEMINI_API_KEY`.

### Setup
1. Open the project in Google Apps Script (this repo is configured for `clasp`).
2. Set script property `GEMINI_API_KEY` with a valid Google Generative Language API key:
   - Apps Script Editor → Project Settings → Script properties → Add `GEMINI_API_KEY`
3. Review required OAuth scopes in `appsscript.json` and authorize when prompted.
4. Deploy as a Gmail Add-on (test or production).

Optional (CLI):
- `clasp login`
- `clasp push` to upload local changes

### Manifest
`appsscript.json` notable fields:
- `runtimeVersion`: V8
- `oauthScopes`:
  - `https://www.googleapis.com/auth/userinfo.email`
  - `https://www.googleapis.com/auth/gmail.addons.execute`
  - `https://www.googleapis.com/auth/gmail.addons.current.action.compose`
  - `https://www.googleapis.com/auth/gmail.readonly`
  - `https://www.googleapis.com/auth/gmail.modify`
  - `https://www.googleapis.com/auth/script.send_mail`
  - `https://www.googleapis.com/auth/script.external_request`
  - `https://www.googleapis.com/auth/script.scriptapp`
- `addOns.gmail.homepageTrigger.runFunction`: `onHomepage`

### Configuration and Stored Properties
- **UserProperties** (per user):
  - `recipientEmail` (string): Where to send briefings and fetch replies.
  - `frequencyHours` (string): Hours between briefings.
  - `briefingTriggerId` (string): Unique ID of the time-based trigger for `forwardAllEmails`.
  - `requestTriggerId` (string, currently not used): ID of trigger for `processEmailRequest`.
  - `triggerId` (legacy): Old property retained for cleanup.
- **ScriptProperties** (shared project):
  - `GEMINI_API_KEY` (string): API key for Gemini endpoints.

### Public APIs (Functions)
Each function below is globally visible in Apps Script and can be executed from the editor or invoked by triggers/UI.

#### onHomepage(e)
- **Description**: Add-on homepage entry point. Reads state from `UserProperties` and returns a UI card via `buildHomepageCard`.
- **Parameters**:
  - `e` (Object): Event object provided by the add-on platform.
- **Returns**: `Card` (CardService)
- **Usage**:
  ```javascript
  const card = onHomepage({});
  ```

#### buildHomepageCard(isRunning, email, frequency)
- **Description**: Pure UI builder that returns the settings/status card.
- **Parameters**:
  - `isRunning` (boolean)
  - `email` (string)
  - `frequency` (string) — hours
- **Returns**: `Card`
- **Usage**:
  ```javascript
  const card = buildHomepageCard(true, 'you@example.com', '6');
  ```

#### handleRunNow(e)
- **Description**: Handles the “Run Briefing Manually Now” button. Validates presence of `recipientEmail` and runs `forwardAllEmails`. Returns an action response with a notification.
- **Parameters**:
  - `e` (Object): Event object (not used for inputs).
- **Returns**: `ActionResponse`
- **Usage**:
  ```javascript
  const response = handleRunNow({});
  ```

#### handleStartService(e)
- **Description**: Handles the “Start / Update Service” button. Validates inputs, stores `recipientEmail` and `frequencyHours`, deletes existing triggers, waits 10 seconds, and creates a new time-based trigger for `forwardAllEmails`.
- **Parameters**:
  - `e` (Object): Must include `formInput.recipient_email_input` (string) and `formInput.frequency_input` (string or number hours).
- **Returns**: `ActionResponse` that updates the current card and shows a notification.
- **Side effects**:
  - Updates `UserProperties`: `recipientEmail`, `frequencyHours`, `briefingTriggerId`.
  - Deletes any previously managed triggers.
- **Usage**:
  ```javascript
  const e = {
    formInput: {
      recipient_email_input: 'you@example.com',
      frequency_input: '6'
    }
  };
  const response = handleStartService(e);
  ```

#### handleStopService(e)
- **Description**: Handles the “Stop Service” button. Deletes managed triggers and clears user properties, updates the UI.
- **Parameters**:
  - `e` (Object): Event object (not used).
- **Returns**: `ActionResponse`
- **Side effects**:
  - Clears `UserProperties` and removes all managed triggers.
- **Usage**:
  ```javascript
  const response = handleStopService({});
  ```

#### deleteManagedTriggers()
- **Description**: Helper to delete all triggers previously created by this add-on (by stored IDs or known handler names). Cleans related `UserProperties`.
- **Parameters**: none
- **Returns**: `void`
- **Usage**:
  ```javascript
  deleteManagedTriggers();
  ```

#### processEmailRequest()
- **Description**: Processes unread inbox emails from the authorized `recipientEmail` with subject `"Fetch Email Body:" ...`. Supports two request formats:
  - `Fetch Email Body: id:<GmailThreadId>` (exact thread fetch)
  - `Fetch Email Body: <search query>` (first result of a Gmail search, excluding messages from yourself)
  Sends the full body of the identified message back to `recipientEmail` as HTML.
- **Parameters**: none
- **Returns**: `void`
- **Prerequisites**: `recipientEmail` set in `UserProperties`.
- **Usage (manual run)**:
  ```javascript
  processEmailRequest();
  ```
- **How to request via email**:
  - Compose an email from your authorized account to yourself with one of:
    - Subject: `Fetch Email Body: id:184f1a2b1c...`
    - Subject: `Fetch Email Body: from:billing@example.com newer_than:7d`
  - Leave it unread in Inbox. The function will read, process, reply with content, and mark the request thread read.

#### forwardAllEmails()
- **Description**: Main briefing function. Finds unread inbox emails not from `recipientEmail`, summarizes each with `getGeminiSummary`, groups by sender, ranks senders via `getSenderRankingFromGemini`, and emails a formatted briefing to `recipientEmail`. Adds per-message actions:
  - “Request Full Body” button: a mailto link pre-filling `Fetch Email Body: id:<threadId>`.
  - “Open in Gmail” link.
  Marks processed threads read after sending.
- **Parameters**: none
- **Returns**: `void`
- **Respects**: `MailApp.getRemainingDailyQuota()`; exits if no quota.
- **Usage**:
  ```javascript
  forwardAllEmails();
  ```

#### getGeminiSummary(text)
- **Description**: Calls Gemini (`gemini-1.5-flash`) to summarize email text into one or two Chinese sentences.
- **Parameters**:
  - `text` (string): Plain text body to summarize.
- **Returns**: `string` — Summary text, or an error placeholder if the call fails or the API key is missing.
- **Requires**: `GEMINI_API_KEY` in Script Properties.
- **Usage**:
  ```javascript
  const summary = getGeminiSummary('Long email body here...');
  ```

#### getSenderRankingFromGemini(senderGroups)
- **Description**: Requests a sender ranking from Gemini based on the summaries. Expects a `Map<string, Array<{summary: string}>>`. Returns a list of sender emails ordered from most to least important. Falls back to input order on failure.
- **Parameters**:
  - `senderGroups` (Map<string, Array<Object>>)
- **Returns**: `Array<string>` — Sender emails in ranked order.
- **Requires**: `GEMINI_API_KEY` in Script Properties.
- **Usage**:
  ```javascript
  const ranked = getSenderRankingFromGemini(new Map([
    ['boss@example.com', [{ summary: 'Approve budget' }]],
    ['newsletter@example.com', [{ summary: 'Weekly news' }]]
  ]));
  ```

### Gmail Add-on UI
- **Fields**:
  - Email input: `recipient_email_input`
  - Frequency dropdown: `frequency_input` (values: `1`, `3`, `6`, `12`, `24`)
- **Buttons**:
  - Start / Update Service → `handleStartService`
  - Stop Service → `handleStopService`
  - Run Briefing Manually Now → `handleRunNow`

### Triggers
- Created by `handleStartService`:
  - Time-based trigger for `forwardAllEmails` set to every `frequencyHours` hours.
- Managed by `deleteManagedTriggers` (and cleaned on Stop).
- Notes:
  - The code includes commented lines that previously created an hourly trigger for `processEmailRequest`. If you want continuous fetch-request processing, you can enable that block and store `requestTriggerId`.

### Error Handling and Logging
- Missing `GEMINI_API_KEY`:
  - `getGeminiSummary` returns an inline error message in the summary slot.
  - `getSenderRankingFromGemini` logs and returns a fallback ranking.
- External request failures log response code/body and return safe fallbacks.
- When the Mail quota is insufficient, `forwardAllEmails` logs a warning and exits.

### Quotas and Limits
- **MailApp quota**: The function checks `MailApp.getRemainingDailyQuota()` and will not send if depleted.
- **Gmail search**: Standard GmailApp limits apply.
- **Gemini API**: Subject to your API plan rate and quota.

### Security & Privacy
- Email content and summaries are sent to the configured `recipientEmail`.
- Summarization and ranking send email text/summaries to the Gemini API endpoint. Ensure your use is compliant with your privacy requirements and policies.

### Local Development Notes
- This repository uses `clasp` configuration (`.clasp.json`). If you are not the owner of the `scriptId`, you can:
  - Create a new Apps Script project and update `.clasp.json` with your `scriptId`, or
  - Remove `scriptId` and run `clasp create` to initialize a new project.

### End-to-End Usage Walkthrough
1. Open the add-on in Gmail (right side panel).
2. On the homepage card:
   - Enter your email in “Email address for briefings and fetched emails”.
   - Choose a frequency (e.g., Every 6 hours).
   - Click “Start / Update Service”.
3. To trigger a briefing immediately, click “Run Briefing Manually Now”.
4. In the briefing email, use “Request Full Body” to receive the full content of a specific email. Alternatively, send yourself an email with subject `Fetch Email Body: <query>` to fetch via Gmail search.
5. To stop, click “Stop Service” in the add-on.

### Troubleshooting
- Not receiving briefings:
  - Confirm `recipientEmail` is set via the UI.
  - Check triggers exist in Apps Script → Triggers.
  - Verify MailApp daily quota.
- Summaries show an error placeholder:
  - Ensure `GEMINI_API_KEY` is set in Script Properties and valid.
- Request email not processed:
  - Subject must start with `Fetch Email Body:` and be unread in Inbox.
  - Make sure the email is from the authorized `recipientEmail`.

---
Maintainers can extend functionality by enabling a dedicated trigger for `processEmailRequest` or adjusting the summarization/ranking prompts to suit different languages or styles.