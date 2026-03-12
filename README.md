# Gmail to Google Sheet Monitor

This project contains a Google Apps Script that logs inbound mailbox activity from `jcatanes@ched.gov.ph` into the production Google Sheet:

- Sheet: [Monitoring Sheet](https://docs.google.com/spreadsheets/d/1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY/edit)
- Drive folder: [Project Folder](https://drive.google.com/drive/folders/13CNh31J9T1d9rhVQHw3nUPSHeDuvw2HG)

## Email Log columns

The `Email Log` sheet keeps only the essential monitoring columns:

1. Date Received
2. From
3. To
4. Cc
5. Subject
6. Message
7. Thread ID
8. Message ID
9. With Reply

## How the sheet works

- Only inbound emails are logged.
- `With Reply` is a checkbox that turns `TRUE` when the same thread contains a later reply sent from `jcatanes@ched.gov.ph`.
- `Message` is generated through Google Gemini when a Gemini API key is configured.
- If no Gemini API key is configured yet, `Message` falls back to a shortened plain-text extract so syncing can continue.

## Setup

1. Sign in to Google as `jcatanes@ched.gov.ph`.
2. Open the production spreadsheet.
3. Go to `Extensions` > `Apps Script`.
4. Replace the default script with the contents of [EmailMonitor.gs](C:/CHED-OPSD/Catanes-Email/apps-script/EmailMonitor.gs).
5. Open `Project Settings`, enable `Show "appsscript.json" manifest file in editor`, then replace the manifest with [appsscript.json](C:/CHED-OPSD/Catanes-Email/apps-script/appsscript.json).
6. Save the project.
7. Create a Gemini API key in [Google AI Studio](https://aistudio.google.com/app/apikey).
8. Run `bootstrapEmailMonitor` once from the Apps Script editor and approve the requested Gmail, Sheets, trigger, and external request permissions.
9. Reload the spreadsheet, open the `Email Monitor` menu, and use `Set Gemini API key`.
10. Run `Sync now`.

## Menu actions

- `Bootstrap monitor`: creates the sheets, installs the hourly trigger, and runs the first sync.
- `Setup sheets only`: rebuilds the `Email Log` and refreshes the `Dashboard`.
- `Sync now`: imports newly found inbound emails.
- `Backfill last 180 days`: imports older inbound email history.
- `Set Gemini API key`: stores the Gemini API key in Apps Script script properties.
- `Clear Gemini API key`: removes the stored Gemini API key.
- `Install hourly trigger`: recreates the time-driven sync trigger.
- `Reset sync state`: clears the sync checkpoint so the next run rescans the recent window.

## Notes

- The script must be authorized while signed in as `jcatanes@ched.gov.ph`; `GmailApp` reads the mailbox of the account that authorizes the script.
- Existing log sheets with the older column layout are preserved by renaming them to a timestamped backup sheet before the simplified `Email Log` is created.
- The email body text used for the `Message` summary is sent to the Gemini API. Review that against your organization's data handling rules before enabling the API key.
