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

## Extra tabs

- `Dashboard`: summary metrics and sender/subject counts.
- `Sender View`: a dropdown-driven tab that lets you choose a sender and see only emails from that sender.

## Apps Script files

The project is split into smaller Apps Script files for easier maintenance:

- [Config.gs](C:/CHED-OPSD/Catanes-Email/apps-script/Config.gs): shared configuration constants.
- [SpreadsheetMenu.gs](C:/CHED-OPSD/Catanes-Email/apps-script/SpreadsheetMenu.gs): menu items and top-level actions.
- [SheetViews.gs](C:/CHED-OPSD/Catanes-Email/apps-script/SheetViews.gs): sheet creation, formatting, dashboard, and sender view setup.
- [MailboxSync.gs](C:/CHED-OPSD/Catanes-Email/apps-script/MailboxSync.gs): Gmail sync pipeline and message cleanup.
- [WebApp.gs](C:/CHED-OPSD/Catanes-Email/apps-script/WebApp.gs): web app entry point.
- [WebData.gs](C:/CHED-OPSD/Catanes-Email/apps-script/WebData.gs): web app data loading and filtering.
- [Dashboard.html](C:/CHED-OPSD/Catanes-Email/apps-script/Dashboard.html): web UI.

## How the sheet works

- Only inbound emails are logged.
- The first sync baseline is fixed at February 1, 2026. After that, syncs continue incrementally for newer emails.
- `With Reply` is a checkbox that turns `TRUE` when the same thread contains a later reply sent from `jcatanes@ched.gov.ph`.
- `Message` is generated from the cleaned plain-text body of the email after removing common quoted-thread markers and trimming the result for the sheet.

## Setup

1. Sign in to Google as `jcatanes@ched.gov.ph`.
2. Open the production spreadsheet.
3. Go to `Extensions` > `Apps Script`.
4. Add the `.gs` files from the [apps-script](C:/CHED-OPSD/Catanes-Email/apps-script) folder into the Apps Script project.
5. Open `Project Settings`, enable `Show "appsscript.json" manifest file in editor`, then replace the manifest with [appsscript.json](C:/CHED-OPSD/Catanes-Email/apps-script/appsscript.json).
6. Save the project.
7. Run `bootstrapEmailMonitor` once from the Apps Script editor and approve the requested Gmail, Sheets, and trigger permissions.
8. Reload the spreadsheet and run `Sync now` when needed.

## Web page

The same Apps Script project also includes a web app UI for browsing the monitored emails outside the spreadsheet.

1. Make sure [WebApp.gs](C:/CHED-OPSD/Catanes-Email/apps-script/WebApp.gs), [WebData.gs](C:/CHED-OPSD/Catanes-Email/apps-script/WebData.gs), and [Dashboard.html](C:/CHED-OPSD/Catanes-Email/apps-script/Dashboard.html) are added to the project.
2. In Apps Script, go to `Deploy` > `New deployment`.
3. Choose `Web app`.
4. Set `Execute as` to the account that owns the mailbox and spreadsheet.
5. Set access according to your organization's policy.
6. Deploy and open the generated web app URL.

The web app includes:

- summary cards for total emails, senders, replied threads, pending replies, and today's volume
- a sender dropdown, reply-status filter, search box, and row limit selector
- a live table of email records loaded directly from `Email Log`

## Menu actions

- `Bootstrap monitor`: creates the sheets, installs the hourly trigger, and runs the first sync.
- `Setup sheets only`: rebuilds the `Email Log` and refreshes the `Dashboard`.
- `Sync now`: imports newly found inbound emails.
- `Resync from Feb 1, 2026`: reruns the import window from February 1, 2026 without creating duplicates.
- `Install hourly trigger`: recreates the time-driven sync trigger.
- `Reset sync state`: clears the sync checkpoint so the next run rescans from February 1, 2026.

## Notes

- The script must be authorized while signed in as `jcatanes@ched.gov.ph`; `GmailApp` reads the mailbox of the account that authorizes the script.
- Existing log sheets with the older column layout are preserved by renaming them to a timestamped backup sheet before the simplified `Email Log` is created.
- This version does not depend on AI Studio or any external AI service.
- Mail earlier than February 1, 2026 is intentionally excluded from sync.
