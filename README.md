# Gmail to Google Sheet Monitor

This project contains a Google Apps Script that logs mailbox activity from `jcatanes@ched.gov.ph` into the production Google Sheet:

- Sheet: [Monitoring Sheet](https://docs.google.com/spreadsheets/d/1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY/edit)
- Drive folder: [Project Folder](https://drive.google.com/drive/folders/13CNh31J9T1d9rhVQHw3nUPSHeDuvw2HG)

## What it creates

The script writes email metadata into an `Email Log` sheet with these columns:

1. Logged At
2. Message Date
3. Direction
4. From
5. To
6. Cc
7. Reply-To
8. Subject
9. Labels
10. Unread
11. Starred
12. Important
13. In Inbox
14. Priority Inbox
15. Has Attachments
16. Attachment Count
17. Attachment Names
18. Body Preview
19. Gmail Link
20. Thread ID
21. Message ID
22. RFC Message-ID
23. Thread Message Count
24. Thread Last Message Date

It also builds a `Dashboard` sheet with counts for unread mail, today's received and sent mail, attachment volume, and top senders/subjects.

## Setup

1. Sign in to Google as `jcatanes@ched.gov.ph`.
2. Open the production spreadsheet.
3. Go to `Extensions` > `Apps Script`.
4. Replace the default script with the contents of [EmailMonitor.gs](C:/CHED-OPSD/Catanes-Email/apps-script/EmailMonitor.gs).
5. Open `Project Settings`, enable `Show "appsscript.json" manifest file in editor`, then replace the manifest with [appsscript.json](C:/CHED-OPSD/Catanes-Email/apps-script/appsscript.json).
6. Save the project. If you want the Apps Script file itself organized in Drive, place the project in the provided Drive folder.
7. Run `bootstrapEmailMonitor` once from the Apps Script editor and approve the requested Gmail, Sheets, and trigger permissions.
8. Reload the spreadsheet and use the `Email Monitor` menu when needed.

## Operations

- `Bootstrap monitor`: creates the sheets, performs a first sync, and installs the hourly trigger.
- `Setup sheets only`: creates or refreshes the `Email Log` and `Dashboard` sheets.
- `Sync now`: imports new mailbox activity using an overlap window to avoid missed emails.
- `Backfill last 180 days`: imports older history without clearing existing rows.
- `Install hourly trigger`: recreates the time-driven sync trigger.
- `Reset sync state`: clears the sync checkpoint so the next run re-scans the recent window.

## Notes

- The script must be authorized while signed in as `jcatanes@ched.gov.ph`; `GmailApp` reads the mailbox of the account that authorizes the script.
- The first regular sync is limited to the last 30 days to keep Apps Script runtime manageable. Use `Backfill last 180 days` if you need more history.
- The script stores only email metadata and a short plain-text preview. It does not archive attachments into Drive.
