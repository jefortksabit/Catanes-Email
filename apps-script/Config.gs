const EMAIL_MONITOR_CONFIG = Object.freeze({
  monitoredMailbox: 'jcatanes@ched.gov.ph',
  spreadsheetId: '1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY',
  logSheetName: 'Email Log',
  dashboardSheetName: 'Dashboard',
  senderViewSheetName: 'Sender View',
  baseQuery: 'in:anywhere -in:trash -in:spam',
  initialSyncStartDate: '2026-02-01',
  overlapDays: 2,
  batchSize: 100,
  timeZone: 'Asia/Manila',
  maxFallbackChars: 280,
  scriptProperties: {
    lastSyncAt: 'EMAIL_MONITOR_LAST_SYNC_AT',
  },
  headers: [
    'Date Received',
    'From',
    'To',
    'Cc',
    'Subject',
    'Message',
    'Thread ID',
    'Message ID',
    'With Reply',
  ],
  columnWidths: [155, 260, 260, 220, 320, 430, 170, 190, 110],
});

const WEB_APP_CONFIG = Object.freeze({
  defaultRowLimit: 150,
  maxRowLimit: 500,
});
