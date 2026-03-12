const EMAIL_STATUS_OPTIONS = Object.freeze([
  'Pending',
  'In Progress',
  'Completed',
]);

const EMAIL_MONITOR_CONFIG = Object.freeze({
  monitoredMailbox: 'jcatanes@ched.gov.ph',
  spreadsheetId: '1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY',
  logSheetName: 'Email Log',
  dashboardSheetName: 'Dashboard',
  senderViewSheetName: 'Sender View',
  personnelSheetName: 'OPSD Personnel',
  baseQuery: 'in:anywhere -in:trash -in:spam',
  initialSyncStartDate: '2026-02-01',
  overlapDays: 2,
  batchSize: 100,
  timeZone: 'Asia/Manila',
  maxFallbackChars: 280,
  statusOptions: EMAIL_STATUS_OPTIONS,
  defaultStatus: EMAIL_STATUS_OPTIONS[0],
  inProgressStatus: EMAIL_STATUS_OPTIONS[1],
  completedStatus: EMAIL_STATUS_OPTIONS[2],
  scriptProperties: {
    lastSyncAt: 'EMAIL_MONITOR_LAST_SYNC_AT',
  },
  legacyHeaders: [
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
  previousHeaders: [
    'Reference Number',
    'Date Received',
    'From',
    'To',
    'Cc',
    'Subject',
    'Message',
    'Thread ID',
    'Message ID',
    'With Reply',
    'Status Update',
  ],
  currentHeadersWithoutPersonnel: [
    'Reference Number',
    'Date Received',
    'From',
    'To',
    'Cc',
    'Subject',
    'Message',
    'Thread ID',
    'Message ID',
    'Status',
    'Status Update',
  ],
  headers: [
    'Reference Number',
    'Date Received',
    'From',
    'To',
    'Cc',
    'Subject',
    'Message',
    'Thread ID',
    'Message ID',
    'OPSD Personnel',
    'Status',
    'Status Update',
  ],
  columnWidths: [170, 155, 260, 260, 220, 320, 430, 170, 190, 240, 130, 220],
  personnelAssignmentSeparator: ', ',
  personnelHeaders: [
    'UserEmail',
    'UserName',
    'Division',
    'Role',
    'Position',
    'SortOrder',
    'IsActive',
  ],
  personnelColumnWidths: [260, 220, 180, 170, 220, 100, 90],
});

const EMAIL_LOG_COLUMN_INDEX = Object.freeze({
  referenceNumber: 1,
  dateReceived: 2,
  from: 3,
  to: 4,
  cc: 5,
  subject: 6,
  message: 7,
  threadId: 8,
  messageId: 9,
  personnel: 10,
  status: 11,
  statusUpdate: 12,
});

const OPSD_PERSONNEL_COLUMN_INDEX = Object.freeze({
  userEmail: 1,
  userName: 2,
  division: 3,
  role: 4,
  position: 5,
  sortOrder: 6,
  isActive: 7,
});

const WEB_APP_CONFIG = Object.freeze({
  defaultRowLimit: 150,
  maxRowLimit: 500,
});

function normalizeEmailStatusValue_(value) {
  const rawValue =
    value === true || value === false
      ? mapReplyFlagToStatus_(value)
      : String(value || '').trim();

  return EMAIL_MONITOR_CONFIG.statusOptions.indexOf(rawValue) !== -1
    ? rawValue
    : EMAIL_MONITOR_CONFIG.defaultStatus;
}

function mapReplyFlagToStatus_(value) {
  return value === true
    ? EMAIL_MONITOR_CONFIG.completedStatus
    : EMAIL_MONITOR_CONFIG.defaultStatus;
}

function isOpenEmailStatus_(value) {
  return (
    normalizeEmailStatusValue_(value) !== EMAIL_MONITOR_CONFIG.completedStatus
  );
}
