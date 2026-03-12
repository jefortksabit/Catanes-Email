const EMAIL_MONITOR_CONFIG = Object.freeze({
  monitoredMailbox: 'jcatanes@ched.gov.ph',
  spreadsheetId: '1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY',
  logSheetName: 'Email Log',
  dashboardSheetName: 'Dashboard',
  baseQuery: 'in:anywhere -in:trash -in:spam',
  initialLookbackDays: 30,
  overlapDays: 2,
  batchSize: 100,
  maxBodyPreviewChars: 250,
  scriptProperties: {
    lastSyncAt: 'EMAIL_MONITOR_LAST_SYNC_AT',
  },
  headers: [
    'Logged At',
    'Message Date',
    'Direction',
    'From',
    'To',
    'Cc',
    'Reply-To',
    'Subject',
    'Labels',
    'Unread',
    'Starred',
    'Important',
    'In Inbox',
    'Priority Inbox',
    'Has Attachments',
    'Attachment Count',
    'Attachment Names',
    'Body Preview',
    'Gmail Link',
    'Thread ID',
    'Message ID',
    'RFC Message-ID',
    'Thread Message Count',
    'Thread Last Message Date',
  ],
  columnWidths: [
    150,
    150,
    110,
    260,
    260,
    220,
    220,
    320,
    180,
    80,
    80,
    90,
    90,
    110,
    120,
    130,
    260,
    420,
    240,
    180,
    180,
    200,
    150,
    170,
  ],
  timeZone: 'Asia/Manila',
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email Monitor')
    .addItem('Bootstrap monitor', 'bootstrapEmailMonitor')
    .addItem('Setup sheets only', 'setupEmailMonitor')
    .addItem('Sync now', 'syncMailbox')
    .addItem('Backfill last 180 days', 'backfillLast180Days')
    .addItem('Install hourly trigger', 'installHourlySyncTrigger')
    .addSeparator()
    .addItem('Reset sync state', 'resetSyncState')
    .addToUi();
}

function bootstrapEmailMonitor() {
  setupEmailMonitor();
  syncMailbox();
  installHourlySyncTrigger();
}

function setupEmailMonitor() {
  const spreadsheet = getTargetSpreadsheet_();
  const logSheet = ensureLogSheet_(spreadsheet);
  const dashboardSheet = ensureDashboardSheet_(spreadsheet);

  configureLogSheet_(logSheet);
  seedDashboard_(dashboardSheet);

  spreadsheet.toast(
    'Email monitor sheets are ready. Use "Sync now" to pull mailbox data.',
    'Email Monitor',
    8
  );
}

function syncMailbox() {
  syncMailboxInternal_({});
}

function backfillLast180Days() {
  syncMailboxInternal_({
    ignoreCheckpoint: true,
    lookbackDays: 180,
  });
}

function installHourlySyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncMailbox') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('syncMailbox')
    .timeBased()
    .everyHours(1)
    .create();

  getTargetSpreadsheet_().toast(
    'Hourly sync trigger installed.',
    'Email Monitor',
    6
  );
}

function resetSyncState() {
  PropertiesService.getScriptProperties().deleteProperty(
    EMAIL_MONITOR_CONFIG.scriptProperties.lastSyncAt
  );

  getTargetSpreadsheet_().toast(
    'Sync checkpoint cleared. The next sync will re-scan the configured window.',
    'Email Monitor',
    8
  );
}

function syncMailboxInternal_(options) {
  const settings = options || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getTargetSpreadsheet_();
    const logSheet = ensureLogSheet_(spreadsheet);
    const dashboardSheet = ensureDashboardSheet_(spreadsheet);
    const existingMessageIds = getExistingMessageIds_(logSheet);
    const query = buildQuery_(settings);
    const rows = [];
    let start = 0;

    while (true) {
      const threads = GmailApp.search(
        query,
        start,
        EMAIL_MONITOR_CONFIG.batchSize
      );

      if (!threads.length) {
        break;
      }

      threads.forEach(function(thread) {
        if (thread.isInTrash() || thread.isInSpam()) {
          return;
        }

        const labels = thread
          .getLabels()
          .map(function(label) {
            return label.getName();
          })
          .sort()
          .join(', ');
        const threadMessageCount = thread.getMessageCount();
        const threadLastMessageDate = thread.getLastMessageDate();
        const threadPermalink = thread.getPermalink();

        thread.getMessages().forEach(function(message) {
          if (shouldSkipMessage_(message)) {
            return;
          }

          const messageId = message.getId();
          if (existingMessageIds.has(messageId)) {
            return;
          }

          const attachmentInfo = getAttachmentInfo_(message);
          rows.push([
            new Date(),
            message.getDate(),
            classifyDirection_(message),
            message.getFrom(),
            message.getTo(),
            message.getCc(),
            message.getReplyTo(),
            message.getSubject(),
            labels,
            message.isUnread(),
            message.isStarred(),
            thread.isImportant(),
            message.isInInbox(),
            message.isInPriorityInbox(),
            attachmentInfo.hasAttachments,
            attachmentInfo.count,
            attachmentInfo.names,
            buildBodyPreview_(message.getPlainBody()),
            threadPermalink,
            thread.getId(),
            messageId,
            safeGetHeader_(message, 'Message-ID'),
            threadMessageCount,
            threadLastMessageDate,
          ]);

          existingMessageIds.add(messageId);
        });
      });

      if (threads.length < EMAIL_MONITOR_CONFIG.batchSize) {
        break;
      }

      start += EMAIL_MONITOR_CONFIG.batchSize;
    }

    if (rows.length) {
      rows.sort(function(left, right) {
        return right[1].getTime() - left[1].getTime();
      });

      appendRows_(logSheet, rows);
      refreshLogSheet_(logSheet);
    }

    seedDashboard_(dashboardSheet);
    PropertiesService.getScriptProperties().setProperty(
      EMAIL_MONITOR_CONFIG.scriptProperties.lastSyncAt,
      new Date().toISOString()
    );

    SpreadsheetApp.flush();
    spreadsheet.toast(
      'Sync complete. ' + rows.length + ' new email(s) logged.',
      'Email Monitor',
      8
    );
  } finally {
    lock.releaseLock();
  }
}

function getTargetSpreadsheet_() {
  return SpreadsheetApp.openById(EMAIL_MONITOR_CONFIG.spreadsheetId);
}

function ensureLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(EMAIL_MONITOR_CONFIG.logSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.logSheetName);
  }
  return sheet;
}

function ensureDashboardSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(EMAIL_MONITOR_CONFIG.dashboardSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.dashboardSheetName);
  }
  return sheet;
}

function configureLogSheet_(sheet) {
  const headerRange = sheet.getRange(
    1,
    1,
    1,
    EMAIL_MONITOR_CONFIG.headers.length
  );
  headerRange.setValues([EMAIL_MONITOR_CONFIG.headers]);
  headerRange
    .setFontWeight('bold')
    .setBackground('#0b57d0')
    .setFontColor('#ffffff')
    .setWrap(true);

  sheet.setFrozenRows(1);
  sheet.setTabColor('#0b57d0');

  EMAIL_MONITOR_CONFIG.columnWidths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.getRange('A:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('X:X').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('R:R').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function seedDashboard_(sheet) {
  sheet.clear();
  sheet.setTabColor('#188038');

  sheet.getRange('A1:B1').setValues([['Mailbox', EMAIL_MONITOR_CONFIG.monitoredMailbox]]);
  sheet.getRange('A2:B2').setValues([['Spreadsheet ID', EMAIL_MONITOR_CONFIG.spreadsheetId]]);
  sheet.getRange('A4:B4').setValues([['Metric', 'Value']]);
  sheet.getRange('A5:A10').setValues([
    ['Total Logged Emails'],
    ['Unread Emails'],
    ['Received Today'],
    ['Sent Today'],
    ['Emails With Attachments'],
    ['Important Emails'],
  ]);
  sheet.getRange('B5:B10').setFormulas([
    ["=MAX(COUNTA('Email Log'!U:U)-1,0)"],
    ["=COUNTIF('Email Log'!J:J,TRUE)"],
    ["=COUNTIFS('Email Log'!B:B,\">=\"&TODAY(),'Email Log'!B:B,\"<\"&TODAY()+1,'Email Log'!C:C,\"Received\")"],
    ["=COUNTIFS('Email Log'!B:B,\">=\"&TODAY(),'Email Log'!B:B,\"<\"&TODAY()+1,'Email Log'!C:C,\"Sent\")"],
    ["=COUNTIF('Email Log'!O:O,TRUE)"],
    ["=COUNTIF('Email Log'!L:L,TRUE)"],
  ]);

  sheet.getRange('D4:E4').setValues([['Top Senders', 'Emails']]);
  sheet.getRange('D5').setFormula(
    "=QUERY('Email Log'!A2:X,\"select D, count(D) where D is not null group by D order by count(D) desc limit 10 label D 'From', count(D) 'Emails'\",0)"
  );

  sheet.getRange('G4:H4').setValues([['Top Subjects', 'Emails']]);
  sheet.getRange('G5').setFormula(
    "=QUERY('Email Log'!A2:X,\"select H, count(H) where H is not null group by H order by count(H) desc limit 10 label H 'Subject', count(H) 'Emails'\",0)"
  );

  sheet.getRange('J4:K4').setValues([['Unread By Sender', 'Unread']]);
  sheet.getRange('J5').setFormula(
    "=QUERY('Email Log'!A2:X,\"select D, count(D) where D is not null and J = TRUE group by D order by count(D) desc limit 10 label D 'From', count(D) 'Unread'\",0)"
  );

  sheet.getRange('A1:K4')
    .setFontWeight('bold')
    .setBackground('#e6f4ea');
  sheet.getRange('A1:K20').setVerticalAlignment('top');
  sheet.setFrozenRows(4);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(7, 320);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(10, 260);
  sheet.setColumnWidth(11, 110);
}

function refreshLogSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const lastColumn = EMAIL_MONITOR_CONFIG.headers.length;
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .sort({ column: 2, ascending: false });
  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
}

function appendRows_(sheet, rows) {
  const startRow = sheet.getLastRow() + 1;
  sheet
    .getRange(startRow, 1, rows.length, EMAIL_MONITOR_CONFIG.headers.length)
    .setValues(rows);
}

function getExistingMessageIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Set();
  }

  const values = sheet.getRange(2, 21, lastRow - 1, 1).getValues();
  return new Set(
    values
      .map(function(row) {
        return String(row[0] || '').trim();
      })
      .filter(function(value) {
        return value !== '';
      })
  );
}

function buildQuery_(settings) {
  const lastSyncAt = PropertiesService.getScriptProperties().getProperty(
    EMAIL_MONITOR_CONFIG.scriptProperties.lastSyncAt
  );
  let startDate;

  if (settings.ignoreCheckpoint && settings.lookbackDays) {
    startDate = new Date();
    startDate.setDate(startDate.getDate() - settings.lookbackDays);
  } else if (lastSyncAt) {
    startDate = new Date(lastSyncAt);
    startDate.setDate(startDate.getDate() - EMAIL_MONITOR_CONFIG.overlapDays);
  } else {
    startDate = new Date();
    startDate.setDate(startDate.getDate() - EMAIL_MONITOR_CONFIG.initialLookbackDays);
  }

  return [
    EMAIL_MONITOR_CONFIG.baseQuery,
    'after:' + Utilities.formatDate(startDate, EMAIL_MONITOR_CONFIG.timeZone, 'yyyy/MM/dd'),
  ].join(' ');
}

function shouldSkipMessage_(message) {
  return (
    message.isDraft() ||
    message.isInChats() ||
    message.isInTrash()
  );
}

function classifyDirection_(message) {
  const mailbox = EMAIL_MONITOR_CONFIG.monitoredMailbox.toLowerCase();
  const from = String(message.getFrom() || '').toLowerCase();
  const recipients = [
    String(message.getTo() || ''),
    String(message.getCc() || ''),
    String(message.getReplyTo() || ''),
  ]
    .join(' ')
    .toLowerCase();

  if (from.indexOf(mailbox) !== -1) {
    return 'Sent';
  }

  if (recipients.indexOf(mailbox) !== -1) {
    return 'Received';
  }

  return 'Related';
}

function getAttachmentInfo_(message) {
  const attachments = message.getAttachments({
    includeInlineImages: false,
    includeAttachments: true,
  });

  return {
    hasAttachments: attachments.length > 0,
    count: attachments.length,
    names: attachments
      .map(function(attachment) {
        return attachment.getName() || '(unnamed attachment)';
      })
      .join(', '),
  };
}

function buildBodyPreview_(plainBody) {
  const normalized = String(plainBody || '')
    .replace(/\s+/g, ' ')
    .trim();

  if (normalized.length <= EMAIL_MONITOR_CONFIG.maxBodyPreviewChars) {
    return normalized;
  }

  return normalized.slice(0, EMAIL_MONITOR_CONFIG.maxBodyPreviewChars - 3) + '...';
}

function safeGetHeader_(message, headerName) {
  try {
    return message.getHeader(headerName);
  } catch (error) {
    return '';
  }
}
