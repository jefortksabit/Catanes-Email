const EMAIL_MONITOR_CONFIG = Object.freeze({
  monitoredMailbox: 'jcatanes@ched.gov.ph',
  spreadsheetId: '1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY',
  logSheetName: 'Email Log',
  dashboardSheetName: 'Dashboard',
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email Monitor')
    .addItem('Bootstrap monitor', 'bootstrapEmailMonitor')
    .addItem('Setup sheets only', 'setupEmailMonitor')
    .addItem('Sync now', 'syncMailbox')
    .addItem('Resync from Feb 1, 2026', 'resyncFromStartDate')
    .addSeparator()
    .addItem('Install hourly trigger', 'installHourlySyncTrigger')
    .addItem('Reset sync state', 'resetSyncState')
    .addToUi();
}

function bootstrapEmailMonitor() {
  setupEmailMonitor();
  installHourlySyncTrigger();
  syncMailbox();
}

function setupEmailMonitor() {
  const spreadsheet = getTargetSpreadsheet_();
  const logSheet = ensureLogSheet_(spreadsheet);
  const dashboardSheet = ensureDashboardSheet_(spreadsheet);

  configureLogSheet_(logSheet);
  refreshLogSheet_(logSheet);
  seedDashboard_(dashboardSheet);

  spreadsheet.toast(
    'Email monitor sheets are ready. Run Sync now to log inbound emails.',
    'Email Monitor',
    8
  );
}

function syncMailbox() {
  syncMailboxInternal_({});
}

function resyncFromStartDate() {
  syncMailboxInternal_({
    forceStartDate: true,
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
    'Sync checkpoint cleared. The next run will re-scan from February 1, 2026.',
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

        const threadMessages = thread.getMessages();
        const replyDates = getReplyDatesFromMailbox_(threadMessages);

        threadMessages.forEach(function(message) {
          if (!shouldLogMessage_(message)) {
            return;
          }

          const messageId = message.getId();
          if (existingMessageIds.has(messageId)) {
            return;
          }

          rows.push([
            message.getDate(),
            message.getFrom(),
            message.getTo(),
            message.getCc(),
            message.getSubject(),
            buildProcessedMessage_(message),
            thread.getId(),
            messageId,
            hasReplyAfterMessage_(message.getDate(), replyDates),
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
        return right[0].getTime() - left[0].getTime();
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
      'Sync complete. ' + rows.length + ' inbound email(s) logged.',
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
    return spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.logSheetName);
  }

  if (isSheetSchemaCurrent_(sheet)) {
    return sheet;
  }

  if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
    sheet.setName(
      EMAIL_MONITOR_CONFIG.logSheetName +
        ' Backup ' +
        Utilities.formatDate(
          new Date(),
          EMAIL_MONITOR_CONFIG.timeZone,
          'yyyyMMdd_HHmmss'
        )
    );
    return spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.logSheetName);
  }

  sheet.clear();
  return sheet;
}

function ensureDashboardSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(EMAIL_MONITOR_CONFIG.dashboardSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.dashboardSheetName);
  }
  return sheet;
}

function isSheetSchemaCurrent_(sheet) {
  if (sheet.getLastRow() < 1) {
    return true;
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn !== EMAIL_MONITOR_CONFIG.headers.length) {
    return false;
  }

  const currentHeaders = sheet
    .getRange(1, 1, 1, EMAIL_MONITOR_CONFIG.headers.length)
    .getValues()[0];

  return EMAIL_MONITOR_CONFIG.headers.every(function(header, index) {
    return currentHeaders[index] === header;
  });
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

  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('F:F').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  applyCheckboxColumn_(sheet);
}

function seedDashboard_(sheet) {
  sheet.clear();
  sheet.setTabColor('#188038');

  sheet.getRange('A1:B2').setValues([
    ['Mailbox', EMAIL_MONITOR_CONFIG.monitoredMailbox],
    ['Spreadsheet ID', EMAIL_MONITOR_CONFIG.spreadsheetId],
  ]);

  sheet.getRange('A4:B4').setValues([['Metric', 'Value']]);
  sheet.getRange('A5:A9').setValues([
    ['Total Logged Emails'],
    ['With Reply'],
    ['Pending Reply'],
    ['Received Today'],
    ['Received This Week'],
  ]);
  sheet.getRange('B5:B9').setFormulas([
    ["=MAX(COUNTA('Email Log'!H:H)-1,0)"],
    ["=COUNTIF('Email Log'!I:I,TRUE)"],
    ["=COUNTIF('Email Log'!I:I,FALSE)"],
    ["=COUNTIFS('Email Log'!A:A,\">=\"&TODAY(),'Email Log'!A:A,\"<\"&TODAY()+1)"],
    ["=COUNTIFS('Email Log'!A:A,\">=\"&(TODAY()-WEEKDAY(TODAY(),2)+1),'Email Log'!A:A,\"<\"&(TODAY()-WEEKDAY(TODAY(),2)+8))"],
  ]);

  sheet.getRange('D4:E4').setValues([['Top Senders', 'Emails']]);
  sheet.getRange('D5').setFormula(
    "=QUERY('Email Log'!A2:I,\"select B, count(B) where B is not null group by B order by count(B) desc limit 10 label B 'From', count(B) 'Emails'\",0)"
  );

  sheet.getRange('G4:H4').setValues([['Pending Reply By Sender', 'Emails']]);
  sheet.getRange('G5').setFormula(
    "=QUERY('Email Log'!A2:I,\"select B, count(B) where B is not null and I = FALSE group by B order by count(B) desc limit 10 label B 'From', count(B) 'Emails'\",0)"
  );

  sheet.getRange('J4:K4').setValues([['Common Subjects', 'Emails']]);
  sheet.getRange('J5').setFormula(
    "=QUERY('Email Log'!A2:I,\"select E, count(E) where E is not null group by E order by count(E) desc limit 10 label E 'Subject', count(E) 'Emails'\",0)"
  );

  sheet.getRange('A1:K4')
    .setFontWeight('bold')
    .setBackground('#e6f4ea');
  sheet.getRange('A1:K20').setVerticalAlignment('top');
  sheet.setFrozenRows(4);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(7, 280);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(10, 280);
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

  applyCheckboxColumn_(sheet);
  sheet.getRange(2, 1, lastRow - 1, lastColumn).sort({
    column: 1,
    ascending: false,
  });
  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
}

function applyCheckboxColumn_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  sheet.getRange(2, 9, lastRow - 1, 1).insertCheckboxes();
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

  const values = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
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
  const baselineDate = getInitialSyncStartDate_();
  let startDate = baselineDate;

  if (!settings.forceStartDate && lastSyncAt) {
    startDate = new Date(lastSyncAt);
    startDate.setDate(startDate.getDate() - EMAIL_MONITOR_CONFIG.overlapDays);
    if (startDate.getTime() < baselineDate.getTime()) {
      startDate = baselineDate;
    }
  }

  return [
    EMAIL_MONITOR_CONFIG.baseQuery,
    'after:' +
      Utilities.formatDate(
        startDate,
        EMAIL_MONITOR_CONFIG.timeZone,
        'yyyy/MM/dd'
      ),
  ].join(' ');
}

function getInitialSyncStartDate_() {
  return new Date(EMAIL_MONITOR_CONFIG.initialSyncStartDate + 'T00:00:00+08:00');
}

function shouldLogMessage_(message) {
  return (
    !message.isDraft() &&
    !message.isInChats() &&
    !message.isInTrash() &&
    !messageAddressContainsMailbox_(message.getFrom())
  );
}

function getReplyDatesFromMailbox_(threadMessages) {
  return threadMessages
    .filter(function(message) {
      return (
        !message.isDraft() &&
        !message.isInTrash() &&
        messageAddressContainsMailbox_(message.getFrom())
      );
    })
    .map(function(message) {
      return message.getDate().getTime();
    })
    .sort(function(left, right) {
      return left - right;
    });
}

function hasReplyAfterMessage_(messageDate, replyDates) {
  const messageTime = messageDate.getTime();
  return replyDates.some(function(replyTime) {
    return replyTime > messageTime;
  });
}

function messageAddressContainsMailbox_(value) {
  return String(value || '')
    .toLowerCase()
    .indexOf(EMAIL_MONITOR_CONFIG.monitoredMailbox.toLowerCase()) !== -1;
}

function buildProcessedMessage_(message) {
  const cleanedBody = cleanEmailBody_(message.getPlainBody());
  const summarySource = cleanedBody || String(message.getSubject() || '').trim();
  return buildFallbackMessage_(summarySource);
}

function cleanEmailBody_(plainBody) {
  let cleaned = String(plainBody || '').replace(/\r\n/g, '\n').trim();
  const markers = [
    /^On .+wrote:$/im,
    /^From:\s.+$/im,
    /^Sent:\s.+$/im,
    /^-+\s*Original Message\s*-+$/im,
    /^Begin forwarded message:$/im,
  ];

  markers.forEach(function(marker) {
    const match = cleaned.match(marker);
    if (match && typeof match.index === 'number') {
      cleaned = cleaned.slice(0, match.index).trim();
    }
  });

  return cleaned.replace(/\n{3,}/g, '\n\n').replace(/[ \t]+/g, ' ').trim();
}

function buildFallbackMessage_(text) {
  const sourceText = String(text || '').trim();
  if (!sourceText) {
    return 'No message content.';
  }

  return truncate_(
    sourceText.replace(/\s+/g, ' '),
    EMAIL_MONITOR_CONFIG.maxFallbackChars
  );
}

function truncate_(value, maxChars) {
  const text = String(value || '').trim();
  if (text.length <= maxChars) {
    return text;
  }

  return text.slice(0, maxChars - 3).trim() + '...';
}
