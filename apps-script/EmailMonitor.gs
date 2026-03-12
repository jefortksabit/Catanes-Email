const EMAIL_MONITOR_CONFIG = Object.freeze({
  monitoredMailbox: 'jcatanes@ched.gov.ph',
  spreadsheetId: '1TYsnQlu8S4CslR42Y18A2d4JD_EGKVmpMEtUug-WqWY',
  logSheetName: 'Email Log',
  dashboardSheetName: 'Dashboard',
  baseQuery: 'in:anywhere -in:trash -in:spam',
  initialLookbackDays: 30,
  overlapDays: 2,
  batchSize: 100,
  timeZone: 'Asia/Manila',
  maxBodyCharsForGemini: 12000,
  maxFallbackChars: 280,
  geminiModel: 'gemini-2.5-flash',
  geminiApiEndpoint:
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent',
  scriptProperties: {
    lastSyncAt: 'EMAIL_MONITOR_LAST_SYNC_AT',
    geminiApiKey: 'GEMINI_API_KEY',
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
    .addItem('Backfill last 180 days', 'backfillLast180Days')
    .addSeparator()
    .addItem('Set Gemini API key', 'setGeminiApiKey')
    .addItem('Clear Gemini API key', 'clearGeminiApiKey')
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
    'Email monitor sheets are ready. Set the Gemini API key, then run Sync now.',
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
    'Sync checkpoint cleared. The next run will re-scan the recent window.',
    'Email Monitor',
    8
  );
}

function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Gemini API Key',
    'Paste the Gemini API key from Google AI Studio. It will be stored in Apps Script script properties.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const apiKey = String(response.getResponseText() || '').trim();
  if (!apiKey) {
    ui.alert('No API key was provided.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(
    EMAIL_MONITOR_CONFIG.scriptProperties.geminiApiKey,
    apiKey
  );

  getTargetSpreadsheet_().toast(
    'Gemini API key saved.',
    'Email Monitor',
    6
  );
}

function clearGeminiApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(
    EMAIL_MONITOR_CONFIG.scriptProperties.geminiApiKey
  );

  getTargetSpreadsheet_().toast(
    'Gemini API key cleared. Message summaries will use the fallback text.',
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
    const geminiApiKey = getGeminiApiKey_();
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
            buildProcessedMessage_(message, geminiApiKey),
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
      buildSyncToastMessage_(rows.length, Boolean(geminiApiKey)),
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

  const geminiStatus = getGeminiApiKey_()
    ? 'Configured (' + EMAIL_MONITOR_CONFIG.geminiModel + ')'
    : 'Not configured';

  sheet.getRange('A1:B3').setValues([
    ['Mailbox', EMAIL_MONITOR_CONFIG.monitoredMailbox],
    ['Spreadsheet ID', EMAIL_MONITOR_CONFIG.spreadsheetId],
    ['Gemini Summary', geminiStatus],
  ]);

  sheet.getRange('A5:B5').setValues([['Metric', 'Value']]);
  sheet.getRange('A6:A10').setValues([
    ['Total Logged Emails'],
    ['With Reply'],
    ['Pending Reply'],
    ['Received Today'],
    ['Received This Week'],
  ]);
  sheet.getRange('B6:B10').setFormulas([
    ["=MAX(COUNTA('Email Log'!H:H)-1,0)"],
    ["=COUNTIF('Email Log'!I:I,TRUE)"],
    ["=COUNTIF('Email Log'!I:I,FALSE)"],
    ["=COUNTIFS('Email Log'!A:A,\">=\"&TODAY(),'Email Log'!A:A,\"<\"&TODAY()+1)"],
    ["=COUNTIFS('Email Log'!A:A,\">=\"&(TODAY()-WEEKDAY(TODAY(),2)+1),'Email Log'!A:A,\"<\"&(TODAY()-WEEKDAY(TODAY(),2)+8))"],
  ]);

  sheet.getRange('D5:E5').setValues([['Top Senders', 'Emails']]);
  sheet.getRange('D6').setFormula(
    "=QUERY('Email Log'!A2:I,\"select B, count(B) where B is not null group by B order by count(B) desc limit 10 label B 'From', count(B) 'Emails'\",0)"
  );

  sheet.getRange('G5:H5').setValues([['Pending Reply By Sender', 'Emails']]);
  sheet.getRange('G6').setFormula(
    "=QUERY('Email Log'!A2:I,\"select B, count(B) where B is not null and I = FALSE group by B order by count(B) desc limit 10 label B 'From', count(B) 'Emails'\",0)"
  );

  sheet.getRange('J5:K5').setValues([['Common Subjects', 'Emails']]);
  sheet.getRange('J6').setFormula(
    "=QUERY('Email Log'!A2:I,\"select E, count(E) where E is not null group by E order by count(E) desc limit 10 label E 'Subject', count(E) 'Emails'\",0)"
  );

  sheet.getRange('A1:K5')
    .setFontWeight('bold')
    .setBackground('#e6f4ea');
  sheet.getRange('A1:K20').setVerticalAlignment('top');
  sheet.setFrozenRows(5);
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
  let startDate;

  if (settings.ignoreCheckpoint && settings.lookbackDays) {
    startDate = new Date();
    startDate.setDate(startDate.getDate() - settings.lookbackDays);
  } else if (lastSyncAt) {
    startDate = new Date(lastSyncAt);
    startDate.setDate(startDate.getDate() - EMAIL_MONITOR_CONFIG.overlapDays);
  } else {
    startDate = new Date();
    startDate.setDate(
      startDate.getDate() - EMAIL_MONITOR_CONFIG.initialLookbackDays
    );
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

function buildProcessedMessage_(message, geminiApiKey) {
  const cleanedBody = cleanEmailBody_(message.getPlainBody());
  const fallbackText = buildFallbackMessage_(cleanedBody, message.getSubject());

  if (!geminiApiKey) {
    return fallbackText;
  }

  const prompt = [
    'Summarize this email for a monitoring spreadsheet.',
    'Return plain text only.',
    'Keep it to at most 80 words.',
    'Include the main request, key facts, dates, names, and the next expected action if any.',
    'Ignore greetings, signatures, disclaimers, and quoted thread history.',
    '',
    'From: ' + String(message.getFrom() || ''),
    'To: ' + String(message.getTo() || ''),
    'Cc: ' + String(message.getCc() || ''),
    'Subject: ' + String(message.getSubject() || ''),
    '',
    'Email body:',
    truncate_(cleanedBody || '(No message body)', EMAIL_MONITOR_CONFIG.maxBodyCharsForGemini),
  ].join('\n');

  const summary = generateGeminiSummary_(prompt, geminiApiKey);
  return summary || fallbackText;
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

function buildFallbackMessage_(cleanedBody, subject) {
  const sourceText = String(cleanedBody || '').trim() || String(subject || '').trim();
  if (!sourceText) {
    return 'No message content.';
  }

  return truncate_(sourceText.replace(/\s+/g, ' '), EMAIL_MONITOR_CONFIG.maxFallbackChars);
}

function generateGeminiSummary_(prompt, apiKey) {
  try {
    const payload = {
      contents: [
        {
          role: 'user',
          parts: [{ text: prompt }],
        },
      ],
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 160,
        responseMimeType: 'text/plain',
      },
    };

    const response = UrlFetchApp.fetch(EMAIL_MONITOR_CONFIG.geminiApiEndpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-goog-api-key': apiKey,
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      console.warn(
        'Gemini summary failed with status ' + response.getResponseCode()
      );
      return '';
    }

    const data = JSON.parse(response.getContentText());
    const candidates = data.candidates || [];
    if (!candidates.length) {
      return '';
    }

    const parts = (((candidates[0] || {}).content || {}).parts || []).map(
      function(part) {
        return part.text || '';
      }
    );

    return truncate_(parts.join(' ').replace(/\s+/g, ' ').trim(), 400);
  } catch (error) {
    console.warn('Gemini summary error: ' + error);
    return '';
  }
}

function truncate_(value, maxChars) {
  const text = String(value || '').trim();
  if (text.length <= maxChars) {
    return text;
  }

  return text.slice(0, maxChars - 3).trim() + '...';
}

function getGeminiApiKey_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(
      EMAIL_MONITOR_CONFIG.scriptProperties.geminiApiKey
    ) || ''
  ).trim();
}

function buildSyncToastMessage_(newRows, geminiEnabled) {
  if (geminiEnabled) {
    return 'Sync complete. ' + newRows + ' inbound email(s) logged.';
  }

  return (
    'Sync complete. ' +
    newRows +
    ' inbound email(s) logged. Gemini key not set, so Message uses fallback text.'
  );
}
