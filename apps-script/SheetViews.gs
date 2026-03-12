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

  if (canMigrateLogSheetSchema_(sheet)) {
    migrateLogSheetSchema_(sheet);
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

function ensureSenderViewSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(EMAIL_MONITOR_CONFIG.senderViewSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.senderViewSheetName);
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

  const currentHeaders = getSheetHeaderValues_(sheet);

  return EMAIL_MONITOR_CONFIG.headers.every(function(header, index) {
    return currentHeaders[index] === header;
  });
}

function canMigrateLogSheetSchema_(sheet) {
  if (sheet.getLastRow() < 1) {
    return true;
  }

  const currentHeaders = getSheetHeaderValues_(sheet);
  return EMAIL_MONITOR_CONFIG.legacyHeaders.every(function(header) {
    return currentHeaders.indexOf(header) !== -1;
  });
}

function migrateLogSheetSchema_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), EMAIL_MONITOR_CONFIG.legacyHeaders.length);
  const currentHeaders = getSheetHeaderValues_(sheet);
  const currentRows =
    lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
      : [];
  const filter = sheet.getFilter();

  if (filter) {
    filter.remove();
  }

  const migratedRows = currentRows.map(function(row) {
    return EMAIL_MONITOR_CONFIG.headers.map(function(header) {
      const sourceIndex = currentHeaders.indexOf(header);
      if (sourceIndex === -1) {
        return header === 'With Reply' ? false : '';
      }
      return row[sourceIndex];
    });
  });

  sheet.clear();
  configureLogSheet_(sheet);

  if (migratedRows.length) {
    appendRows_(sheet, migratedRows);
    refreshLogSheet_(sheet);
  }
}

function getSheetHeaderValues_(sheet) {
  const headerWidth = Math.max(
    sheet.getLastColumn(),
    EMAIL_MONITOR_CONFIG.headers.length
  );
  if (headerWidth < 1) {
    return [];
  }

  return sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
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

  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('G:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('K:K').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
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
    ["=MAX(COUNTA('Email Log'!I:I)-1,0)"],
    ["=COUNTIF('Email Log'!J:J,TRUE)"],
    ["=COUNTIF('Email Log'!J:J,FALSE)"],
    ["=COUNTIFS('Email Log'!B:B,\">=\"&TODAY(),'Email Log'!B:B,\"<\"&TODAY()+1)"],
    ["=COUNTIFS('Email Log'!B:B,\">=\"&(TODAY()-WEEKDAY(TODAY(),2)+1),'Email Log'!B:B,\"<\"&(TODAY()-WEEKDAY(TODAY(),2)+8))"],
  ]);

  sheet.getRange('D4:E4').setValues([['Top Senders', 'Emails']]);
  sheet.getRange('D5').setFormula(
    "=QUERY('Email Log'!A2:K,\"select C, count(C) where C is not null group by C order by count(C) desc limit 10 label C 'From', count(C) 'Emails'\",0)"
  );

  sheet.getRange('G4:H4').setValues([['Pending Reply By Sender', 'Emails']]);
  sheet.getRange('G5').setFormula(
    "=QUERY('Email Log'!A2:K,\"select C, count(C) where C is not null and J = FALSE group by C order by count(C) desc limit 10 label C 'From', count(C) 'Emails'\",0)"
  );

  sheet.getRange('J4:K4').setValues([['Common Subjects', 'Emails']]);
  sheet.getRange('J5').setFormula(
    "=QUERY('Email Log'!A2:K,\"select F, count(F) where F is not null group by F order by count(F) desc limit 10 label F 'Subject', count(F) 'Emails'\",0)"
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

function seedSenderView_(sheet) {
  const selectedSender = String(sheet.getRange('B1').getValue() || '').trim();

  sheet.clear();
  sheet.getRange('A:M').breakApart();
  sheet.setTabColor('#f9ab00');

  sheet.getRange('A1').setValue('Sender');
  sheet.getRange('B1:D1').merge().setValue('');
  sheet.getRange('A2:K2').setValues([
    ['Choose a sender in B1 to view matching emails from Email Log.', '', '', '', '', '', '', '', '', '', ''],
  ]);
  sheet.getRange('A4:K4').setValues([EMAIL_MONITOR_CONFIG.headers]);
  sheet.getRange('A5').setFormula(
    "=IF($B$1=\"\",\"\",IFERROR(FILTER('Email Log'!A2:K,'Email Log'!C2:C=$B$1),\"\"))"
  );
  sheet.getRange('M1').setValue('Sender List');
  sheet.getRange('M2').setFormula(
    "=IFERROR(SORT(UNIQUE(FILTER('Email Log'!C2:C,'Email Log'!C2:C<>\"\"))),\"\")"
  );

  sheet.getRange('A1:K1')
    .setFontWeight('bold')
    .setBackground('#fef7e0');
  sheet.getRange('A2:K2')
    .merge()
    .setWrap(true)
    .setBackground('#fff8d7');
  sheet.getRange('A4:K4')
    .setFontWeight('bold')
    .setBackground('#f9ab00')
    .setFontColor('#202124');

  sheet.setFrozenRows(4);
  EMAIL_MONITOR_CONFIG.columnWidths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('G:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('K:K').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.hideColumns(13);

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('M2:M'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B1').setDataValidation(validation);

  SpreadsheetApp.flush();
  const availableSenders = sheet
    .getRange('M2:M')
    .getDisplayValues()
    .map(function(row) {
      return String(row[0] || '').trim();
    })
    .filter(function(value) {
      return value !== '';
    });

  const senderToUse =
    availableSenders.indexOf(selectedSender) !== -1
      ? selectedSender
      : (availableSenders[0] || '');

  if (senderToUse) {
    sheet.getRange('B1').setValue(senderToUse);
  }
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
    column: EMAIL_LOG_COLUMN_INDEX.dateReceived,
    ascending: true,
  });
  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
}

function applyCheckboxColumn_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  sheet
    .getRange(2, EMAIL_LOG_COLUMN_INDEX.withReply, lastRow - 1, 1)
    .insertCheckboxes();
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

  const values = sheet
    .getRange(2, EMAIL_LOG_COLUMN_INDEX.messageId, lastRow - 1, 1)
    .getValues();
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
