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

function ensurePersonnelSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(EMAIL_MONITOR_CONFIG.personnelSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_MONITOR_CONFIG.personnelSheetName);
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
  return (
    headersExistInRow_(
      currentHeaders,
      EMAIL_MONITOR_CONFIG.currentHeadersWithoutPersonnel
    ) ||
    headersExistInRow_(currentHeaders, EMAIL_MONITOR_CONFIG.previousHeaders) ||
    headersExistInRow_(currentHeaders, EMAIL_MONITOR_CONFIG.legacyHeaders)
  );
}

function migrateLogSheetSchema_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(
    sheet.getLastColumn(),
    EMAIL_MONITOR_CONFIG.headers.length,
    EMAIL_MONITOR_CONFIG.currentHeadersWithoutPersonnel.length,
    EMAIL_MONITOR_CONFIG.previousHeaders.length,
    EMAIL_MONITOR_CONFIG.legacyHeaders.length
  );
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
      return getMigratedCellValue_(header, currentHeaders, row);
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

function headersExistInRow_(currentHeaders, expectedHeaders) {
  return expectedHeaders.every(function(header) {
    return currentHeaders.indexOf(header) !== -1;
  });
}

function getMigratedCellValue_(header, currentHeaders, row) {
  const sourceIndex = currentHeaders.indexOf(header);
  if (sourceIndex !== -1) {
    return header === 'Status'
      ? normalizeEmailStatusValue_(row[sourceIndex])
      : row[sourceIndex];
  }

  if (header === 'Status') {
    const legacyStatusIndex = currentHeaders.indexOf('With Reply');
    return legacyStatusIndex === -1
      ? EMAIL_MONITOR_CONFIG.defaultStatus
      : normalizeEmailStatusValue_(row[legacyStatusIndex]);
  }

  return '';
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
  sheet
    .getRange(
      1,
      EMAIL_LOG_COLUMN_INDEX.personnel,
      sheet.getMaxRows(),
      1
    )
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet
    .getRange(
      1,
      EMAIL_LOG_COLUMN_INDEX.statusUpdate,
      sheet.getMaxRows(),
      1
    )
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet
    .getRange(1, EMAIL_LOG_COLUMN_INDEX.status, sheet.getMaxRows(), 1)
    .setHorizontalAlignment('center');
  sheet
    .getRange(1, EMAIL_LOG_COLUMN_INDEX.personnel)
    .setNote(
      'Choose one OPSD personnel name from the dropdown, then choose additional names to build a multi-select list. Clear the cell to remove all selections.'
    );
  applyPersonnelColumnValidation_(sheet);
  applyStatusColumnValidation_(sheet);
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
    ['Completed'],
    ['Open Items'],
    ['Received Today'],
    ['Received This Week'],
  ]);
  sheet.getRange('B5:B9').setFormulas([
    ["=MAX(COUNTA('Email Log'!I:I)-1,0)"],
    ["=COUNTIF('Email Log'!K:K,\"Completed\")"],
    ["=COUNTIFS('Email Log'!K:K,\"<>Completed\",'Email Log'!K:K,\"<>\")"],
    ["=COUNTIFS('Email Log'!B:B,\">=\"&TODAY(),'Email Log'!B:B,\"<\"&TODAY()+1)"],
    ["=COUNTIFS('Email Log'!B:B,\">=\"&(TODAY()-WEEKDAY(TODAY(),2)+1),'Email Log'!B:B,\"<\"&(TODAY()-WEEKDAY(TODAY(),2)+8))"],
  ]);

  sheet.getRange('D4:E4').setValues([['Top Senders', 'Emails']]);
  sheet.getRange('D5').setFormula(
    "=QUERY('Email Log'!A2:L,\"select C, count(C) where C is not null group by C order by count(C) desc limit 10 label C 'From', count(C) 'Emails'\",0)"
  );

  sheet.getRange('G4:H4').setValues([['Open Items By Sender', 'Emails']]);
  sheet.getRange('G5').setFormula(
    "=QUERY('Email Log'!A2:L,\"select C, count(C) where C is not null and K <> 'Completed' and K is not null group by C order by count(C) desc limit 10 label C 'From', count(C) 'Emails'\",0)"
  );

  sheet.getRange('J4:K4').setValues([['Common Subjects', 'Emails']]);
  sheet.getRange('J5').setFormula(
    "=QUERY('Email Log'!A2:L,\"select F, count(F) where F is not null group by F order by count(F) desc limit 10 label F 'Subject', count(F) 'Emails'\",0)"
  );

  sheet.getRange('A1:L4')
    .setFontWeight('bold')
    .setBackground('#e6f4ea');
  sheet.getRange('A1:L20').setVerticalAlignment('top');
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
  sheet.getRange('A2:L2').setValues([
    ['Choose a sender in B1 to view matching emails from Email Log.', '', '', '', '', '', '', '', '', '', '', ''],
  ]);
  sheet.getRange('A4:L4').setValues([EMAIL_MONITOR_CONFIG.headers]);
  sheet.getRange('A5').setFormula(
    "=IF($B$1=\"\",\"\",IFERROR(FILTER('Email Log'!A2:L,'Email Log'!C2:C=$B$1),\"\"))"
  );
  sheet.getRange('M1').setValue('Sender List');
  sheet.getRange('M2').setFormula(
    "=IFERROR(SORT(UNIQUE(FILTER('Email Log'!C2:C,'Email Log'!C2:C<>\"\"))),\"\")"
  );

  sheet.getRange('A1:L1')
    .setFontWeight('bold')
    .setBackground('#fef7e0');
  sheet.getRange('A2:L2')
    .merge()
    .setWrap(true)
    .setBackground('#fff8d7');
  sheet.getRange('A4:L4')
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
  sheet.getRange('J:J').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('L:L').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('K:K').setHorizontalAlignment('center');
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

function configurePersonnelSheet_(sheet) {
  const headerRange = sheet.getRange(
    1,
    1,
    1,
    EMAIL_MONITOR_CONFIG.personnelHeaders.length
  );

  headerRange.setValues([EMAIL_MONITOR_CONFIG.personnelHeaders]);
  headerRange
    .setFontWeight('bold')
    .setBackground('#5f6368')
    .setFontColor('#ffffff')
    .setWrap(true);

  sheet.setFrozenRows(1);
  sheet.setTabColor('#5f6368');

  EMAIL_MONITOR_CONFIG.personnelColumnWidths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('F:F').setNumberFormat('0');
  sheet.getRange('G:G').setHorizontalAlignment('center');
  sheet.getRange('A:G').setVerticalAlignment('middle');
  sheet
    .getRange(1, OPSD_PERSONNEL_COLUMN_INDEX.isActive)
    .setNote('Use TRUE/FALSE or Yes/No. Blank values are treated as active.');
}

function getPersonnelRecords_() {
  const sheet = ensurePersonnelSheet_(getTargetSpreadsheet_());
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, EMAIL_MONITOR_CONFIG.personnelHeaders.length)
    .getValues()
    .map(function(row) {
      return {
        userName: normalizePersonnelOptionValue_(
          row[OPSD_PERSONNEL_COLUMN_INDEX.userName - 1]
        ),
        division: normalizePersonnelOptionValue_(
          row[OPSD_PERSONNEL_COLUMN_INDEX.division - 1]
        ),
        position: normalizePersonnelOptionValue_(
          row[OPSD_PERSONNEL_COLUMN_INDEX.position - 1]
        ),
        sortOrder: normalizePersonnelSortOrder_(
          row[OPSD_PERSONNEL_COLUMN_INDEX.sortOrder - 1]
        ),
        isActive: normalizePersonnelIsActive_(
          row[OPSD_PERSONNEL_COLUMN_INDEX.isActive - 1]
        ),
      };
    })
    .filter(function(record) {
      return record.userName !== '';
    });
}

function getPersonnelDirectory_() {
  const directory = {};

  getPersonnelRecords_()
    .slice()
    .sort(function(left, right) {
      return (
        left.sortOrder - right.sortOrder ||
        left.userName.localeCompare(right.userName)
      );
    })
    .forEach(function(record) {
      const key = record.userName.toLowerCase();
      if (!record.userName || directory[key]) {
        return;
      }

      directory[key] = {
        userName: record.userName,
        division: record.division,
        position: record.position,
      };
    });

  return directory;
}

function getActivePersonnelOptions_() {
  const seen = {};
  return getPersonnelRecords_()
    .filter(function(record) {
      return record.isActive;
    })
    .sort(function(left, right) {
      return (
        left.sortOrder - right.sortOrder ||
        left.userName.localeCompare(right.userName)
      );
    })
    .map(function(record) {
      return record.userName;
    })
    .filter(function(userName) {
      const key = userName.toLowerCase();
      if (seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
}

function normalizePersonnelOptionValue_(value) {
  return String(value || '').trim();
}

function normalizePersonnelSortOrder_(value) {
  const normalizedValue = String(value || '').trim();
  if (normalizedValue === '') {
    return Number.MAX_SAFE_INTEGER;
  }

  const numericValue = Number(normalizedValue);
  return isNaN(numericValue) ? Number.MAX_SAFE_INTEGER : numericValue;
}

function normalizePersonnelIsActive_(value) {
  if (value === false) {
    return false;
  }
  if (value === true || value === 1) {
    return true;
  }

  const normalizedValue = String(value || '').trim().toLowerCase();
  if (!normalizedValue) {
    return true;
  }

  return (
    ['false', '0', 'no', 'n', 'inactive'].indexOf(normalizedValue) === -1
  );
}

function splitPersonnelAssignments_(value) {
  const seen = {};
  return String(value || '')
    .split(/\s*,\s*|\s*;\s*|\r?\n+/)
    .map(function(item) {
      return String(item || '').trim();
    })
    .filter(function(item) {
      const key = item.toLowerCase();
      if (!item || seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
}

function normalizePersonnelAssignmentsInput_(value, options) {
  const optionList = options || [];
  const optionLookup = optionList.reduce(function(map, option) {
    map[option.toLowerCase()] = option;
    return map;
  }, {});
  const rawValues = Array.isArray(value)
    ? value
    : splitPersonnelAssignments_(value);
  const seen = {};

  return sortPersonnelAssignments_(
    rawValues
      .map(function(item) {
        const normalizedItem = String(item || '').trim();
        if (!normalizedItem) {
          return '';
        }

        return optionLookup[normalizedItem.toLowerCase()] || normalizedItem;
      })
      .filter(function(item) {
        const key = item.toLowerCase();
        if (!item || seen[key]) {
          return false;
        }
        seen[key] = true;
        return true;
      }),
    optionList
  );
}

function sortPersonnelAssignments_(values, options) {
  const optionOrder = (options || []).reduce(function(map, option, index) {
    map[option.toLowerCase()] = index;
    return map;
  }, {});

  return values.slice().sort(function(left, right) {
    const leftKey = left.toLowerCase();
    const rightKey = right.toLowerCase();
    const leftOrder = Object.prototype.hasOwnProperty.call(optionOrder, leftKey)
      ? optionOrder[leftKey]
      : Number.MAX_SAFE_INTEGER;
    const rightOrder = Object.prototype.hasOwnProperty.call(optionOrder, rightKey)
      ? optionOrder[rightKey]
      : Number.MAX_SAFE_INTEGER;

    return leftOrder - rightOrder || left.localeCompare(right);
  });
}

function joinPersonnelAssignments_(values) {
  return normalizePersonnelAssignmentsInput_(
    values,
    getActivePersonnelOptions_()
  ).join(EMAIL_MONITOR_CONFIG.personnelAssignmentSeparator);
}

function togglePersonnelAssignment_(
  existingValues,
  selectedValue,
  options
) {
  const normalizedSelection = normalizePersonnelAssignmentsInput_(
    [selectedValue],
    options
  )[0];

  if (!normalizedSelection) {
    return normalizePersonnelAssignmentsInput_(existingValues, options);
  }

  const nextValues = normalizePersonnelAssignmentsInput_(existingValues, options);
  const existingIndex = nextValues.findIndex(function(value) {
    return value.toLowerCase() === normalizedSelection.toLowerCase();
  });

  if (existingIndex === -1) {
    nextValues.push(normalizedSelection);
  } else {
    nextValues.splice(existingIndex, 1);
  }

  return sortPersonnelAssignments_(nextValues, options);
}

function applyPersonnelColumnValidation_(sheet) {
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  const validationRange = sheet.getRange(
    2,
    EMAIL_LOG_COLUMN_INDEX.personnel,
    rowCount,
    1
  );
  const options = getActivePersonnelOptions_();

  validationRange.clearDataValidations();
  if (!options.length) {
    return;
  }

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(true)
    .build();

  validationRange.setDataValidation(validation);
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    return;
  }

  const sheet = range.getSheet();
  if (sheet.getName() === EMAIL_MONITOR_CONFIG.personnelSheetName) {
    if (range.getRow() > 1) {
      applyPersonnelColumnValidation_(ensureLogSheet_(sheet.getParent()));
    }
    return;
  }

  if (
    sheet.getName() !== EMAIL_MONITOR_CONFIG.logSheetName ||
    range.getRow() <= 1 ||
    range.getColumn() !== EMAIL_LOG_COLUMN_INDEX.personnel
  ) {
    return;
  }

  applyPersonnelMultiSelectEdit_(range, e.value, e.oldValue);
}

function applyPersonnelMultiSelectEdit_(range, newValue, oldValue) {
  if (typeof newValue === 'undefined') {
    return;
  }

  const options = getActivePersonnelOptions_();
  const previousValues = normalizePersonnelAssignmentsInput_(oldValue, options);
  const incomingValues = normalizePersonnelAssignmentsInput_(newValue, options);
  let nextValues = [];

  if (!incomingValues.length) {
    range.clearContent();
    return;
  }

  if (!previousValues.length || incomingValues.length > 1) {
    nextValues = incomingValues;
  } else {
    nextValues = togglePersonnelAssignment_(
      previousValues,
      incomingValues[0],
      options
    );
  }

  range.setValue(
    nextValues.join(EMAIL_MONITOR_CONFIG.personnelAssignmentSeparator)
  );
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

  applyPersonnelColumnValidation_(sheet);
  applyStatusColumnValidation_(sheet);
  sheet.getRange(2, 1, lastRow - 1, lastColumn).sort({
    column: EMAIL_LOG_COLUMN_INDEX.dateReceived,
    ascending: true,
  });
  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
}

function applyStatusColumnValidation_(sheet) {
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(EMAIL_MONITOR_CONFIG.statusOptions, true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(2, EMAIL_LOG_COLUMN_INDEX.status, rowCount, 1)
    .setDataValidation(validation);
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
