function getWebAppBootstrapData(filters) {
  const records = getEmailLogRecordsForWeb_();
  return buildWebAppPayload_(records, filters || {});
}

function getFilteredEmailRecords(filters) {
  const records = getEmailLogRecordsForWeb_();
  return buildFilteredEmailPayload_(records, filters || {});
}

function getEmailLogRecordsForWeb_() {
  const sheet = getTargetSpreadsheet_().getSheetByName(
    EMAIL_MONITOR_CONFIG.logSheetName
  );
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, EMAIL_MONITOR_CONFIG.headers.length)
    .getValues();

  return values
    .map(function(row) {
      return {
        referenceNumber: String(
          row[EMAIL_LOG_COLUMN_INDEX.referenceNumber - 1] || ''
        ).trim(),
        dateReceived:
          row[EMAIL_LOG_COLUMN_INDEX.dateReceived - 1] instanceof Date
            ? row[EMAIL_LOG_COLUMN_INDEX.dateReceived - 1]
            : new Date(row[EMAIL_LOG_COLUMN_INDEX.dateReceived - 1]),
        from: String(row[EMAIL_LOG_COLUMN_INDEX.from - 1] || '').trim(),
        to: String(row[EMAIL_LOG_COLUMN_INDEX.to - 1] || '').trim(),
        cc: String(row[EMAIL_LOG_COLUMN_INDEX.cc - 1] || '').trim(),
        subject: String(row[EMAIL_LOG_COLUMN_INDEX.subject - 1] || '').trim(),
        message: String(row[EMAIL_LOG_COLUMN_INDEX.message - 1] || '').trim(),
        threadId: String(row[EMAIL_LOG_COLUMN_INDEX.threadId - 1] || '').trim(),
        messageId: String(
          row[EMAIL_LOG_COLUMN_INDEX.messageId - 1] || ''
        ).trim(),
        status: normalizeEmailStatusValue_(
          row[EMAIL_LOG_COLUMN_INDEX.status - 1]
        ),
        statusUpdate: String(
          row[EMAIL_LOG_COLUMN_INDEX.statusUpdate - 1] || ''
        ).trim(),
      };
    })
    .filter(function(record) {
      return record.dateReceived && !isNaN(record.dateReceived.getTime());
    })
    .sort(function(left, right) {
      return left.dateReceived.getTime() - right.dateReceived.getTime();
    });
}

function buildWebAppPayload_(records, filters) {
  const filteredPayload = buildFilteredEmailPayload_(records, filters);
  return Object.assign(filteredPayload, {
    summary: buildWebSummary_(records),
    senderOptions: getSenderOptions_(records),
    topSenders: buildTopSenders_(records),
    metadata: {
      lastSyncAt: getLastSyncLabel_(),
      totalRows: records.length,
      baselineDateLabel: 'February 1, 2026',
    },
  });
}

function buildFilteredEmailPayload_(records, filters) {
  const normalizedFilters = normalizeWebFilters_(filters);
  const filteredRecords = records.filter(function(record) {
    return recordMatchesFilters_(record, normalizedFilters);
  });

  return {
    filters: normalizedFilters,
    totalFiltered: filteredRecords.length,
    rows: filteredRecords
      .slice(0, normalizedFilters.limit)
      .map(formatEmailRecordForWeb_),
  };
}

function normalizeWebFilters_(filters) {
  const limit = Math.min(
    Math.max(parseInt(filters.limit, 10) || WEB_APP_CONFIG.defaultRowLimit, 25),
    WEB_APP_CONFIG.maxRowLimit
  );
  const status = String(filters.status || 'all').trim();

  return {
    sender: String(filters.sender || '').trim(),
    query: String(filters.query || '').trim(),
    status:
      EMAIL_MONITOR_CONFIG.statusOptions.indexOf(status) !== -1 ? status : 'all',
    limit: limit,
  };
}

function recordMatchesFilters_(record, filters) {
  if (filters.sender && record.from !== filters.sender) {
    return false;
  }

  if (filters.status !== 'all' && record.status !== filters.status) {
    return false;
  }

  if (!filters.query) {
    return true;
  }

  const haystack = [
    record.referenceNumber,
    record.from,
    record.to,
    record.cc,
    record.subject,
    record.message,
    record.threadId,
    record.messageId,
    record.status,
    record.statusUpdate,
  ]
    .join(' ')
    .toLowerCase();

  return haystack.indexOf(filters.query.toLowerCase()) !== -1;
}

function formatEmailRecordForWeb_(record) {
  return {
    referenceNumber: record.referenceNumber,
    dateReceived: Utilities.formatDate(
      record.dateReceived,
      EMAIL_MONITOR_CONFIG.timeZone,
      'yyyy-MM-dd HH:mm'
    ),
    from: record.from,
    to: record.to,
    cc: record.cc,
    subject: record.subject || '(No subject)',
    message: record.message || 'No message content.',
    threadId: record.threadId,
    messageId: record.messageId,
    status: record.status,
    statusUpdate: record.statusUpdate,
  };
}

function buildWebSummary_(records) {
  const todayKey = Utilities.formatDate(
    new Date(),
    EMAIL_MONITOR_CONFIG.timeZone,
    'yyyy-MM-dd'
  );
  const uniqueSenders = new Set();
  let completedCount = 0;
  let openCount = 0;
  let receivedToday = 0;

  records.forEach(function(record) {
    if (record.from) {
      uniqueSenders.add(record.from);
    }

    if (isOpenEmailStatus_(record.status)) {
      openCount += 1;
    } else {
      completedCount += 1;
    }

    const recordDay = Utilities.formatDate(
      record.dateReceived,
      EMAIL_MONITOR_CONFIG.timeZone,
      'yyyy-MM-dd'
    );
    if (recordDay === todayKey) {
      receivedToday += 1;
    }
  });

  return {
    totalEmails: records.length,
    uniqueSenders: uniqueSenders.size,
    completed: completedCount,
    openItems: openCount,
    receivedToday: receivedToday,
  };
}

function getSenderOptions_(records) {
  return Array.from(
    records.reduce(function(senderSet, record) {
      if (record.from) {
        senderSet.add(record.from);
      }
      return senderSet;
    }, new Set())
  ).sort();
}

function buildTopSenders_(records) {
  const counts = records.reduce(function(map, record) {
    if (record.from) {
      map[record.from] = (map[record.from] || 0) + 1;
    }
    return map;
  }, {});

  return Object.keys(counts)
    .map(function(sender) {
      return {
        sender: sender,
        count: counts[sender],
      };
    })
    .sort(function(left, right) {
      return right.count - left.count || left.sender.localeCompare(right.sender);
    })
    .slice(0, 8);
}

function getLastSyncLabel_() {
  const rawValue = PropertiesService.getScriptProperties().getProperty(
    EMAIL_MONITOR_CONFIG.scriptProperties.lastSyncAt
  );
  if (!rawValue) {
    return 'Not synced yet';
  }

  return Utilities.formatDate(
    new Date(rawValue),
    EMAIL_MONITOR_CONFIG.timeZone,
    'yyyy-MM-dd HH:mm'
  );
}

function updateEmailRecordManualFields(payload) {
  const messageId = String((payload && payload.messageId) || '').trim();
  if (!messageId) {
    throw new Error('Message ID is required to update the row.');
  }

  const referenceNumber = normalizeManualSheetValue_(
    payload && payload.referenceNumber
  );
  const status = normalizeEmailStatusValue_(payload && payload.status);
  const statusUpdate = normalizeManualSheetValue_(
    payload && payload.statusUpdate
  );
  const sheet = ensureLogSheet_(getTargetSpreadsheet_());
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    throw new Error('No email records are available to update.');
  }

  const messageIds = sheet
    .getRange(2, EMAIL_LOG_COLUMN_INDEX.messageId, lastRow - 1, 1)
    .getDisplayValues();
  let matchedRow = 0;

  messageIds.some(function(row, index) {
    if (String(row[0] || '').trim() === messageId) {
      matchedRow = index + 2;
      return true;
    }
    return false;
  });

  if (!matchedRow) {
    throw new Error('The selected email could not be found in the sheet.');
  }

  sheet
    .getRange(matchedRow, EMAIL_LOG_COLUMN_INDEX.referenceNumber)
    .setValue(referenceNumber);
  sheet.getRange(matchedRow, EMAIL_LOG_COLUMN_INDEX.status).setValue(status);
  sheet
    .getRange(matchedRow, EMAIL_LOG_COLUMN_INDEX.statusUpdate)
    .setValue(statusUpdate);
  SpreadsheetApp.flush();

  return {
    messageId: messageId,
    referenceNumber: referenceNumber,
    status: status,
    statusUpdate: statusUpdate,
  };
}

function normalizeManualSheetValue_(value) {
  return String(value || '').trim();
}
