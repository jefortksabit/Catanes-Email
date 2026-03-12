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
        dateReceived: row[0] instanceof Date ? row[0] : new Date(row[0]),
        from: String(row[1] || '').trim(),
        to: String(row[2] || '').trim(),
        cc: String(row[3] || '').trim(),
        subject: String(row[4] || '').trim(),
        message: String(row[5] || '').trim(),
        threadId: String(row[6] || '').trim(),
        messageId: String(row[7] || '').trim(),
        withReply: row[8] === true,
      };
    })
    .filter(function(record) {
      return record.dateReceived && !isNaN(record.dateReceived.getTime());
    })
    .sort(function(left, right) {
      return right.dateReceived.getTime() - left.dateReceived.getTime();
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
  const replyStatus = String(filters.replyStatus || 'all').trim();

  return {
    sender: String(filters.sender || '').trim(),
    query: String(filters.query || '').trim(),
    replyStatus:
      replyStatus === 'with_reply' || replyStatus === 'pending_reply'
        ? replyStatus
        : 'all',
    limit: limit,
  };
}

function recordMatchesFilters_(record, filters) {
  if (filters.sender && record.from !== filters.sender) {
    return false;
  }

  if (filters.replyStatus === 'with_reply' && !record.withReply) {
    return false;
  }

  if (filters.replyStatus === 'pending_reply' && record.withReply) {
    return false;
  }

  if (!filters.query) {
    return true;
  }

  const haystack = [
    record.from,
    record.to,
    record.cc,
    record.subject,
    record.message,
    record.threadId,
    record.messageId,
  ]
    .join(' ')
    .toLowerCase();

  return haystack.indexOf(filters.query.toLowerCase()) !== -1;
}

function formatEmailRecordForWeb_(record) {
  return {
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
    withReply: record.withReply,
  };
}

function buildWebSummary_(records) {
  const todayKey = Utilities.formatDate(
    new Date(),
    EMAIL_MONITOR_CONFIG.timeZone,
    'yyyy-MM-dd'
  );
  const uniqueSenders = new Set();
  let withReplyCount = 0;
  let receivedToday = 0;

  records.forEach(function(record) {
    if (record.from) {
      uniqueSenders.add(record.from);
    }

    if (record.withReply) {
      withReplyCount += 1;
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
    withReply: withReplyCount,
    pendingReply: records.length - withReplyCount,
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
