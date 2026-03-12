function syncMailboxInternal_(options) {
  const settings = options || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getTargetSpreadsheet_();
    const logSheet = ensureLogSheet_(spreadsheet);
    const dashboardSheet = ensureDashboardSheet_(spreadsheet);
    const senderViewSheet = ensureSenderViewSheet_(spreadsheet);
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
    seedSenderView_(senderViewSheet);
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
