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
  const senderViewSheet = ensureSenderViewSheet_(spreadsheet);

  configureLogSheet_(logSheet);
  refreshLogSheet_(logSheet);
  seedDashboard_(dashboardSheet);
  seedSenderView_(senderViewSheet);

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
