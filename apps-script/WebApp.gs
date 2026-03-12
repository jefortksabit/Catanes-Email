function doGet() {
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.appBootstrap = {
    monitoredMailbox: EMAIL_MONITOR_CONFIG.monitoredMailbox,
    baselineDateLabel: 'February 1, 2026',
    spreadsheetId: EMAIL_MONITOR_CONFIG.spreadsheetId,
    statusOptions: EMAIL_MONITOR_CONFIG.statusOptions,
  };

  return template
    .evaluate()
    .setTitle('Catanes Email Monitor')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
