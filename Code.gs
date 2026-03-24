/**
 * Code.gs
 * Main entry points: doGet(), onOpen(), onEdit(), includes()
 */

/**
 * Serves the student form as a web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AI Use Documentation')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Adds custom menu when the spreadsheet opens.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AI Documenter')
    .addItem('Setup Wizard', 'showSetupWizard')
    .addSeparator()
    .addItem('Sync Assignments from Canvas', 'syncAllAssignments')
    .addItem('Retry Failed Processing', 'retryFailedProcessing')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('Reformat Tabs', 'reformatTabs')
    .addToUi();
}

/**
 * Handles edits to the Dashboard tab for interactive filtering.
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== 'Dashboard') return;

    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();

    // React to changes in the dropdown cells (B1, B2, B3)
    if (col === 2 && row >= 1 && row <= 3) {
      // If view type changed, repopulate the student/assignment dropdown
      if (row === 1) {
        populateDashboardSelectDropdown();
      } else if (row === 2) {
        populateDashboardSelectDropdown();
      } else if (row === 3) {
        refreshDashboard();
      }
    }
  } catch (err) {
    // Silently fail on edit triggers to avoid disrupting the user
    Logger.log('onEdit error: ' + err.message);
  }
}

/**
 * Opens the setup wizard dialog.
 */
function showSetupWizard() {
  var html = HtmlService.createHtmlOutputFromFile('setup')
    .setWidth(600)
    .setHeight(500)
    .setTitle('AI Documenter Setup');
  SpreadsheetApp.getUi().showModalDialog(html, 'AI Documenter Setup');
}
