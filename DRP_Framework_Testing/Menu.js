/**
 * ****************************************
 * Google Apps Script - User Interface
 * File: Menu.js
 * Description: Refresh of the dashboard menu and navigation.
 * ****************************************
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 DRP Framework V3')
    .addItem('1. Initialize Empty System', 'menuInitSystem')
    .addSeparator()
    .addSubMenu(ui.createMenu('2. Data Operations')
      .addItem('Refresh Partner DB (BQ)', 'menuRefreshDB')
      .addItem('Refresh Scoring Matrix', 'menuRefreshScoring')
      .addItem('Refresh Deep Dive Data', 'menuRefreshDeepDive')
      .addItem('Run Full Update', 'menuFullUpdate'))
    .addSubMenu(ui.createMenu('3. Automation & Triggers')
      .addItem('✅ Setup Daily 1AM Sync', 'menuSetupTrigger')
      .addItem('❌ Remove All Triggers', 'menuRemoveTriggers'))
    .addSeparator()
    .addItem('⚙️ System Migration (Master -> Local)', 'runMigration')
    .addItem('✨ Enrich from BigQuery (Add Missing)', 'syncBigQueryToLocalDB')
    .addItem('🌍 Sync Country Presence (Profile Data)', 'enrichPartnerCountries')
    .addToUi();
}

/**
 * UI Wrapper for Initialization
 */
function menuInitSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('System Initialization', 'This will create/reset:\n1. DB_Partners (Managed)\n2. DB_Reference\n\nContinue?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    initSystem();
    ui.alert('Initialization complete. Please run "System Migration" next.');
  }
}

// UI Route to modular local functions
function menuRefreshDB() { runBigQueryLoader(); }
function menuRefreshScoring() { runScoringLoader(); }
function menuRefreshDeepDive() { runDeepDiveLoader(); }
function menuFullUpdate() { runDailySync(); } // Full Update = Daily Sync logic

function menuSetupTrigger() {
  setupDailySync();
  SpreadsheetApp.getUi().alert("Daily 1AM Trigger Established.");
}

function menuRemoveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) ScriptApp.deleteTrigger(t);
  SpreadsheetApp.getUi().alert("All triggers removed.");
}

function menuBatchEmail() {
  SpreadsheetApp.getUi().alert("Batch Email logic not yet implemented in V3.");
}
