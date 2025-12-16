/**
 * ****************************************
 * Google Apps Script - UI & Menu
 * File: Menu.js
 * Description: Application Entry Point & Custom Menu.
 * NOTE: UI Disabled for Console-First Debugging.
 * ****************************************
 */

/*
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
*/

/**
 * UI Wrapper for Initialization
 */
function menuInitSystem() {
  // Console Version
  Logger.log("Running initSystem() from console...");
  initSystem();
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
