/**
 * ****************************************
 * Google Apps Script - Automated Triggers
 * File: Triggers.js
 * Description: Daily 1AM Sync for BigQuery Caches.
 * ****************************************
 */

/**
 * Setup the 1AM daily trigger.
 * Run this ONCE manually or via the Menu.
 */
function setupDailySync() {
  // Clear existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailySync') {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Create new trigger: 1 AM - 2 AM
  ScriptApp.newTrigger('runDailySync')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  Logger.log("Daily 1AM Trigger for 'runDailySync' established.");
}

/**
 * The Main Sync Function.
 * Wraps all automated updates.
 */
function runDailySync() {
  Logger.log(">>> STARTING DAILY SYNC <<<");
  
  // 1. Refresh Scoring Cache
  try {
    Logger.log("Step 1: Refreshing Scoring Cache...");
    runScoringLoader(); // Defined in Scoring.js
  } catch (e) {
    Logger.log("ERROR in Scoring Loader: " + e.toString());
  }

  // 2. Refresh Deep Dive Cache (To be implemented)
  try {
    Logger.log("Step 2: Refreshing Deep Dive Cache...");
    // runDeepDiveLoader(); 
    Logger.log("Deep Dive Loader pending implementation.");
  } catch (e) {
    Logger.log("ERROR in Deep Dive Loader: " + e.toString());
  }
  
  Logger.log(">>> COMPLETED DAILY SYNC <<<");
}
