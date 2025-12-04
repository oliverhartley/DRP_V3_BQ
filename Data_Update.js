/**
 * ****************************************
 * Google Apps Script - Full Data Update
 * File: Data_Update.gs
 * Version: 1.0
 * ****************************************
 */

function runFullDataUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. Update Partner DB
    ss.toast("Starting Step 1 of 5: Updating Partner DB...", "Full Update", 5);
    runBigQueryQuery();
    ss.toast("Step 1 Complete.", "Full Update", 5);
    
    // 2. Update Scoring Matrix
    ss.toast("Starting Step 2 of 5: Updating Scoring Matrix...", "Full Update", 5);
    runPartnerScorePivot();
    ss.toast("Step 2 Complete.", "Full Update", 5);
    
    // 3. Update Profile Source
    ss.toast("Starting Step 3 of 5: Updating Profile Source...", "Full Update", 5);
    runDeepDiveQuerySource();
    ss.toast("Step 3 Complete.", "Full Update", 5);

    // 4. Update Link Cache
    ss.toast("Starting Step 4 of 5: Updating Link Cache...", "Full Update", 5);
    const linkCount = updateLinkCache();
    ss.toast(`Step 4 Complete. Cached ${linkCount} links.`, "Full Update", 5);

    // 5. Update Dashboard Cache
    ss.toast("Starting Step 5 of 5: Updating Dashboard Cache...", "Full Update", 5);
    updateDashboardCache();
    ss.toast("Step 5 Complete.", "Full Update", 5);

    ss.toast("Full Data Update Complete!", "Success", 5);
    
  } catch (e) {
    ss.toast("Error during Full Update: " + e.toString(), "Error", 10);
    Logger.log("Full Update Error: " + e.toString());
  }
}
