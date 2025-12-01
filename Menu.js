/**
 * ****************************************
 * Google Apps Script - Custom Menu
 * File: Menu.gs
 * Version: 12.0 (Added Locking System)
 * ****************************************
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 Partner Engine')
      .addItem('1️⃣ Update Partner DB', 'runBigQueryQuery')
      .addItem('2️⃣ Update Scoring Matrix', 'runPartnerScorePivot')
      .addItem('3️⃣ Update Profile Source', 'runDeepDiveQuerySource')
      .addSeparator()
      
      .addSubMenu(ui.createMenu('📄 Generate Decks')
          .addItem('⭐ MANAGED Partners', 'runManagedBatch')
        .addItem('🌍 GSI Partners', 'runGSIBatch') // Added GSI
          .addItem('📂 UNMANAGED Partners', 'runUnManagedBatch')
          .addSeparator()
          .addItem('🇧🇷 Brazil', 'runBrazilBatch')
          .addItem('🇲🇽 Mexico', 'runMexicoBatch')
        .addItem('🌎 MCO', 'runMCOBatch')
        .addItem('💼 PS', 'runPSBatch'))
          
      .addSubMenu(ui.createMenu('🔒 Lock Decks')
          .addItem('⭐ Lock MANAGED', 'lockManagedBatch')
        .addItem('🌍 Lock GSI', 'lockGSIBatch') // Added GSI
          .addItem('📂 Lock UNMANAGED', 'lockUnManagedBatch')
          .addSeparator()
          .addItem('🇧🇷 Lock Brazil', 'lockBrazilBatch')
          .addItem('🇲🇽 Lock Mexico', 'lockMexicoBatch')
        .addItem('🌎 Lock MCO', 'lockMCOBatch')
        .addItem('💼 Lock PS', 'lockPSBatch'))
      
      .addSeparator()
      .addItem('🔗 Refresh Links (Manual)', 'runLinkUpdateManual') 
      .addItem('⚠️ Reset Dropdowns', 'setupDashboard')
      .addItem('🕒 Timestamp', 'updateTimestamp')
      .addToUi();
}

// ... (Keep the rest of your Menu.gs functions: runLinkUpdateManual, updateTimestamp) ...
// Make sure you keep the helper functions at the bottom of this file!
function runLinkUpdateManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Scanning Drive for files...", "Update Started", 5);
    const count = updateLinkCache(); 
    ss.toast(`Found ${count} partner files. Slicer is ready.`, "Update Complete", 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function updateTimestamp() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
    if (!sheet) return;
    const now = new Date();
    const timeString = "Last Data Refresh: " + Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm");
    const targetRange = sheet.getRange("E1:I1");
    targetRange.merge().setValue(timeString).setBackground('#fff2cc').setFontColor('#666666').setHorizontalAlignment('right').setFontWeight('bold');
    SpreadsheetApp.flush();
  } catch (e) {}
}