/**
 * ****************************************
 * Google Apps Script - Custom Menu
 * File: Menu.gs
 * Version: 2026.1 (Cleaned and Refactored)
 * ****************************************
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 Partner Engine 2026')
    .addItem('1️⃣ Update Deep Dive Data', 'runDeepDive2026')
    .addItem('2️⃣ Update Scoring Matrix', 'runPartnerScorePivot2026')
    .addSeparator()
    .addItem('📊 Format Dashboard View', 'setupDashboard2026')
    .addSeparator()
    .addItem('📄 Generate All Partner Decks', 'runFullBatchDecks2026')
    .addItem('📧 Send Batch Summary Emails', 'runBatchEmailSender2026')
    .addItem('📨 Send Single Partner Email (Test)', 'runSinglePartnerEmailSender2026')
    .addToUi();
}