/**
 * ****************************************
 * Google Apps Script - Single Partner Email Summary 2026
 * File: Partner_Summary_Email_Single_2026.js
 * Description: Allows sending a Gemini-generated executive summary for a specific 2026 partner.
 * Supports both Spreadsheet UI (Prompt) and Apps Script Console execution.
 * ****************************************
 */

/**
 * Main entry point for sending a single 2026 partner summary.
 * Can be called from the UI (no args) or programmatically with a partner name.
 * 
 * @param {string} [partnerNameFromArgs] Optional partner name for non-UI execution.
 */
function runSinglePartnerEmailSender2026(partnerNameFromArgs) {
  let partnerNameInput = partnerNameFromArgs;

  if (!partnerNameInput) {
    try {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt(
        'Send Partner Summary Email 2026',
        'Please enter the exact Partner Name as it appears in LATAM_Partner_DB_2026:',
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      partnerNameInput = response.getResponseText().trim();
    } catch (e) {
      Logger.log("ERROR: Spreadsheet UI not available. If running from Console, use runSinglePartnerEmailManual2026().");
      return;
    }
  }

  if (!partnerNameInput) {
    try { SpreadsheetApp.getUi().alert('Partner Name cannot be empty.'); } catch (e) { Logger.log('Partner Name cannot be empty.'); }
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAME_2026);
  if (!dbSheet) {
    try { SpreadsheetApp.getUi().alert(`ERROR: '${SHEET_NAME_2026}' sheet not found.`); } catch (e) { }
    return;
  }

  const dataRange = dbSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());

  // Dynamically find columns
  const colPartnerName = 0; // Partner Name is A (0)
  let colToEmail = -1, colCcEmail = -1, colStatus = -1, colSpreadsheetId = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h.includes("to") && h.includes("email")) colToEmail = c;
    if (h.includes("cc") && h.includes("email")) colCcEmail = c;
    if (h.includes("email") && h.includes("sent")) colStatus = c;
    if (h.includes("spreadsheet") && h.includes("id")) colSpreadsheetId = c;
  }

  if (colToEmail === -1 || colCcEmail === -1 || colSpreadsheetId === -1 || colStatus === -1) {
    const errMsg = "ERROR: Could not find required columns (To Email, CC Email, Spreadsheet ID, Status) in DB.";
    Logger.log(errMsg);
    try { SpreadsheetApp.getUi().alert(errMsg); } catch (e) { }
    return;
  }

  // Aggregate Data for the requested Partner
  let partnerFound = false;
  let spreadsheetId = "";
  const toEmailsSet = new Set();
  const ccEmailsSet = new Set();
  const rowsToUpdate = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partnerNameInSheet = String(row[colPartnerName]).trim();

    if (partnerNameInSheet.toLowerCase() === partnerNameInput.toLowerCase()) {
      partnerFound = true;
      rowsToUpdate.push(i);

      if (!spreadsheetId && row[colSpreadsheetId]) {
        spreadsheetId = String(row[colSpreadsheetId]).trim();
      }

      const to = String(row[colToEmail] || "");
      const cc = String(row[colCcEmail] || "");
      if (to) to.split(',').forEach(e => toEmailsSet.add(e.trim()));
      if (cc) cc.split(',').forEach(e => ccEmailsSet.add(e.trim()));
    }
  }

  if (!partnerFound) {
    const errorMsg = `Partner "${partnerNameInput}" not found in ${SHEET_NAME_2026}.`;
    Logger.log(errorMsg);
    try { SpreadsheetApp.getUi().alert(errorMsg); } catch (e) { }
    return;
  }

  if (!spreadsheetId) {
    const errorMsg = `Partner "${partnerNameInput}" has no Spreadsheet ID. Please generate their deck first.`;
    Logger.log(errorMsg);
    try { SpreadsheetApp.getUi().alert(errorMsg); } catch (e) { }
    return;
  }

  Logger.log(`>>> Processing Single Partner 2026: ${partnerNameInput} <<<`);
  const toEmailStr = Array.from(toEmailsSet).join(",");
  const ccEmailStr = Array.from(ccEmailsSet).join(",");

  try {
    if (ss) ss.toast(`Generating 2026 summary for ${partnerNameInput}...`, "Process Started");
    
    generateAndSendPartnerSummary2026(partnerNameInput, spreadsheetId, toEmailStr, ccEmailStr);

    const currentBatchId = getBatchId2026();
    rowsToUpdate.forEach(r => {
      dbSheet.getRange(r + 1, colStatus + 1).setValue(`MANUAL_${currentBatchId}`);
    });
    SpreadsheetApp.flush();

    Logger.log(`SUCCESS: Summary email sent for ${partnerNameInput}.`);
    try { SpreadsheetApp.getUi().alert(`Success: Summary email sent for ${partnerNameInput}.`); } catch (e) { }
  } catch (e) {
    Logger.log(`ERROR: ${e.toString()}`);
    try { SpreadsheetApp.getUi().alert(`ERROR: ${e.toString()}`); } catch (e) { }
  }
}

/**
 * Helper function for manual execution from the Apps Script console.
 */
function runSinglePartnerEmailManual2026() {
  const testPartnerName = "Accenture"; // Modify this to test a specific partner
  runSinglePartnerEmailSender2026(testPartnerName);
}
