/**
 * ****************************************
 * Google Apps Script - Single Partner Email Summary
 * File: Partner_Summary_Email_Single.gs
 * Description: Allows sending a Gemini-generated executive summary for a specific partner.
 * Supports both Spreadsheet UI (Prompt) and Apps Script Console execution.
 * ****************************************
 */

/**
 * Main entry point for sending a single partner summary.
 * Can be called from the UI (no args) or programmatically with a partner name.
 * 
 * @param {string} [partnerNameFromArgs] Optional partner name for non-UI execution.
 */
function runSinglePartnerEmailSender(partnerNameFromArgs) {
  let partnerNameInput = partnerNameFromArgs;

  // Use UI prompt if no name is provided via arguments
  if (!partnerNameInput) {
    try {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt(
        'Send Partner Summary Email',
        'Please enter the exact Partner Name as it appears in the "Consolidate by Partner" sheet:',
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      partnerNameInput = response.getResponseText().trim();
    } catch (e) {
      Logger.log("ERROR: Spreadsheet UI not available. If running from Console, use runSinglePartnerEmailManual('Partner Name').");
      return;
    }
  }

  if (!partnerNameInput) {
    try { SpreadsheetApp.getUi().alert('Partner Name cannot be empty.'); } catch (e) { Logger.log('Partner Name cannot be empty.'); }
    return;
  }

  const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sheet = ss.getSheetByName("Consolidate by Partner");
  if (!sheet) {
    try { SpreadsheetApp.getUi().alert("ERROR: 'Consolidate by Partner' sheet not found."); } catch (e) { Logger.log("ERROR: 'Consolidate by Partner' sheet not found."); }
    return;
  }

  const data = sheet.getDataRange().getValues();
  let partnerFound = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partnerNameInSheet = row[COL_PARTNER_NAME];

    if (String(partnerNameInSheet).trim().toLowerCase() === partnerNameInput.toLowerCase()) {
      partnerFound = true;
      const toEmails = row[COL_TO_EMAIL];
      const ccEmails = row[COL_CC_EMAIL];

      Logger.log(`>>> Processing Single Partner: ${partnerNameInSheet} <<<`);
      
      const fileId = findPartnerFileId(partnerNameInSheet);
      if (fileId) {
        try {
          if (ss) ss.toast(`Generating summary for ${partnerNameInSheet}...`, "Process Started");
          generateAndSendPartnerSummary(partnerNameInSheet, fileId, toEmails, ccEmails);

          const currentBatchId = getBatchId();
          sheet.getRange(i + 1, COL_STATUS + 1).setValue(`MANUAL_${currentBatchId}`);
          
          Logger.log(`SUCCESS: Summary email sent for ${partnerNameInSheet}.`);
          try { SpreadsheetApp.getUi().alert(`Success: Summary email sent for ${partnerNameInSheet}.`); } catch (e) { }
        } catch (e) {
          Logger.log(`ERROR: ${e.toString()}`);
          try { SpreadsheetApp.getUi().alert(`ERROR: ${e.toString()}`); } catch (e) { }
        }
      } else {
        const errorMsg = `WARNING: Partner file not found for ${partnerNameInSheet}.`;
        Logger.log(errorMsg);
        try { SpreadsheetApp.getUi().alert(errorMsg); } catch (e) { }
      }
      break;
    }
  }

  if (!partnerFound) {
    const errorMsg = `Partner "${partnerNameInput}" not found in the "Consolidate by Partner" sheet.`;
    Logger.log(errorMsg);
    try { SpreadsheetApp.getUi().alert(errorMsg); } catch (e) { }
  }
}

/**
 * Helper function for manual execution from the Apps Script console.
 * Usage: Select this function and click "Run", or modify the hardcoded name.
 */
function runSinglePartnerEmailManual() {
  // EDIT THE PARTNER NAME BELOW FOR CONSOLE TESTING
  const testPartnerName = "Accenture";
  runSinglePartnerEmailSender(testPartnerName);
}
