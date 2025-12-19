/**
 * ****************************************
 * Google Apps Script - Single Partner Email Summary
 * File: Partner_Summary_Email_Single.gs
 * Description: Allows sending a Gemini-generated executive summary for a specific partner.
 * ****************************************
 */

/**
 * Prompts the user for a partner name and sends the summary email.
 * This reuses the logic from Partner_Summary_Email.js but targets one specific row.
 */
function runSinglePartnerEmailSender() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Send Partner Summary Email',
    'Please enter the exact Partner Name as it appears in the "Consolidate by Partner" sheet:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const partnerNameInput = response.getResponseText().trim();
  if (!partnerNameInput) {
    ui.alert('Partner Name cannot be empty.');
    return;
  }

  const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sheet = ss.getSheetByName("Consolidate by Partner");
  if (!sheet) {
    ui.alert("ERROR: 'Consolidate by Partner' sheet not found in Source SS.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  let partnerFound = false;

  // Search for the partner in the sheet
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partnerNameInSheet = row[COL_PARTNER_NAME];
    
    // Case-insensitive comparison just in case, but strict preferred
    if (String(partnerNameInSheet).trim().toLowerCase() === partnerNameInput.toLowerCase()) {
      partnerFound = true;
      const toEmails = row[COL_TO_EMAIL];
      const ccEmails = row[COL_CC_EMAIL];

      Logger.log(`Processing Single Partner: ${partnerNameInSheet}`);
      
      const fileId = findPartnerFileId(partnerNameInSheet);
      if (fileId) {
        try {
          ss.toast(`Generating summary for ${partnerNameInSheet}...`, "Process Started");
          generateAndSendPartnerSummary(partnerNameInSheet, fileId, toEmails, ccEmails);
          
          // Optionally update status column to record that this specific email was sent
          const currentBatchId = getBatchId();
          sheet.getRange(i + 1, COL_STATUS + 1).setValue(`MANUAL_${currentBatchId}`);
          
          ui.alert(`Success: Summary email sent for ${partnerNameInSheet}.`);
        } catch (e) {
          ui.alert(`ERROR: ${e.toString()}`);
        }
      } else {
        ui.alert(`WARNING: Partner file not found for ${partnerNameInSheet}. Please check the Partner Folder ID.`);
      }
      break;
    }
  }

  if (!partnerFound) {
    ui.alert(`Partner "${partnerNameInput}" not found in the "Consolidate by Partner" sheet.`);
  }
}
