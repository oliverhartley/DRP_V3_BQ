/**
 * ****************************************
 * Google Apps Script - Lock System (Protection)
 * File: Lock_System.gs
 * Version: 1.0
 * Description: Protects specific sheets in partner files to prevent editing.
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

// Sheets to Protect
const SHEETS_TO_LOCK = ["Tier Dashboard", "Profile Deep Dive"];

// Re-using DB Column Indices for Batch Logic
const COL_IDX_MANAGED = 5; 
const COL_IDX_BRAZIL = 7; 
const COL_IDX_MCO = 8;     
const COL_IDX_MEXICO = 9;  

// ==========================================
// 1. BATCH LOCKERS (The Buttons)
// ==========================================

function lockManagedBatch() { runLockBatch(COL_IDX_MANAGED, "MANAGED PARTNERS", true); }
function lockUnManagedBatch() { runLockBatch(COL_IDX_MANAGED, "UNMANAGED PARTNERS", false); }
function lockBrazilBatch() { runLockBatch(COL_IDX_BRAZIL, "Brazil", true); }
function lockMCOBatch() { runLockBatch(COL_IDX_MCO, "MCO", true); }
function lockMexicoBatch() { runLockBatch(COL_IDX_MEXICO, "Mexico", true); }

// ==========================================
// 2. BATCH CONTROLLER
// ==========================================

function runLockBatch(colIndex, batchName, targetValue = true) {
  Logger.log(`>>> STARTING LOCK BATCH: ${batchName} <<<`);
  
  const partnerList = getPartnersForLocking(colIndex, targetValue);
  
  if (partnerList.length === 0) {
    Logger.log(`No partners found for ${batchName}`);
    return;
  }

  Logger.log(`Found ${partnerList.length} partners. Starting protection...`);

  for (let i = 0; i < partnerList.length; i++) {
    const pName = partnerList[i];
    Logger.log(`[${i + 1}/${partnerList.length}] Locking: ${pName}...`);
    
    try {
      const result = lockPartnerFile(pName);
      Logger.log(`  -> ${result}`);
      // Small sleep to be kind to API
      Utilities.sleep(500); 
    } catch (e) {
      Logger.log(`  -> ERROR: ${e.toString()}`);
    }
  }
  
  Logger.log(`>>> LOCK BATCH COMPLETE <<<`);
}

// ==========================================
// 3. CORE LOCKER FUNCTION
// ==========================================

function lockPartnerFile(partnerName) {
  const fileName = `${partnerName} - Partner Dashboard`;
  const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID); 
  const files = folder.getFilesByName(fileName);
  
  if (!files.hasNext()) {
    return "[SKIPPED] File not found.";
  }
  
  const file = files.next();
  const ss = SpreadsheetApp.open(file);
  const me = Session.getEffectiveUser(); // You (The Admin)

  let lockedCount = 0;

  // Iterate through the target sheets
  SHEETS_TO_LOCK.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      
      // 1. Check if already protected, if so, remove old protection to reset
      const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (let i = 0; i < existingProtections.length; i++) {
        existingProtections[i].remove();
      }

      // 2. Create New Protection
      const protection = sheet.protect().setDescription("Locked by Automation");
      
      // 3. Ensure YOU are the only editor
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors()); // Removes everyone else
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false); // Disable "Anyone in domain can edit"
      }
      
      lockedCount++;
    }
  });

  return `[LOCKED] Protected ${lockedCount} sheets.`;
}

// ==========================================
// 4. DATA HELPER
// ==========================================

function getPartnersForLocking(colIndex, targetValue) {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_DB); 
  const data = sheet.getDataRange().getValues();
  const matches = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] === targetValue) {
      matches.push(data[i][1]);
    }
  }
  return matches;
}