/**
 * ****************************************
 * Google Apps Script - Link Caching System
 * File: Link_System.gs
 * Version: 2.1 (Added Clear Cache)
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

/**
 * 1. AUTO-SAVE: Called by the Batch Generator.
 * Saves a list of [{name, url}] to the cache instantly.
 */
function saveBatchLinks(linkArray) {
  if (!linkArray || linkArray.length === 0) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_LINKS);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_LINKS); sheet.hideSheet(); }
  
  // Get existing data to append/update
  const lastRow = sheet.getLastRow();
  let existingMap = new Map();
  
  if (lastRow > 1) {
    const currentData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    currentData.forEach(row => existingMap.set(row[0], row[1]));
  } else {
    sheet.getRange("A1").setValue("Partner Name");
    sheet.getRange("B1").setValue("Dashboard URL");
  }
  
  // Merge new links
  linkArray.forEach(item => {
    existingMap.set(item.name, item.url);
  });
  
  // Convert back to array and sort
  const output = Array.from(existingMap.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  
  // Write back in one go
  sheet.getRange(2, 1, output.length, 2).setValues(output);
}

/**
 * 2. MANUAL SCAN: Optimized with Search Query (Faster than iteration)
 */
function updateLinkCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_LINKS);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_LINKS); sheet.hideSheet(); }
  
  sheet.clear();
  sheet.getRange("A1:B1").setValues([["Partner Name", "Dashboard URL"]]);

  // Search Query: Much faster than getting all files
  const query = `'${PARTNER_FOLDER_ID}' in parents and title contains ' - Partner Dashboard' and trashed = false`;
  const files = DriveApp.searchFiles(query);
  
  const output = [];
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().replace(" - Partner Dashboard", "").trim();
    output.push([name, file.getUrl()]);
  }

  if (output.length > 0) {
    output.sort((a, b) => a[0].localeCompare(b[0]));
    sheet.getRange(2, 1, output.length, 2).setValues(output);
  }
  return output.length;
}

/**
 * 3. CLEAR CACHE: Removes all cached links.
 */
function clearLinkCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_LINKS);
  if (sheet) {
    sheet.clear();
    sheet.getRange("A1:B1").setValues([["Partner Name", "Dashboard URL"]]);
  }
}