/**
 * ****************************************
 * Google Apps Script - Performance Analysis
 * File: Performance_Analysis.js
 * Description: Calculates the difference in partner scores between two snapshots.
 * ****************************************
 */

function calculatePerformanceDelta() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const baselineSheetName = "LATAM_Partner_Score_DRP_Nov1"; // Corrected name based on logs
  const currentSheetName = SHEET_NAME_SCORE; // "LATAM_Partner_Score_DRP" from Config.gs
  const outputSheetName = "Performance Delta";

  try {
    ss.toast('Starting performance delta calculation...', 'In Progress', 10);
    Logger.log('Starting performance delta calculation...');

    // 1. Get sheet objects
    const baselineSheet = ss.getSheetByName(baselineSheetName);
    const currentSheet = ss.getSheetByName(currentSheetName);

    if (!baselineSheet || !currentSheet) {
      const allSheets = ss.getSheets().map(s => s.getName());
      Logger.log("Available Sheets: " + allSheets.join(", "));
      throw new Error(`Could not find baseline ('${baselineSheetName}') or current ('${currentSheetName}') sheet. Available: ${allSheets.join(", ")}`);
    }

    // 2. Read data from both sheets
    // 2. Read data from sheets
    const baselineData = baselineSheet.getDataRange().getValues();
    const currentData = currentSheet.getDataRange().getValues();
    const dbData = ss.getSheetByName(SHEET_NAME_DB).getDataRange().getValues(); // Get DB data for Total Profiles

    const headers = currentData.slice(0, 3); // Keep the top 3 header rows

    // 3a. Create a map of the baseline data
    const baselineMap = new Map();
    baselineData.slice(3).forEach(row => {
      const partnerId = row[0];
      baselineMap.set(partnerId, row);
    });

    // 3b. Create a map of the DB data to get the correct Total Profiles count
    const dbMap = new Map();
    const dbHeaders = dbData[0];
    const totalProfilesIdx = dbHeaders.indexOf("Total_Profiles");
    if (totalProfilesIdx === -1) throw new Error("Total_Profiles column not found in DB sheet.");

    dbData.slice(1).forEach(row => {
      const partnerId = row[0]; // Assuming Partner ID is col 0 in DB too? Let's check.
      // Actually, in DB sheet, Partner Name is col 1, ID is likely col 0 but let's be safe.
      // Based on Partner_Region_Solution_Selector.gs: idxName = 1.
      // Let's assume Partner Name is the key if ID is missing or unreliable, but ID is safer.
      // Let's assume ID is col 0 for now as per standard.
      dbMap.set(row[1], row[totalProfilesIdx]); // Map Name -> Total Profiles (Name is safer for joining if IDs differ)
    });

    // 4. Calculate the delta
    const deltaData = [];
    currentData.slice(3).forEach(currentRow => {
      const partnerId = currentRow[0];
      const partnerName = currentRow[1];
      const baselineRow = baselineMap.get(partnerId);

      // Get the "Real" Total Profiles from DB if available, otherwise fallback to Score sheet
      let currentTotalProfiles = currentRow[2]; // Default from Score sheet
      if (dbMap.has(partnerName)) {
        currentTotalProfiles = dbMap.get(partnerName);
      }

      const deltaRow = [partnerId, partnerName];

      if (baselineRow) {
        // Total Profiles Delta (Col 2)
        const baselineTotalProfiles = parseFloat(baselineRow[2]) || 0;
        deltaRow.push(currentTotalProfiles - baselineTotalProfiles);

        // Other Columns Delta (Col 3+)
        for (let i = 3; i < currentRow.length; i++) {
          const currentValue = parseFloat(currentRow[i]) || 0;
          const baselineValue = parseFloat(baselineRow[i]) || 0;
          deltaRow.push(currentValue - baselineValue);
        }
      } else {
        // New partner
        deltaRow.push(parseFloat(currentTotalProfiles) || 0); // Total Profiles
        for (let i = 3; i < currentRow.length; i++) {
          deltaRow.push(parseFloat(currentRow[i]) || 0);
        }
      }
      deltaData.push(deltaRow);
    });

    // 5. Write the results to a new sheet
    let outputSheet = ss.getSheetByName(outputSheetName);
    if (outputSheet) {
      outputSheet.clear();
    } else {
      outputSheet = ss.insertSheet(outputSheetName);
    }

    // Write headers
    if (headers.length > 0 && headers[0].length > 0) {
      outputSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
    }

    // Write data if there is any
    if (deltaData.length > 0) {
      outputSheet.getRange(headers.length + 1, 1, deltaData.length, deltaData[0].length).setValues(deltaData);
    }

    ss.toast('Performance delta calculation complete!', 'Success', 10);
    Logger.log('Performance delta calculation complete.');

  } catch (e) {
    ss.toast(`Error: ${e.toString()}`, 'Failed', 20);
    Logger.log(`Error in calculatePerformanceDelta: ${e.toString()}`);
  }
}
