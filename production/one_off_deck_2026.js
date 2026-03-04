/**
 * ****************************************
 * Google Apps Script - 2026 Deck Generator PROTOTYPE
 * File: one_off_deck_2026.js
 * Description: Safely tests generating a single consolidated deck for Accenture 
 *              using the new 2026 data structure.
 * ****************************************
 */

const TEST_PARTNER_NAME = "Accenture";
const TEST_DECK_NAME = "Accenture - TEST 2026 Deck";

function runTestDeck2026() {
  const ssMain = SpreadsheetApp.openById(DESTINATION_SS_ID);
  
  // 1. Fetch Scoring Data
  const scoreSheet = ssMain.getSheetByName(SHEET_NAME_SCORE_2026);
  if (!scoreSheet) throw new Error("Could not find " + SHEET_NAME_SCORE_2026);
  
  const scoreDataRange = scoreSheet.getDataRange();
  const scoreValues = scoreDataRange.getValues();
  
  // Find all rows for Accenture
  const partnerRows = [];
  const headersSol = scoreValues[0];
  const headersProd = scoreValues[1];
  const headersTier = scoreValues[2];
  
  let totalProfilesAcrossRegions = 0;
  
  for (let r = 3; r < scoreValues.length; r++) {
    const rowPartnerName = String(scoreValues[r][2]).trim(); // Column C
    if (rowPartnerName.toLowerCase() === TEST_PARTNER_NAME.toLowerCase()) {
      partnerRows.push({
        subRegion: String(scoreValues[r][3]).trim(), // Column D
        profiles: Number(scoreValues[r][6]) || 0,     // Column G
        data: scoreValues[r]
      });
      totalProfilesAcrossRegions += (Number(scoreValues[r][6]) || 0);
    }
  }
  
  if (partnerRows.length === 0) {
    Logger.log("No data found for " + TEST_PARTNER_NAME);
    return;
  }
  
  Logger.log(`Found ${partnerRows.length} region rows for ${TEST_PARTNER_NAME}. Total Profiles: ${totalProfilesAcrossRegions}`);
  
  // Aggregate scores across all regions
  const aggregatedRow = new Array(scoreValues[0].length).fill(0);
  
  // We only sum up the numeric tier columns (index 7 onwards)
  for (let r = 0; r < partnerRows.length; r++) {
    const rowData = partnerRows[r].data;
    for (let c = 7; c < rowData.length; c++) {
      aggregatedRow[c] += (Number(rowData[c]) || 0);
    }
  }
  
  // Format the consolidated dashboard view
  const dashboardData = [];
  dashboardData.push(["Solutions", "Products", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]);
  
  let currentSolution = "";
  for (let c = 7; c < headersSol.length; c += 4) {
    let sol = String(headersSol[c]).trim();
    if (sol !== "") currentSolution = sol;
    else {
      // Look back for merged header
      for (let k = c - 1; k >= 7; k--) {
        if (String(headersSol[k]).trim() !== "") { currentSolution = String(headersSol[k]).trim(); break; }
      }
    }
    
    let prod = String(headersProd[c]).trim();
    if (prod === "") {
        for (let k = c - 1; k >= 7; k--) {
            if (String(headersProd[k]).trim() !== "") { prod = String(headersProd[k]).trim(); break; }
        }
    }
    
    if (prod && prod !== "") {
      dashboardData.push([
        currentSolution, 
        prod, 
        aggregatedRow[c],     // Tier 1
        aggregatedRow[c+1],   // Tier 2
        aggregatedRow[c+2],   // Tier 3
        aggregatedRow[c+3]    // Tier 4
      ]);
    }
  }
  
  // 2. Fetch Deep Dive Data
  const deepDiveSheet = ssMain.getSheetByName(SHEET_NAME_DEEPDIVE_2026);
  let deepDiveData = [];
  if (deepDiveSheet) {
    const rawDive = deepDiveSheet.getDataRange().getValues();
    // Start from row 2 (skip header)
    for (let r = 1; r < rawDive.length; r++) {
       const divePartnerName = String(rawDive[r][2]).trim(); // Col C
       if (divePartnerName.toLowerCase() === TEST_PARTNER_NAME.toLowerCase()) {
         deepDiveData.push(rawDive[r]);
       }
    }
  }
  
  Logger.log(`Found ${deepDiveData.length} Deep Dive profile records.`);
  
  // Pivot the deep dive data
  // 2026 Deep Dive mapping:
  // [0] Internal ID, [1] Partner Name (D/E?), [2] Partner Name, [3] Sub Region, [4] PDM, [5] Type
  // [6] Timestamp, [7] Profile ID, [8] Country, [9] Job Title, [10] Solution, [11] Product, [12] Score, [13] Tier
  const pivotMap = new Map();
  deepDiveData.forEach(row => {
     const profileId = String(row[7]);
     const country = String(row[8]);
     const jobTitle = String(row[9]);
     const product = String(row[11]);
     const tier = String(row[13]);
     const subRegion = String(row[3]); // WE NEED THIS FOR THE SLICER NOW
     
     if (!pivotMap.has(profileId)) {
        // We include Sub Region here so it can be filtered
        pivotMap.set(profileId, { info: [profileId, subRegion, jobTitle], scores: {} });
     }
     pivotMap.get(profileId).scores[product] = tier;
  });
  
  const pivotedRows = [];
  pivotMap.forEach((value, key) => {
     const row = [...value.info];
     PRODUCT_SCHEMA.forEach(group => { 
        group.products.forEach(prodName => { 
            row.push(value.scores[prodName] || "-"); 
        }); 
     });
     pivotedRows.push(row);
  });
  
  // 3. Create the New Spreadsheet
  let testSS;
  const files = DriveApp.getFilesByName(TEST_DECK_NAME);
  if (files.hasNext()) {
      testSS = SpreadsheetApp.open(files.next());
  } else {
      testSS = SpreadsheetApp.create(TEST_DECK_NAME);
  }
  
  // Let's use the core logic from updatePartnerSpreadsheet but adapted for the test
  let sheet = testSS.getSheetByName(DECK_SHEET_NAME) || testSS.insertSheet(DECK_SHEET_NAME);
  sheet.clear();
  
  let diveOutSheet = testSS.getSheetByName("Profile Deep Dive") || testSS.insertSheet("Profile Deep Dive");
  diveOutSheet.clear();
  
  // WRITE DASHBOARD
  if (dashboardData.length > 0) {
      sheet.getRange(1, 1, dashboardData.length, dashboardData[0].length).setValues(dashboardData);
      
      const focusColIndex = 7;
      sheet.getRange(1, focusColIndex).setValue("Es Producto Foco");
      if (dashboardData.length > 1) {
          sheet.getRange(2, focusColIndex, dashboardData.length - 1, 1).insertCheckboxes();
      }
      
      const totalProfilesActual = pivotedRows.length;
      
      sheet.getRange("I1").setValue("Profiles with Tier");
      sheet.getRange("J1").setValue("Profiles with no Tiers");
      sheet.getRange("K1").setValue("Total Profiles");
      sheet.getRange("I2").setValue(totalProfilesAcrossRegions);
      sheet.getRange("J2").setValue(Math.max(0, totalProfilesActual - totalProfilesAcrossRegions));
      sheet.getRange("K2").setValue(totalProfilesActual);
      
      // NEW 2026 SLICER: Sub-Region
      sheet.getRange("M1").setValue("Select Sub-Region");
      sheet.getRange("M2").setValue("All"); 
      sheet.getRange("M1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      sheet.getRange("M2").setBackground("#fff2cc").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
      
      const subRegions = [...new Set(partnerRows.map(r => r.subRegion))].sort();
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(["All", ...subRegions]).build();
      sheet.getRange("M2").setDataValidation(rule);
      
      sheet.getRange("N1").setValue("Profiles in Selection");
      sheet.getRange("N1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      
      // Formula looking at column B (which is now Sub-Region in Deep Dive)
      sheet.getRange("N2").setFormula(`=IF(M2="All", ${totalProfilesActual}, SUMPRODUCT((TRIM('Profile Deep Dive'!$B$1000:$B)=M2)*1))`);
      sheet.getRange("N2").setBackground("white").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
      
      // Formatting
      formatDeckSheet(sheet, dashboardData.length, dashboardData[0].length, "Profile Deep Dive");
  }
  
  // WRITE DEEP DIVE
  if (pivotedRows.length > 0) {
      for (let i = 0; i < pivotedRows.length; i++) {
        const row = pivotedRows[i];
        let tier1Count = 0;
        for (let j = 3; j < row.length; j++) { if (row[j] === "Tier 1") tier1Count++; }
        row.splice(3, 0, tier1Count);
        
        const profileId = row[0];
        if (profileId && typeof profileId === 'string' && !profileId.startsWith('=HYPERLINK')) {
          row[0] = `=HYPERLINK("https://delivery-readiness-portal.cloud.google/app/profiles/detailed-profile-view/${profileId}", "${profileId}")`;
        }
      }
      
      const rawDataStartRow = 1000;
      diveOutSheet.getRange(rawDataStartRow, 1, pivotedRows.length, pivotedRows[0].length).setValues(pivotedRows);
      
      diveOutSheet.getRange("A1:D4").setBackground("#f3f3f3").setBorder(true, true, true, true, true, true);
      diveOutSheet.getRange("A1").setValue("Partner & Solution Selector").setFontWeight("bold").setFontSize(12);
      diveOutSheet.getRange("A2").setValue("Select Sub-Region:");
      diveOutSheet.getRange("A3").setValue("Select Product:");
      
      const regions = [...new Set(pivotedRows.map(r => r[1]))].sort();
      regions.unshift("All");
      const regionRule = SpreadsheetApp.newDataValidation().requireValueInList(regions).setAllowInvalid(false).build();
      diveOutSheet.getRange("B2").setDataValidation(regionRule).setValue("All");
      diveOutSheet.getRange("B3").setValue("All (Column Filters)");
      
      formatTestDeepDivePivot(diveOutSheet, pivotedRows.length + 2, pivotedRows[0].length, rawDataStartRow);
  }
  
  const defaultSheet = testSS.getSheetByName("Sheet1"); if (defaultSheet) testSS.deleteSheet(defaultSheet);
  
  Logger.log("DONE! Link: " + testSS.getUrl());
}

// Duplicated so we can safely edit it for the test
function formatTestDeepDivePivot(sheet, lastRow, lastCol, rawDataStartRow) {
  try {
    const startRow = 6;
    sheet.setFrozenRows(0); sheet.setFrozenColumns(0);
    if (sheet.getMaxColumns() < lastCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
    }
    
    // UPDATED FIXED HEADERS TO INCLUDE SUB-REGION INSTEAD OF COUNTRY
    const fixedHeaders = ["Profile ID", "Sub Region", "Job Title", "Tier 1 Count"];
    sheet.getRange(startRow, 1, 1, 4).setValues([fixedHeaders]);
    sheet.getRange(startRow - 1, 1, 1, 4).merge().setValue("Profile Details").setBackground("#666666").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(startRow, 1, 1, 4).setBackground("#d9d9d9").setFontWeight("bold");
    let currentCol = 5; 
    PRODUCT_SCHEMA.forEach(group => {
      const numProducts = group.products.length;
      if (numProducts > 0) {
        const solRange = sheet.getRange(startRow - 1, currentCol, 1, numProducts);
        solRange.merge().setValue(group.solution).setBackground(group.color).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        const prodRange = sheet.getRange(startRow, currentCol, 1, numProducts);
        prodRange.setValues([group.products]).setBackground(group.color).setFontWeight("bold").setHorizontalAlignment("center").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
        sheet.setColumnWidths(currentCol, numProducts, 100);
        currentCol += numProducts;
      }
    });

    const lastColLetter = columnToLetter(lastCol);
    sheet.getRange(startRow + 1, 1).setFormula(`=IFERROR(FILTER(A${rawDataStartRow}:${lastColLetter}${rawDataStartRow + 1000}, (B${rawDataStartRow}:B${rawDataStartRow + 1000} = B2) + (B2="All")), "No data found")`);

    const dataRange = sheet.getRange(startRow + 1, 1, 500, lastCol);
    dataRange.setHorizontalAlignment("center");
    sheet.getRange(startRow + 1, 1, 500, 1).setFontColor("#1155cc").setFontLine("underline");
    const scoreArea = sheet.getRange(startRow + 1, 5, 500, lastCol - 4);
    const rule1 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 1").setBackground("#d9ead3").setRanges([scoreArea]).build();
    const rule2 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 2").setBackground("#fff2cc").setRanges([scoreArea]).build();
    const rule3 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 3").setBackground("#fce5cd").setRanges([scoreArea]).build();
    const rule4 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 4").setBackground("#f4cccc").setRanges([scoreArea]).build();
    sheet.setConditionalFormatRules([rule1, rule2, rule3, rule4]);

    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(4); 

  } catch (e) { Logger.log("Matrix Formatting Error: " + e.toString()); }
}
