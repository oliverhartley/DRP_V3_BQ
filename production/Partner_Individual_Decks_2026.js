/**
 * ****************************************
 * Google Apps Script - 2026 Batch Deck Generator
 * File: Partner_Individual_Decks_2026.js
 * Description: Generates individual partner decks based on the new 2026 data model.
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.js

function runAccentureTestBatch2026() {
  const testPartner = "Accenture";
  const currentBatchId = getDeckBatchId2026();
  Logger.log(`>>> STARTING TARTGETED TEST BATCH FOR: ${testPartner} [Batch ID: ${currentBatchId}] <<<`);

  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const dbSheet = ss.getSheetByName(SHEET_NAME_2026);

  if (!dbSheet) {
    Logger.log(`Error: Database sheet ${SHEET_NAME_2026} not found.`);
    return;
  }

  const dbData = dbSheet.getDataRange().getValues();
  const partnerRows = [];
  let existingDeckId = null;

  // Find Accenture rows in DB (Column A is index 0)
  for (let i = 1; i < dbData.length; i++) {
    if (String(dbData[i][0]).trim().toLowerCase() === testPartner.toLowerCase()) {
      partnerRows.push(i + 1); // 1-based row index
      // Col I is index 8
      if (!existingDeckId && dbData[i][8]) {
        existingDeckId = String(dbData[i][8]).trim();
      }
    }
  }

  if (partnerRows.length === 0) {
    Logger.log(`No entries found for ${testPartner} in Database.`);
    return;
  }

  Logger.log(`Found ${partnerRows.length} region entries for ${testPartner}. Existing Deck ID: ${existingDeckId || "None"}. Processing deck...`);

  const result = generateDeckForPartner2026(testPartner, existingDeckId);

  if (result && result.url) {
    Logger.log(`Deck Generated! Updating Database Links...`);
    // Write links back to all corresponding DB rows in Col I (9) and J (10)
    for (const r of partnerRows) {
      dbSheet.getRange(r, 9).setValue(result.id);
      dbSheet.getRange(r, 10).setValue(result.url);
    }
  }

  Logger.log(`>>> TARGETED TEST COMPLETE <<<`);
}

function getDeckBatchId2026() {
  const now = new Date();
  const shiftedDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const year = shiftedDate.getFullYear();
  const onejan = new Date(year, 0, 1);
  const week = Math.ceil((((shiftedDate.getTime() - onejan.getTime()) / 86400000) + onejan.getDay() + 1) / 7);
  return `UPDATED_2026_${year}_${week}`;
}

// Core generation logic derived from the prototype
function generateDeckForPartner2026(partnerName, existingDeckId = null) {
  const ssMain = SpreadsheetApp.openById(DESTINATION_SS_ID);

  // 1. Fetch Scoring Data
  const scoreSheet = ssMain.getSheetByName(SHEET_NAME_SCORE_2026);
  if (!scoreSheet) { Logger.log("Error: Score sheet not found."); return null; }

  const scoreValues = scoreSheet.getDataRange().getValues();
  const headersSol = scoreValues[0];
  const headersProd = scoreValues[1];

  const partnerScoreRows = [];
  let totalProfilesAcrossRegions = 0;

  // Find all rows for this partner in the SCORE sheet (Partner Name is Col C / Index 2)
  for (let r = 3; r < scoreValues.length; r++) {
    if (String(scoreValues[r][2]).trim().toLowerCase() === partnerName.toLowerCase()) {
      partnerScoreRows.push({
        subRegion: String(scoreValues[r][3]).trim(), // Column D
        profiles: Number(scoreValues[r][6]) || 0,     // Column G
        data: scoreValues[r]
      });
      totalProfilesAcrossRegions += (Number(scoreValues[r][6]) || 0);
    }
  }

  if (partnerScoreRows.length === 0) return null;

  // Aggregate scores across all regions
  const aggregatedRow = new Array(scoreValues[0].length).fill(0);
  for (let r = 0; r < partnerScoreRows.length; r++) {
    const rowData = partnerScoreRows[r].data;
    for (let c = 7; c < rowData.length; c++) {
      aggregatedRow[c] += (Number(rowData[c]) || 0);
    }
  }

  // Format dashboard
  const dashboardData = [["Solutions", "Products", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]];
  let currentSolution = "";
  for (let c = 7; c < headersSol.length; c += 4) {
    let sol = String(headersSol[c]).trim();
    if (sol !== "") currentSolution = sol;
    else {
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
        currentSolution, prod,
        aggregatedRow[c], aggregatedRow[c + 1], aggregatedRow[c + 2], aggregatedRow[c + 3]
      ]);
    }
  }

  // 2. Fetch Deep Dive Data
  const deepDiveSheet = ssMain.getSheetByName(SHEET_NAME_DEEPDIVE_2026);
  let deepDiveData = [];
  if (deepDiveSheet) {
    const rawDive = deepDiveSheet.getDataRange().getValues();
    // Col 2 is partner_name
    for (let r = 1; r < rawDive.length; r++) {
      if (String(rawDive[r][2]).trim().toLowerCase() === partnerName.toLowerCase()) {
        deepDiveData.push(rawDive[r]);
      }
    }
  }

  // Pivot Deep Dive
  const pivotMap = new Map();
  deepDiveData.forEach(row => {
    const profileId = String(row[6]);
    const jobTitle = String(row[8]);
    const product = String(row[9]);
    const tier = String(row[11]);
    const subRegion = String(row[3]);

    if (!pivotMap.has(profileId)) {
      pivotMap.set(profileId, { info: [profileId, subRegion, jobTitle], scores: {} });
    }
    pivotMap.get(profileId).scores[product] = tier;
  });

  const pivotedRows = [];
  pivotMap.forEach((value, key) => {
    const row = [...value.info];
    const userSolutions = new Set();
    PRODUCT_SCHEMA.forEach((group) => {
      row.push(""); // Spacer
      group.products.forEach(prodName => {
        const t = value.scores[prodName] || "-";
        row.push(t);
        if (t !== "-") userSolutions.add(group.solution);
      });
    });
    row.push(Array.from(userSolutions).join(","));
    pivotedRows.push(row);
  });

  // 3. Create/Update Spreadsheet
  const deckName = `${partnerName} - Partner Dashboard 2026`;
  let targetSS;
  const folderId = "1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC"; // Force using the designated 2025 folder
  const folder = DriveApp.getFolderById(folderId);

  if (existingDeckId) {
    try {
      targetSS = SpreadsheetApp.openById(existingDeckId);
      Logger.log(`Successfully opened existing deck: ${targetSS.getName()}`);
    } catch (e) {
      Logger.log(`WARNING: Could not open deck by ID ${existingDeckId}. Searching by name...`);
    }
  }

  if (!targetSS) {
    const oldDeckName = `${partnerName} - Partner Dashboard`; // Exactly match the 2025 naming convention
    const files = folder.getFilesByName(oldDeckName); 
    if (files.hasNext()) {
      targetSS = SpreadsheetApp.open(files.next());
      Logger.log(`Found existing deck by name: ${targetSS.getName()}`);
    } else {
      // Fallback search with generic suffix
      const suffixFiles = folder.getFilesByName(deckName);
      if (suffixFiles.hasNext()) {
        targetSS = SpreadsheetApp.open(suffixFiles.next());
      } else {
        targetSS = SpreadsheetApp.create(deckName);
        Logger.log(`Created NEW deck: ${deckName}`);
        const file = DriveApp.getFileById(targetSS.getId());
        file.moveTo(folder);
      }
    }
  }

  let sheet = targetSS.getSheetByName(DECK_SHEET_NAME) || targetSS.insertSheet(DECK_SHEET_NAME);
  sheet.clear();

  let diveOutSheet = targetSS.getSheetByName("Profile Deep Dive") || targetSS.insertSheet("Profile Deep Dive");
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

    sheet.getRange("M1").setValue("Select Sub-Region");
    sheet.getRange("M2").setValue("All");
    sheet.getRange("M1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
    sheet.getRange("M2").setBackground("#fff2cc").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

    const subRegions = [...new Set(partnerScoreRows.map(r => r.subRegion))].sort();
    if (subRegions.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(["All", ...subRegions]).build();
      sheet.getRange("M2").setDataValidation(rule);
    }

    sheet.getRange("N1").setValue("Profiles in Selection");
    sheet.getRange("N1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

    sheet.getRange("N2").setFormula(`=IF(M2="All", ${totalProfilesActual}, SUMPRODUCT((TRIM('Profile Deep Dive'!$B$1000:$B)=M2)*1))`);
    sheet.getRange("N2").setBackground("white").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

    formatDeckSheet2026(sheet, dashboardData.length, dashboardData[0].length, "Profile Deep Dive");
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

    diveOutSheet.getRange("A1:D4").clearFormat();
    diveOutSheet.getRange("A1:D1").merge().setValue(" Partner & Solution Selector").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setFontSize(14).setVerticalAlignment("middle");
    diveOutSheet.setRowHeight(1, 35);

    diveOutSheet.getRange("A2:A3").setBackground("#e8f0fe").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
    diveOutSheet.getRange("A2").setValue("Select Sub-Region:");
    diveOutSheet.getRange("A3").setValue("Select Product:");

    diveOutSheet.getRange("B2:B3").setBackground('white').setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
    diveOutSheet.setColumnWidth(1, 150);
    diveOutSheet.setColumnWidth(2, 250);

    const regions = [...new Set(pivotedRows.map(r => r[1]))].sort();
    regions.unshift("All");
    const regionRule = SpreadsheetApp.newDataValidation().requireValueInList(regions).setAllowInvalid(false).build();
    diveOutSheet.getRange("B2").setDataValidation(regionRule).setValue("All");

    const solutions = ["All", ...PRODUCT_SCHEMA.map(g => g.solution)];
    const solutionRule = SpreadsheetApp.newDataValidation().requireValueInList(solutions).setAllowInvalid(false).build();
    diveOutSheet.getRange("B3").setDataValidation(solutionRule).setValue("All");

    formatTestDeepDivePivot(diveOutSheet, pivotedRows.length + 2, pivotedRows[0].length, rawDataStartRow);
  }

  const defaultSheet = targetSS.getSheetByName("Sheet1"); if (defaultSheet) targetSS.deleteSheet(defaultSheet);

  ensurePartnerImages(sheet);

  return { id: targetSS.getId(), url: targetSS.getUrl() };
}

function formatDeckSheet2026(sheet, lastRow, lastCol, diveSheetName) {
  try {
    const colorMap = { 'Infrastructure Modernization': '#fce5cd', 'Application Modernization': '#fff2cc', 'Databases': '#d9ead3', 'Data & Analytics': '#d0e0e3', 'Artificial Intelligence': '#c9daf8', 'Security': '#cfe2f3', 'Workspace': '#d9d2e9' };
    sheet.getRange(1, 1, 1, lastCol).setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    const fullTable = sheet.getRange(1, 1, lastRow, lastCol);
    fullTable.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
    const solutionCol = sheet.getRange(2, 1, lastRow - 1, 1);
    solutionCol.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setHorizontalAlignment("center").setTextRotation(90).setFontWeight("bold");
    sheet.getRange(2, 3, lastRow - 1, 4).setHorizontalAlignment("center");

    // --- FORMAT COLUMN G (Es Producto Foco) ---
    sheet.getRange(1, 7).setBackground("#e69138").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setColumnWidth(7, 100);
    sheet.getRange(2, 7, lastRow - 1, 1).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");

    // --- MAP EXACT COLUMNS FOR FORMULAS ---
    const productColMap = {};
    let currentDeepDiveColIdx = 5; // A=1(Profile), B=2(SubRegion), C=3(JobTitle), D=4(Tier1 Count)

    // In pivotedRows, we push: Spacer -> Prod 1 -> Prod 2 ...
    PRODUCT_SCHEMA.forEach(group => {
      currentDeepDiveColIdx++; // Skip Spacer
      group.products.forEach(prod => {
        productColMap[prod.toLowerCase()] = columnToLetter(currentDeepDiveColIdx);
        currentDeepDiveColIdx++;
      });
    });

    for (let i = 2; i <= lastRow; i++) {
      const product = String(sheet.getRange(i, 2).getValue()).trim();
      const solution = String(sheet.getRange(i, 1).getValue()).trim();
      if (solution && colorMap[solution]) {
        sheet.getRange(i, 1, 1, lastCol).setBackground(colorMap[solution]);
      }

      if (product) {
        const colLetter = productColMap[product.toLowerCase()];
        if (colLetter) {
          const rangeB = `'${diveSheetName}'!$B$1000:$B`;
          const rangeCol = `'${diveSheetName}'!$${colLetter}$1000:$${colLetter}`;

          sheet.getRange(i, 3).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 1"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 1")))`);
          sheet.getRange(i, 4).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 2"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 2")))`);
          sheet.getRange(i, 5).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 3"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 3")))`);
          sheet.getRange(i, 6).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 4"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 4")))`);
        }
      }
    }
    // Apply custom number format to show hyphen for zero
    sheet.getRange(2, 3, lastRow - 1, 4).setNumberFormat('0;-0;"-"');

    const headerRange = sheet.getRange("I1:K1");
    headerRange.setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    const valueRange = sheet.getRange("I2:K2");
    valueRange.setBackground("white").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

    // Auto-resize product column
    sheet.setColumnWidth(2, 230);

    // --- MERGE SOLUTION CELLS VERTICALLY ---
    let startRow = 2;
    let currentVal = String(sheet.getRange(2, 1).getValue()).trim();
    for (let i = 3; i <= lastRow + 1; i++) {
      let val = (i <= lastRow) ? String(sheet.getRange(i, 1).getValue()).trim() : null;
      if (val !== currentVal) {
        if (i - startRow > 1) {
          sheet.getRange(startRow, 1, i - startRow, 1).mergeVertical().setVerticalAlignment("middle");
        }
        currentVal = val;
        startRow = i;
      }
    }

  } catch (e) {
    Logger.log("Error in formatDeckSheet2026: " + e.toString());
  }
}
