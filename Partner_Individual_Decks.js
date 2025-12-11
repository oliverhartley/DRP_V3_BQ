/**
 * ****************************************
 * Google Apps Script - Individual Partner Decks
 * File: Partner_Individual_Decks.gs
 * Version: 11.8 (Corrected Indices V2)
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

const DECK_SHEET_NAME = "Tier Dashboard";
const SOURCE_DEEPDIVE_SHEET = "TEST_DeepDive_Data"; 

// DB Column Indices
const COL_INDEX_MANAGED = 5; 
const COL_INDEX_GSI = 6;
const COL_INDEX_BRAZIL = 7;
const COL_INDEX_MCO = 8;
const COL_INDEX_MEXICO = 9;
const COL_INDEX_PS = 10;
const COL_INDEX_DECK_STATUS = 19; // Column T
const MAX_EXECUTION_TIME_MS = 1500000; // 25 minutes

const PRODUCT_SCHEMA = [
  { solution: 'Infrastructure Modernization', color: '#fce5cd', products: ['Google Compute Engine', 'Google Cloud Networking', 'SAP on Google Cloud', 'Google Cloud VMware Engine', 'Google Distributed Cloud'] },
  { solution: 'Application Modernization', color: '#d9d2e9', products: ['Google Kubernetes Engine', 'Apigee API Management'] },
  { solution: 'Databases', color: '#fce5cd', products: ['Cloud SQL', 'AlloyDB for PostgreSQL', 'Spanner', 'Cloud Run', 'Oracle'] },
  { solution: 'Data & Analytics', color: '#d9ead3', products: ['BigQuery', 'Looker', 'Dataflow', 'Dataproc'] },
  { solution: 'Artificial Intelligence', color: '#c9daf8', products: ['Vertex AI Platform', 'AI Applications', 'Gemini Enterprise', 'Customer Engagement Suite'] },
  { solution: 'Security', color: '#f4cccc', products: ['Cloud Security', 'Security Command Center', 'Security Operations', 'Google Threat Intelligence'] },
  { solution: 'Workspace', color: '#fff2cc', products: ['Workspace'] }
];

// --- BATCH RUNNERS ---
function runManagedBatch() { runBatchByColumnIndex(COL_INDEX_MANAGED, "MANAGED PARTNERS", true); }
function runUnManagedBatch() { runBatchByColumnIndex(COL_INDEX_MANAGED, "UNMANAGED PARTNERS", false); }
function runGSIBatch() { runBatchByColumnIndex(COL_INDEX_GSI, "GSI PARTNERS", true); }
function runBrazilBatch() { runBatchByColumnIndex(COL_INDEX_BRAZIL, "BRAZIL PARTNERS", true); }
function runMCOBatch() { runBatchByColumnIndex(COL_INDEX_MCO, "MCO PARTNERS", true); }
function runMexicoBatch() { runBatchByColumnIndex(COL_INDEX_MEXICO, "MEXICO PARTNERS", true); }
function runPSBatch() { runBatchByColumnIndex(COL_INDEX_PS, "PS PARTNERS", true); }

// --- CONTROLLER ---
function runBatchByColumnIndex(colIndex, batchName, targetValue = true) {
  const startTime = new Date().getTime();
  const currentBatchId = getBatchId();
  Logger.log(`>>> STARTING BATCH: ${batchName} [Batch ID: ${currentBatchId}] <<<`);
  
  const partnerList = getPartnersByFlag(colIndex, targetValue);
  if (partnerList.length === 0) return;

  Logger.log(`Found ${partnerList.length} partners.`);
  const newLinks = [];

  // Need to open sheet to write status
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const dbSheet = ss.getSheetByName(SHEET_NAME_DB); // Ensure this sheet exists and matches Config

  for (let i = 0; i < partnerList.length; i++) {
    if (isTimeLimitApproaching(startTime)) {
      Logger.log("WARNING: Time limit approaching. Stopping to allow safe resume on next trigger.");
      break;
    }

    const pData = partnerList[i]; // Now an object { name, rowIndex, status }
    const pName = pData.name;
    const pRowIndex = pData.rowIndex;
    const pStatus = pData.status;

    if (pStatus === currentBatchId) {
      Logger.log(`[${i + 1}/${partnerList.length}] Skipping ${pName} (Already processed).`);
      continue;
    }

    Logger.log(`[${i + 1}/${partnerList.length}] Processing: ${pName}...`);
    try {
      const result = generateDeckForPartner(pName);
      if (result && result.url) {
        newLinks.push({ name: pName, url: result.url });
        // Update Status
        dbSheet.getRange(pRowIndex, COL_INDEX_DECK_STATUS + 1).setValue(currentBatchId);
      }
      Utilities.sleep(1000); 
    } catch (e) {
      Logger.log(`ERROR processing ${pName}: ${e.toString()}`);
    }
  }
  
  // *** AUTO-SAVE TO CACHE ***
  saveBatchLinks(newLinks); 
  Logger.log(`>>> BATCH COMPLETE: ${batchName} <<<`);
}

function getBatchId() {
  const now = new Date();
  const shiftedDate = new Date(now.getTime() - 24 * 60 * 60 * 1000); // Shift for Tue start
  const year = shiftedDate.getFullYear();
  const onejan = new Date(year, 0, 1);
  const week = Math.ceil((((shiftedDate.getTime() - onejan.getTime()) / 86400000) + onejan.getDay() + 1) / 7);
  return `DECK_${year}_${week}`;
}

function isTimeLimitApproaching(startTime) {
  return (new Date().getTime() - startTime) > MAX_EXECUTION_TIME_MS;
}

// --- GENERATOR ---
function generateDeckForPartner(targetPartner) {
  const scoreData = getPartnerDataFromMaster(targetPartner);
  if (!scoreData) { Logger.log(`  -> [SKIPPED]`); return null; }

  const deepDiveData = getDeepDiveData(targetPartner);
  const dashboardData = transposeDataForDeck(scoreData);
  const pivotedProfileData = pivotDeepDiveData(deepDiveData); 

  const result = updatePartnerSpreadsheet(targetPartner, dashboardData, scoreData.totalProfiles, pivotedProfileData);
  Logger.log(`  -> [${result.status}] Success: ${result.url}`);
  
  return result;
}

// --- HELPERS (Data & Transpose) ---
function getPartnersByFlag(colIndex, targetValue) {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_DB); 
  const data = sheet.getDataRange().getValues();
  const matches = [];
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const cellValue = data[i][colIndex];
    // Robust match for both boolean and string "TRUE"/"FALSE"
    const isMatch = (String(cellValue).toUpperCase() === String(targetValue).toUpperCase());

    if (isMatch) {
      const partnerName = data[i][1]; // Column B
      const status = data[i][COL_INDEX_DECK_STATUS];
      Logger.log(`Match found: ${partnerName} (Col ${colIndex} value: ${cellValue})`);
      matches.push({ name: partnerName, rowIndex: i + 1, status: status });
    }
  }
  return matches;
}

function getPartnerDataFromMaster(partnerName) {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_SCORE);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 3, lastCol).getValues();
  const partnerNames = sheet.getRange(4, 2, lastRow - 3, 1).getValues().flat();
  const partnerIndex = partnerNames.indexOf(partnerName);
  if (partnerIndex === -1) return null;
  const partnerRow = sheet.getRange(partnerIndex + 4, 1, 1, lastCol).getValues()[0];
  return { solutions: headers[0], products: headers[1], tiers: headers[2], data: partnerRow, totalProfiles: partnerRow[2] };
}

function generatePartnerDeck(partnerName) {
  // Version 11.2
  const ss = SpreadsheetApp.create(`${partnerName} - Partner Dashboard`);
  const sheet = ss.getSheetByName(SOURCE_DEEPDIVE_SHEET);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
}

function getDeepDiveData(partnerName) { // Version 11.2
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SOURCE_DEEPDIVE_SHEET);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  return data.filter(row => String(row[0]).trim().toLowerCase() === partnerName.toLowerCase());
}

function pivotDeepDiveData(rawRows) {
  if (rawRows.length === 0) return [];
  const profileMap = new Map();
  rawRows.forEach(row => {
    const pId = row[1]; const country = row[2]; const role = row[3]; const product = row[4]; const tier = row[6]; 
    if (!profileMap.has(pId)) { profileMap.set(pId, { info: [pId, country, role], scores: {} }); }
    profileMap.get(pId).scores[product] = tier;
  });
  const matrixRows = [];
  profileMap.forEach((value, key) => {
    const row = [...value.info]; 
    PRODUCT_SCHEMA.forEach(group => { group.products.forEach(prodName => { row.push(value.scores[prodName] || "-"); }); });
    matrixRows.push(row);
  });
  matrixRows.sort((a, b) => { if (a[1] < b[1]) return -1; if (a[1] > b[1]) return 1; return 0; });
  return matrixRows;
}

function transposeDataForDeck(source) {
  let output = [];
  output.push(["Solutions", "Products", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]);
  const colStart = 3; const totalCols = source.solutions.length;
  let currentSolution = "";
  for (let c = colStart; c < totalCols; c += 4) {
    let sol = source.solutions[c]; if (sol !== "") currentSolution = sol;
    let prod = source.products[c];
    if (prod && prod !== "") { output.push([currentSolution, prod, source.data[c], source.data[c+1], source.data[c+2], source.data[c+3]]); }
  }
  return output;
}

function updatePartnerSpreadsheet(partnerName, dashData, totalProfilesFromScoreData, pivotData) {
  const fileName = `${partnerName} - Partner Dashboard`;
  const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID); 
  let file, ss, actionStatus;
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) { file = files.next(); ss = SpreadsheetApp.open(file); actionStatus = "UPDATED"; } 
  else { ss = SpreadsheetApp.create(fileName); file = DriveApp.getFileById(ss.getId()); file.moveTo(folder); actionStatus = "CREATED"; }
  
  let sheet = ss.getSheetByName(DECK_SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(DECK_SHEET_NAME); }

  // --- PRESERVE COLUMN G (Es Producto Foco) ---
  let existingFocusData = [];
  const focusColIndex = 7; // Column G
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const headerVal = sheet.getRange(1, focusColIndex).getValue();
      // Check if header indicates focus column or if data exists (Column G is index 7)
      if (headerVal === "Es Producto Foco" || lastRow > 1) {
        existingFocusData = sheet.getRange(2, focusColIndex, lastRow - 1, 1).getValues();
      }
    }
  } catch (e) { Logger.log("Error reading existing focus data: " + e.toString()); }

  sheet.clear();

  let diveSheet = ss.getSheetByName("Profile Deep Dive");
  if (!diveSheet) { diveSheet = ss.insertSheet("Profile Deep Dive"); }
  else { diveSheet.clear(); }
  const actualDiveSheetName = diveSheet.getName();

  if (dashData.length > 0) {
    sheet.getRange(1, 1, dashData.length, dashData[0].length).setValues(dashData);

    // --- RESTORE / INIT COLUMN G ---
    sheet.getRange(1, focusColIndex).setValue("Es Producto Foco");
    const dataRows = dashData.length - 1;
    if (dataRows > 0) {
      const focusRange = sheet.getRange(2, focusColIndex, dataRows, 1);
      focusRange.insertCheckboxes();

      if (existingFocusData.length > 0) {
        // Restore values, handling potential length mismatch
        const restoredValues = [];
        for (let i = 0; i < dataRows; i++) {
          const val = (i < existingFocusData.length) ? existingFocusData[i][0] : false;
          restoredValues.push([val]);
        }
        focusRange.setValues(restoredValues);
      }
    }

    const totalProfiles = pivotData.length;
    const profilesWithTier = totalProfiles > 0 ? Number(totalProfilesFromScoreData) : 0; // Rename arg for clarity
    const profilesWithoutTier = Math.max(0, totalProfiles - profilesWithTier);

    sheet.getRange("I1").setValue("Profiles with Tier");
    sheet.getRange("J1").setValue("Profiles with no Tiers");
    sheet.getRange("K1").setValue("Total Profiles");
    sheet.getRange("I2").setValue(profilesWithTier);
    sheet.getRange("J2").setValue(profilesWithoutTier);
    sheet.getRange("K2").setValue(totalProfiles);

    // --- NEW: Region Slicer ---
    sheet.getRange("M1").setValue("Select Region / Country");
    sheet.getRange("M2").setValue("All"); // Default to All
    sheet.getRange("M1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
    sheet.getRange("M2").setBackground("#fff2cc").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

    // Add dropdown for Region/Country
    const countries = [...new Set(pivotData.map(row => row[1]))].sort(); // Column 1 is Country
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(["All", ...countries]).build();
    sheet.getRange("M2").setDataValidation(rule);

    // Ensure diveSheet has enough columns before creating formulas that reference them
    if (diveSheet.getMaxColumns() < pivotData[0].length) {
      diveSheet.insertColumnsAfter(diveSheet.getMaxColumns(), pivotData[0].length - diveSheet.getMaxColumns());
    }

    // Add formula for count in N1/N2
    sheet.getRange("N1").setValue("Profiles in Selection");
    sheet.getRange("N1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
    sheet.getRange("N2").setFormula(`=IF(M2="All", ${totalProfiles}, SUMPRODUCT((TRIM('${actualDiveSheetName}'!$B$1000:$B)=M2)*1))`);
    sheet.getRange("N2").setBackground("white").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

    formatDeckSheet(sheet, dashData.length, dashData[0].length, actualDiveSheetName);
  }

  if (diveSheet.getFilter()) { diveSheet.getFilter().remove(); }
  if (pivotData.length > 0) {
    // Add hyperlinks and count Tier 1s
    for (let i = 0; i < pivotData.length; i++) {
      const row = pivotData[i];
      let tier1Count = 0;
      for (let j = 3; j < row.length; j++) { // Products start at index 3
        if (row[j] === "Tier 1") {
          tier1Count++;
        }
      }
      // Insert count at index 3
      row.splice(3, 0, tier1Count);

      const profileId = row[0];
      if (profileId && typeof profileId === 'string' && !profileId.startsWith('=HYPERLINK')) {
        row[0] = `=HYPERLINK("https://delivery-readiness-portal.cloud.google/app/profiles/detailed-profile-view/${profileId}", "${profileId}")`;
      }
    }

    // Setup Profile Deep Dive sheet with Selector
    diveSheet.clear();
    if (diveSheet.getFilter()) { diveSheet.getFilter().remove(); }
    diveSheet.getSlicers().forEach(s => s.remove());

    // Write raw data far down (Row 1000) to keep it hidden but on same sheet
    const rawDataStartRow = 1000;
    diveSheet.getRange(rawDataStartRow, 1, pivotData.length, pivotData[0].length).setValues(pivotData);

    // Selector UI
    diveSheet.getRange("A1:D4").setBackground("#f3f3f3").setBorder(true, true, true, true, true, true);
    diveSheet.getRange("A1").setValue("Partner & Solution Selector").setFontWeight("bold").setFontSize(12);
    diveSheet.getRange("A2").setValue("Select Country:");
    diveSheet.getRange("A3").setValue("Select Product:");

    // Country Dropdown
    const countries = [...new Set(pivotData.map(r => r[1]).filter(c => c && c !== "Country"))].sort();
    countries.unshift("All");
    const countryRule = SpreadsheetApp.newDataValidation().requireValueInList(countries).setAllowInvalid(false).build();
    diveSheet.getRange("B2").setDataValidation(countryRule).setValue("All");

    // Product Dropdown (Placeholder for now, but can be added later)
    diveSheet.getRange("B3").setValue("All (Column Filters)");

    // Filter Formula
    const lastColLetter = columnToLetter(pivotData[0].length);
    // Note: We include Row 1 for headers, but we will handle headers separately for better formatting.
    // Actually, it's better to have headers static and filter only data.

    // Static Headers for Deep Dive
    formatDeepDivePivot(diveSheet, pivotData.length + 2, pivotData[0].length, rawDataStartRow);

  } else { diveSheet.getRange(1,1).setValue("No profile details available."); }

  const defaultSheet = ss.getSheetByName("Sheet1"); if (defaultSheet) ss.deleteSheet(defaultSheet);

  // --- NEW: Auto-Insert Images ---
  ensurePartnerImages(sheet);

  return { url: ss.getUrl(), status: actionStatus };
}

function formatDeckSheet(sheet, lastRow, lastCol, diveSheetName) {
  // Version 11.2
  try {
    const colorMap = { 'Infrastructure Modernization': '#fce5cd', 'Application Modernization': '#fff2cc', 'Databases': '#d9ead3', 'Data & Analytics': '#d0e0e3', 'Artificial Intelligence': '#c9daf8', 'Security': '#cfe2f3', 'Workspace': '#d9d2e9' };
    sheet.getRange(1, 1, 1, lastCol).setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    const fullTable = sheet.getRange(1, 1, lastRow, lastCol);
    fullTable.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
    const solutionCol = sheet.getRange(2, 1, lastRow - 1, 1);
    solutionCol.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setHorizontalAlignment("center").setTextRotation(90).setFontWeight("bold");
    sheet.getRange(2, 3, lastRow - 1, 4).setHorizontalAlignment("center"); 

    // NEW: Apply dynamic formulas to the table
    // Column A: Solutions, Column B: Products, Col C: Tier 1, Col D: Tier 2, Col E: Tier 3, Col F: Tier 4
    // Profile Deep Dive: Col B is Country, Col E onwards are products (due to new Tier 1 Count column)

    // --- FORMAT COLUMN G (Es Producto Foco) ---
    sheet.getRange(1, 7).setBackground("#e69138").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setColumnWidth(7, 100);
    sheet.getRange(2, 7, lastRow - 1, 1).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");

    let currentProductColIndex = 5; // Start at Column E (Google Compute Engine) in Profile Deep Dive
    for (let i = 2; i <= lastRow; i++) {
      const product = sheet.getRange(i, 2).getValue();
      if (product) {
        const colLetter = columnToLetter(currentProductColIndex);
        const rangeB = `'${diveSheetName}'!$B$1000:$B`;
        const rangeCol = `'${diveSheetName}'!$${colLetter}$1000:$${colLetter}`;

        sheet.getRange(i, 3).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 1"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 1")))`);
        sheet.getRange(i, 4).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 2"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 2")))`);
        sheet.getRange(i, 5).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 3"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 3")))`);
        sheet.getRange(i, 6).setFormula(`=IF($M$2="All", COUNTIFS(${rangeCol}, "Tier 4"), SUMPRODUCT((TRIM(${rangeB})=$M$2)*(${rangeCol}="Tier 4")))`);
        currentProductColIndex++;
      }
    }
    // Apply custom number format to show hyphen for zero
    sheet.getRange(2, 3, lastRow - 1, 4).setNumberFormat('0;-0;"-"');

    const headerRange = sheet.getRange("I1:K1");
    headerRange.setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    const valueRange = sheet.getRange("I2:K2");
    valueRange.setBackground("white").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
    const values = solutionCol.getValues();
    let mergeStartRow = 2; let currentVal = values[0][0];
    const applyBlockFormat = (startRow, endRow, val) => {
      const span = endRow - startRow;
      sheet.getRange(startRow, 1, span, 7).setBackground(colorMap[val] || '#ffffff');
        if (span > 0) { sheet.getRange(startRow, 1, span, 1).merge(); sheet.setRowHeights(startRow, span === 1 ? 1 : span, span === 1 ? 90 : 35); }
    };
    for (let i = 1; i < values.length; i++) { if (values[i][0] !== currentVal) { applyBlockFormat(mergeStartRow, i+2, currentVal); mergeStartRow = i+2; currentVal = values[i][0]; } }
    applyBlockFormat(mergeStartRow, lastRow + 1, currentVal);
    sheet.setColumnWidth(1, 40); // Adjusted width for vertical text
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidths(3, 4, 60);
    sheet.setColumnWidth(9, 100); sheet.setColumnWidth(10, 100); sheet.setColumnWidth(11, 100);
    sheet.setColumnWidth(13, 150); sheet.setColumnWidth(14, 150);
  } catch (e) {}
}

function formatDeepDivePivot(sheet, lastRow, lastCol, rawDataStartRow) {
  // Version 11.2
  try {
    const startRow = 6;
    sheet.setFrozenRows(0); sheet.setFrozenColumns(0);
    if (sheet.getMaxColumns() < lastCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
    }
    const fixedHeaders = ["Profile ID", "Country", "Job Title", "Tier 1 Count"];
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

    // Apply FILTER formula
    // Since we want to keep formatting, we might need to apply the formula to data rows only, but FILTER returns multiple rows.
    // Better to just use the formula in A7 (first data row) and keep A6 as static headers.
    const lastColLetter = columnToLetter(lastCol);
    sheet.getRange(startRow + 1, 1).setFormula(`=IFERROR(FILTER(A${rawDataStartRow}:${lastColLetter}${rawDataStartRow + 1000}, (B${rawDataStartRow}:B${rawDataStartRow + 1000} = B2) + (B2="All")), "No data found")`);

    // Formatting
    const dataRange = sheet.getRange(startRow + 1, 1, 500, lastCol); // Reduced to 500 for stability
    dataRange.setHorizontalAlignment("center");
    sheet.getRange(startRow + 1, 1, 500, 1).setFontColor("#1155cc").setFontLine("underline"); // Restore hyperlink formatting
    // Note: Conditional formatting might not work perfectly with dynamic filter, but let's try.
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

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function ensurePartnerImages(sheet) {
  const images = [
    { id: "1RrY--a7cZ9gYZKFZJa0v4ZIAT75aM0VH", row: 5, col: 9 },
    { id: "1Gf9sghdhjs-tnszdSP00IXlWR52UBaQs", row: 15, col: 9 }
  ];

  try {
    // Check existing images to avoid duplicates
    const existingImages = sheet.getImages();
    const occupiedCells = new Set();
    existingImages.forEach(img => {
      try {
        const anchor = img.getAnchorCell();
        occupiedCells.add(`${anchor.getRow()}_${anchor.getColumn()}`);
      } catch (e) { }
    });

    const token = ScriptApp.getOAuthToken();

    images.forEach(img => {
      if (occupiedCells.has(`${img.row}_${img.col}`)) return;

      try {
        // Fetch resized thumbnail to avoid "Blob too large" error
        const resizeUrl = `https://drive.google.com/thumbnail?id=${img.id}&sz=w1000`;
        const response = UrlFetchApp.fetch(resizeUrl, {
          headers: { 'Authorization': 'Bearer ' + token },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          sheet.insertImage(response.getBlob(), img.col, img.row);
        } else {
          // Fallback to direct Drive fetch
          try {
            sheet.insertImage(DriveApp.getFileById(img.id).getBlob(), img.col, img.row);
          } catch (e) {
            Logger.log(`Failed to insert image ${img.id}: ${e.toString()}`);
          }
        }
      } catch (e) {
        Logger.log(`Error processing image ${img.id}: ${e.toString()}`);
      }
    });
  } catch (e) {
    Logger.log(`Critical Error in ensurePartnerImages: ${e.toString()}`);
  }
}