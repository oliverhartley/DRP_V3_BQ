/**
 * ****************************************
 * Google Apps Script - Individual Partner Decks
 * File: Partner_Individual_Decks.gs
 * Version: 8.0 (Auto-Save Links)
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

const DECK_SHEET_NAME = "Tier Dashboard";
const DEEPDIVE_SHEET_NAME = "Profile Deep Dive";
const SOURCE_DEEPDIVE_SHEET = "TEST_DeepDive_Data"; 

// DB Column Indices
const COL_INDEX_MANAGED = 5; 
const COL_INDEX_BRAZIL = 7; 
const COL_INDEX_MCO = 8;     
const COL_INDEX_MEXICO = 9;  

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
function runBrazilBatch() { runBatchByColumnIndex(COL_INDEX_BRAZIL, "Brazil", true); }
function runMCOBatch() { runBatchByColumnIndex(COL_INDEX_MCO, "MCO", true); }
function runMexicoBatch() { runBatchByColumnIndex(COL_INDEX_MEXICO, "Mexico", true); }

// --- CONTROLLER ---
function runBatchByColumnIndex(colIndex, batchName, targetValue = true) {
  Logger.log(`>>> STARTING BATCH: ${batchName} <<<`);
  
  const partnerList = getPartnersByFlag(colIndex, targetValue);
  if (partnerList.length === 0) return;

  Logger.log(`Found ${partnerList.length} partners.`);
  const newLinks = [];

  for (let i = 0; i < partnerList.length; i++) {
    const pName = partnerList[i];
    Logger.log(`[${i + 1}/${partnerList.length}] Processing: ${pName}...`);
    try {
      const result = generateDeckForPartner(pName);
      if (result && result.url) newLinks.push({ name: pName, url: result.url });
      Utilities.sleep(1000); 
    } catch (e) {
      Logger.log(`ERROR processing ${pName}: ${e.toString()}`);
    }
  }
  
  // *** AUTO-SAVE TO CACHE ***
  saveBatchLinks(newLinks); 
  Logger.log(`>>> BATCH COMPLETE: ${batchName} <<<`);
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] === targetValue) matches.push(data[i][1]);
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

function getDeepDiveData(partnerName) {
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

function updatePartnerSpreadsheet(partnerName, dashData, totalProfiles, pivotData) {
  const fileName = `${partnerName} - Partner Dashboard`;
  const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID); 
  let file, ss, actionStatus;
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) { file = files.next(); ss = SpreadsheetApp.open(file); actionStatus = "UPDATED"; } 
  else { ss = SpreadsheetApp.create(fileName); file = DriveApp.getFileById(ss.getId()); file.moveTo(folder); actionStatus = "CREATED"; }
  
  let sheet = ss.getSheetByName(DECK_SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(DECK_SHEET_NAME); }
  sheet.clear();
  if (dashData.length > 0) {
    sheet.getRange(1, 1, dashData.length, dashData[0].length).setValues(dashData);
    sheet.getRange("I1").setValue("Total Profiles"); sheet.getRange("I2").setValue(totalProfiles);
    formatDeckSheet(sheet, dashData.length, dashData[0].length);
  }

  let diveSheet = ss.getSheetByName(DEEPDIVE_SHEET_NAME);
  if (!diveSheet) { diveSheet = ss.insertSheet(DEEPDIVE_SHEET_NAME); }
  diveSheet.clear(); 
  if (diveSheet.getFilter()) { diveSheet.getFilter().remove(); }
  if (pivotData.length > 0) {
    diveSheet.getRange(3, 1, pivotData.length, pivotData[0].length).setValues(pivotData);
    formatDeepDivePivot(diveSheet, pivotData.length + 2, pivotData[0].length);
  } else { diveSheet.getRange(1,1).setValue("No profile details available."); }

  const defaultSheet = ss.getSheetByName("Sheet1"); if (defaultSheet) ss.deleteSheet(defaultSheet);
  return { url: ss.getUrl(), status: actionStatus };
}

function formatDeckSheet(sheet, lastRow, lastCol) {
  try {
    const colorMap = { 'Infrastructure Modernization': '#fce5cd', 'Application Modernization': '#fff2cc', 'Databases': '#d9ead3', 'Data & Analytics': '#d0e0e3', 'Artificial Intelligence': '#c9daf8', 'Security': '#cfe2f3', 'Workspace': '#d9d2e9' };
    sheet.getRange(1, 1, 1, lastCol).setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    const fullTable = sheet.getRange(1, 1, lastRow, lastCol);
    fullTable.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
    const solutionCol = sheet.getRange(2, 1, lastRow - 1, 1);
    solutionCol.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setHorizontalAlignment("center").setTextRotation(90).setFontWeight("bold");
    sheet.getRange(2, 3, lastRow - 1, 4).setHorizontalAlignment("center"); 
    sheet.getRange("I1:J1").merge().setBackground("#4285f4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true,true,true,true,true,true);
    sheet.getRange("I2:J2").merge().setBackground("white").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true,true,true,true,true,true);
    const values = solutionCol.getValues();
    let mergeStartRow = 2; let currentVal = values[0][0];
    const applyBlockFormat = (startRow, endRow, val) => {
        const span = endRow - startRow; sheet.getRange(startRow, 1, span, lastCol).setBackground(colorMap[val] || '#ffffff');
        if (span > 0) { sheet.getRange(startRow, 1, span, 1).merge(); sheet.setRowHeights(startRow, span === 1 ? 1 : span, span === 1 ? 90 : 35); }
    };
    for (let i = 1; i < values.length; i++) { if (values[i][0] !== currentVal) { applyBlockFormat(mergeStartRow, i+2, currentVal); mergeStartRow = i+2; currentVal = values[i][0]; } }
    applyBlockFormat(mergeStartRow, lastRow + 1, currentVal);
    sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 250); sheet.setColumnWidths(3, 4, 60);
  } catch (e) {}
}

function formatDeepDivePivot(sheet, lastRow, lastCol) {
  try {
    if (sheet.getFilter()) { sheet.getFilter().remove(); }
    const fixedHeaders = ["Profile ID", "Country", "Job Title"];
    sheet.getRange(2, 1, 1, 3).setValues([fixedHeaders]);
    sheet.getRange(1, 1, 1, 3).merge().setValue("Profile Details").setBackground("#666666").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(2, 1, 1, 3).setBackground("#d9d9d9").setFontWeight("bold");
    let currentCol = 4; 
    PRODUCT_SCHEMA.forEach(group => {
      const numProducts = group.products.length;
      if (numProducts > 0) {
        const solRange = sheet.getRange(1, currentCol, 1, numProducts);
        solRange.merge().setValue(group.solution).setBackground(group.color).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        const prodRange = sheet.getRange(2, currentCol, 1, numProducts);
        prodRange.setValues([group.products]).setBackground(group.color).setFontWeight("bold").setHorizontalAlignment("center").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
        sheet.setColumnWidths(currentCol, numProducts, 100);
        currentCol += numProducts;
      }
    });
    const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);
    dataRange.setHorizontalAlignment("center");
    dataRange.setBorder(true, true, true, true, true, true);
    const scoreArea = sheet.getRange(3, 4, lastRow - 2, lastCol - 3);
    const rule1 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 1").setBackground("#d9ead3").setRanges([scoreArea]).build();
    const rule2 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 2").setBackground("#fff2cc").setRanges([scoreArea]).build();
    const rule3 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 3").setBackground("#fce5cd").setRanges([scoreArea]).build();
    const rule4 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Tier 4").setBackground("#f4cccc").setRanges([scoreArea]).build();
    sheet.setConditionalFormatRules([rule1, rule2, rule3, rule4]);
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(3); 
    sheet.getRange(2, 1, lastRow - 1, lastCol).createFilter();
  } catch (e) { Logger.log("Matrix Formatting Error: " + e.toString()); }
}