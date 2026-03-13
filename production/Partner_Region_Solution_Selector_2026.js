/**
 * ****************************************
 * Google Apps Script - Partner Dashboard Slicer (2026 Version)
 * File: Partner_Region_Solution_Selector_2026.js
 * Description: Google Sheets-based interactive dashboard slicing for 2026 data.
 * ****************************************
 */

const CELL_TYPE_2026 = {r: 3, c: 2};     
const CELL_SUB_REGION_2026 = {r: 4, c: 2};   
const CELL_PDM_2026 = {r: 5, c: 2};  
const CELL_SOLUTION_2026 = {r: 6, c: 2}; 
const CELL_PRODUCT_2026 = {r: 7, c: 2};  
const CELL_STATUS_2026 = {r: 3, c: 4};   
const DATA_START_ROW_2026 = 9;

function setLoadingStatus2026(sheet, isLoading) {
  const cell = sheet.getRange(CELL_STATUS_2026.r, CELL_STATUS_2026.c);
  if (isLoading) {
    cell.setValue("⏳ UPDATING...")
        .setBackground("#f4cccc")
        .setFontColor("#cc0000")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
  } else {
    cell.clearContent().setBackground(null);
  }
  SpreadsheetApp.flush();
}



/**
 * Creates the dashboard layout.
 */
function setupDashboard2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_DASHBOARD_2026);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_DASHBOARD_2026); }
  sheet.clear();
  setLoadingStatus2026(sheet, true);

  // Set Header Titles
  sheet.getRange("A1").setValue("DRP Dashboard 2026").setFontWeight("bold").setFontSize(18);
  sheet.getRange("A2").setValue("Status by partner").setFontWeight("bold").setFontSize(15);
  sheet.getRange("A3").setValue("Questions / Feedback: oliverhartley@").setFontWeight("bold").setFontSize(10);
  
  // Initialize Cache
  updateDashboardCache2026();

  try {
    refreshDashboardData2026(sheet);
  } catch (e) {
    sheet.getRange(DATA_START_ROW_2026, 1).setValue("Error loading initial data: " + e.toString());
  }
  
  setLoadingStatus2026(sheet, false);
}

/**
 * Updates the Dashboard Cache Sheet directly from LATAM_Partner_Score_2026.
 * Because all metadata is now inline, this just copies the values/formatting exactly,
 * replacing the complex merging from 2025.
 */
function updateDashboardCache2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoreSheet = ss.getSheetByName(SHEET_NAME_SCORE_2026);
  
  if (!scoreSheet) throw new Error("2026 Score Sheet missing.");
  ss.toast("Updating Dashboard Cache...", "Processing", 30);

  const scoreRange = scoreSheet.getDataRange();
  const scoreValues = scoreRange.getValues();
  const scoreBackgrounds = scoreRange.getBackgrounds();
  const scoreFontWeights = scoreRange.getFontWeights();
  const scoreFontColors = scoreRange.getFontColors();

  let cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE_2026);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(SHEET_NAME_CACHE_2026);
    cacheSheet.hideSheet();
  }
  cacheSheet.clear();

  if (scoreValues.length > 0) {
    const range = cacheSheet.getRange(1, 1, scoreValues.length, scoreValues[0].length);
    range.setValues(scoreValues);
    range.setBackgrounds(scoreBackgrounds);
    range.setFontWeights(scoreFontWeights);
    range.setFontColors(scoreFontColors);
  }

  ss.toast("Dashboard Cache Updated!", "Success", 5);
}


/**
 * Reads from the fast cache, applies the slicers, and drops the view.
 */
function refreshDashboardData2026(dashSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE_2026);

  if (!cacheSheet) {
    dashSheet.getRange(DATA_START_ROW_2026, 1).setValue("Error: Cache missing. Run Setup.");
    return;
  }


  // 2. Read Cache
  const cacheRange = cacheSheet.getDataRange();
  const cacheValues = cacheRange.getValues();
  const cacheBackgrounds = cacheRange.getBackgrounds();
  const cacheWeights = cacheRange.getFontWeights();
  const cacheFontColors = cacheRange.getFontColors(); 

  if (cacheValues.length < 3) return;

  const rowSol = cacheValues[0];
  const rowProd = cacheValues[1];

  const columnsToKeepInfo = [
    { type: 'meta', index: 2 },
    { type: 'meta', index: 3 },
    { type: 'meta', index: 4 },
    { type: 'meta', index: 5 },
    { type: 'meta', index: 6 }
  ];

  const effectiveHeaders = { sol: {}, prod: {} }; 
  let currentEffectiveSol = "";
  for (let c = 7; c < rowSol.length; c++) {
    let prod = String(rowProd[c]).trim();
    let sol = String(rowSol[c]).trim(); 
    
    let effectiveSol = sol;
    if (effectiveSol === "") { 
        for (let k = c - 1; k >= 7; k--) { 
            if (String(rowSol[k]).trim() !== "") { effectiveSol = String(rowSol[k]).trim(); break; } 
        } 
    }
    
    let effectiveProd = prod;
    if (effectiveProd === "") { 
        for (let k = c - 1; k >= 7; k--) { 
            if (String(rowProd[k]).trim() !== "") { effectiveProd = String(rowProd[k]).trim(); break; } 
        } 
    }
    
    effectiveHeaders.sol[c] = effectiveSol; 
    effectiveHeaders.prod[c] = effectiveProd;

    if (effectiveSol !== currentEffectiveSol) {
       const color = cacheBackgrounds[0][c] && cacheBackgrounds[0][c] !== "#ffffff" ? cacheBackgrounds[0][c] : "#f3f3f3";
       columnsToKeepInfo.push({ type: 'spacer', solution: effectiveSol, color: color });
       currentEffectiveSol = effectiveSol;
    }
    columnsToKeepInfo.push({ type: 'data', index: c, sol: effectiveSol, prod: effectiveProd });
  }

  // Read DB to get URLs for Hyperlinking
  const dbSheet = ss.getSheetByName(SHEET_NAME_2026);
  const dbData = dbSheet ? dbSheet.getDataRange().getValues() : [];
  const partnerUrlMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const pName = String(dbData[i][0]).trim();
    const url = String(dbData[i][9]).trim(); // Col J (Index 9)
    if (pName && url) partnerUrlMap.set(pName.toLowerCase(), url);
  }

  // 4. Build Output Data
  let outputValues = [], outputBackgrounds = [], outputWeights = [], outputFontColors = [];

  // Headers (2 rows: Row 0 is Product, Row 1 is Tier)
  for (let r = 1; r < 3; r++) {
    let rowV = [], rowB = [], rowW = [], rowFC = [];
    columnsToKeepInfo.forEach(info => {
      if (info.type === 'meta') {
        // Since we skip r=0, we must force the text for the top header row (r=1) to use cacheValues[0] where the text actually lives.
        let val = cacheValues[r][info.index];
        if (r === 1) val = cacheValues[0][info.index];
        if (r === 2) val = ""; // Clear text on row 10 so it's a blank cell for the filter
        rowV.push(val);
        rowB.push(cacheBackgrounds[r][info.index]);
        rowW.push(cacheWeights[r][info.index]);
        rowFC.push(cacheFontColors[r][info.index]);
      } else if (info.type === 'spacer') {
        rowV.push(""); // Spacer header is blank for the horizontal rows
        rowB.push(info.color);
        rowW.push("bold");
        rowFC.push("#000000"); // Standard text color for spacer
      } else if (info.type === 'data') {
        let val = cacheValues[r][info.index];
        if (r === 1) val = info.prod;
        rowV.push(val);
        rowB.push(cacheBackgrounds[r][info.index]);
        rowW.push(cacheWeights[r][info.index]);
        rowFC.push(cacheFontColors[r][info.index]);
      }
    });
    outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW); outputFontColors.push(rowFC);
  }

  // Data Rows
  for (let r = 3; r < cacheValues.length; r++) {
    let rowV = [], rowB = [], rowW = [], rowFC = [];
    columnsToKeepInfo.forEach(info => {
      if (info.type === 'meta' || info.type === 'data') {
        rowV.push(cacheValues[r][info.index]);
        rowB.push(cacheBackgrounds[r][info.index]);
        rowW.push(cacheWeights[r][info.index]);
        rowFC.push(cacheFontColors[r][info.index]);
      } else if (info.type === 'spacer') {
        rowV.push(""); // Spacer body is blank
        rowB.push(info.color);
        rowW.push("normal");
        rowFC.push("#000000");
      }
    });
    outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW); outputFontColors.push(rowFC);
  }

  // 5. Apply Output to Sheet
  const lastRow = dashSheet.getLastRow(); 
  const lastCol = dashSheet.getLastColumn();
  if (lastRow >= DATA_START_ROW_2026) {
      dashSheet.getRange(DATA_START_ROW_2026, 1, lastRow - DATA_START_ROW_2026 + 1, lastCol || 1).clear();
  }
  
  if (outputValues.length > 3) {
    const outRows = outputValues.length; const outCols = outputValues[0].length;
    const targetRange = dashSheet.getRange(DATA_START_ROW_2026, 1, outRows, outCols);

    targetRange.setValues(outputValues);
    targetRange.setBackgrounds(outputBackgrounds);
    targetRange.setFontWeights(outputWeights);
    targetRange.setFontColors(outputFontColors);

    // Apply Hyperlinks to Partner Names
    let richTexts = [];
    for (let i = 2; i < outputValues.length; i++) { // offset by 2 header rows instead of 3
      const pName = String(outputValues[i][0]);
      const url = partnerUrlMap.get(pName.trim().toLowerCase());
      if (url && url !== "") {
        richTexts.push([SpreadsheetApp.newRichTextValue().setText(pName).setLinkUrl(url).build()]);
      } else {
        richTexts.push([SpreadsheetApp.newRichTextValue().setText(pName).build()]);
      }
    }
    if (richTexts.length > 0) {
      dashSheet.getRange(DATA_START_ROW_2026 + 2, 1, richTexts.length, 1).setRichTextValues(richTexts); // offset by 2 header rows
    }

    targetRange.setHorizontalAlignment("center");
    dashSheet.getRange(DATA_START_ROW_2026, 1, outRows, 1).setHorizontalAlignment("left"); // Partner Name
    dashSheet.getRange(DATA_START_ROW_2026, 1, outRows, outCols).setBorder(true, true, true, true, true, true);
    dashSheet.getRange(DATA_START_ROW_2026, 1, 2, outCols).setBorder(true, true, true, true, true, true); // borders for 2 header rows
    
    // Auto Resize width for the first 5 metadata columns
    // Clear all existing groups first
    const maxC = dashSheet.getMaxColumns();
    for (let c = 1; c <= maxC; c++) {
      let depth = dashSheet.getColumnGroupDepth(c);
      if (depth > 0) dashSheet.getRange(1, c).shiftColumnGroupDepth(-depth);
    }

    dashSheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);

    const solutionRowIndex = DATA_START_ROW_2026; 
    const productRowIndex = DATA_START_ROW_2026; 

    let colIdx = 6; // 1-based index for columns (1-5 are metadata)
    for (let i = 5; i < columnsToKeepInfo.length; i++) {
        const info = columnsToKeepInfo[i];
        if (info.type === 'spacer') {
            dashSheet.setColumnWidth(colIdx, 40);
            
            // Format the spacer header (row 1 to 8 vertically merged with Solution name)
            dashSheet.getRange(1, colIdx, DATA_START_ROW_2026 - 1, 1).merge()
               .setValue(info.solution)
               .setBackground(info.color)
               .setFontWeight("bold")
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle")
               .setTextRotation(90)
               .setWrap(true);

            // Format the lower spacer header (rows 9 and 10 separately)
            dashSheet.getRange(DATA_START_ROW_2026, colIdx, 2, 1)
               .setValue("") // Clear the text from rows 9 and 10
               .setVerticalAlignment("middle")
               .setTextRotation(90)
               .setWrap(true);
               
            // Find how many data columns follow it
            let productCount = 0;
            for (let j = i + 1; j < columnsToKeepInfo.length; j++) {
                if (columnsToKeepInfo[j].type === 'spacer') break;
                productCount++;
            }
            
            if (productCount > 0) {
                const prodStartCol = colIdx + 1;
                // We no longer merge Solution horizontally here, as it's purely vertical in spacer.
                // Group the product columns natively
                dashSheet.getRange(1, prodStartCol, 1, productCount).shiftColumnGroupDepth(1);
            }
        } else if (info.type === 'data') {
            dashSheet.setColumnWidth(colIdx, 70);
        }
        colIdx++;
    }

    // Now merge the product row (row 9)
    let prodMergeStart = 6; 
    let currentProd = outputValues[0][5];
    for (let c = 6; c <= outCols; c++) {
       const nextProd = c < outCols ? outputValues[0][c-1] : null; // outputValues is 0-indexed
       if (c === outCols || String(nextProd).trim() !== String(currentProd).trim() || String(currentProd).trim() === "") {
           const span = c - prodMergeStart;
           if (span > 1) {
               dashSheet.getRange(productRowIndex, prodMergeStart, 1, span).merge();
               // Create product-level Depth-2 grouping for Tier 2 - Tier 4, leaving Tier 1 as the visible toggle
               dashSheet.getRange(1, prodMergeStart + 1, 1, span - 1).shiftColumnGroupDepth(1);
           }
           if (c < outCols) {
               prodMergeStart = c;
               currentProd = nextProd;
           }
       }
    }

    // Wrap text on product row (row 9) from column 6 onwards
    dashSheet.getRange(DATA_START_ROW_2026, 6, 1, outCols - 5).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Style the 5 metadata header columns cleanly
    dashSheet.getRange(DATA_START_ROW_2026, 1, 2, 5)
        .setVerticalAlignment("middle")
        .setBackground("#f3f3f3")
        .setFontColor("black")
        .setFontWeight("bold")
        .setWrap(true);

    // Apply filter on Row 10 (DATA_START_ROW_2026 + 1) across ALL columns (starting from 1)
    const filterStartRow = DATA_START_ROW_2026 + 1; // row 10
    const existingFilter = dashSheet.getFilter();
    if (existingFilter) existingFilter.remove();
    dashSheet.getRange(filterStartRow, 1, dashSheet.getMaxRows() - filterStartRow + 1, dashSheet.getMaxColumns()).createFilter();

  } else {
    dashSheet.getRange(DATA_START_ROW_2026, 1).setValue("No partners found for this selection.");
  }
}
