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
 * Automatically triggers when a dropdown cell in LATAM_Partner_Dashboard_2026 is modified.
 */
function onEdit2026(e) {
  if (!e || !e.source) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME_DASHBOARD_2026) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Only react to changes in Column B, rows 3-7 (the slicers)
  if (col === 2 && (row >= 3 && row <= 7)) {
    try {
      // Handle Multi-Select for Solution
      if (row === CELL_SOLUTION_2026.r) {
        const newValue = e.value; 
        const oldValue = e.oldValue;   
        if (newValue) {
          if (newValue === "All") e.range.setValue("All");
          else if (oldValue && oldValue !== "All") {
            const currentItems = oldValue.split(',').map(s => s.trim());
            const index = currentItems.indexOf(newValue);
            if (index > -1) { 
                currentItems.splice(index, 1); 
                e.range.setValue(currentItems.length === 0 ? "All" : currentItems.join(', ')); 
            } 
            else { e.range.setValue(oldValue + ', ' + newValue); }
          } else e.range.setValue(newValue);
        } else e.range.setValue("All");
        SpreadsheetApp.flush(); 
      }
      
      setLoadingStatus2026(sheet, true);
      Utilities.sleep(10); 
      
      // Cascading Dropdown Logic
      if (row === CELL_SUB_REGION_2026.r) { 
          sheet.getRange(CELL_PDM_2026.r, CELL_PDM_2026.c).setValue("All"); 
          updatePDMDropdown2026(sheet); 
      }
      if (row === CELL_SOLUTION_2026.r) { 
          sheet.getRange(CELL_PRODUCT_2026.r, CELL_PRODUCT_2026.c).setValue("All"); 
          updateProductDropdown2026(sheet); 
      }
      
      refreshDashboardData2026(sheet);
      
    } catch (err) {
      e.source.toast("Error: " + err.toString(), "Slicer Failed", 10);
      try { sheet.getRange(1, 5).setValue("Error: " + err.toString()).setBackground("red").setFontColor("white"); } catch (e2) { }
    } finally {
      setLoadingStatus2026(sheet, false);
    }
  }
}

/**
 * Creates the dashboard layout and sets up initial dropdowns.
 */
function setupDashboard2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_DASHBOARD_2026);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_DASHBOARD_2026); }
  sheet.clear();
  setLoadingStatus2026(sheet, true);
  
  // UI setup
  sheet.getRange("A1").setValue("2026 Partner & Solution Slicer").setFontSize(14).setFontWeight("bold");
  sheet.getRange("A3").setValue("Select Partner Type:"); 
  sheet.getRange("A4").setValue("Select Sub-Region:");
  sheet.getRange("A5").setValue("Select PDM:");
  sheet.getRange("A6").setValue("Select Solution (Multi):");
  sheet.getRange("A7").setValue("Select Product:");
  
  sheet.getRange("A3:A7").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("B3:B7").setBackground("#fff2cc").setFontWeight("bold");
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 300);

  // Initialize Cache before building dropdowns
  updateDashboardCache2026();
  
  const cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE_2026);
  if(!cacheSheet) {
      sheet.getRange(DATA_START_ROW_2026, 1).setValue("Error establishing cache.");
      return;
  }
  
  // Extract unique types and sub-regions
  const data = cacheSheet.getDataRange().getValues();
  const types = new Set();
  const regions = new Set();
  
  for(let i=3; i<data.length; i++) {
     types.add(String(data[i][5]).trim()); // F: Type
     regions.add(String(data[i][3]).trim()); // D: Sub-Region
  }

  // Set Dropdowns
  const typeList = ["All"].concat(Array.from(types).sort());
  sheet.getRange(CELL_TYPE_2026.r, CELL_TYPE_2026.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(typeList).build()).setValue("All");

  const regionList = ["All"].concat(Array.from(regions).sort());
  sheet.getRange(CELL_SUB_REGION_2026.r, CELL_SUB_REGION_2026.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(regionList).build()).setValue("All");
  
  sheet.getRange(CELL_PDM_2026.r, CELL_PDM_2026.c).setValue("All");
  updateSolutionDropdown2026(sheet, cacheSheet); 
  sheet.getRange(CELL_PRODUCT_2026.r, CELL_PRODUCT_2026.c).setValue("All");
  
  updatePDMDropdown2026(sheet);
  updateProductDropdown2026(sheet);

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

function updateSolutionDropdown2026(sheet, cacheSheet) {
  const headers = cacheSheet.getRange(1, 8, 1, cacheSheet.getLastColumn() - 7).getValues()[0];
  let solutions = new Set();
  solutions.add("All");
  headers.forEach(sol => { const cleanSol = String(sol).trim(); if (cleanSol !== "") solutions.add(cleanSol); });
  sheet.getRange(CELL_SOLUTION_2026.r, CELL_SOLUTION_2026.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from(solutions)).build()).setValue("All");
}

function updatePDMDropdown2026(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE_2026);
  if (!cacheSheet) return;
  
  const regionSelection = sheet.getRange(CELL_SUB_REGION_2026.r, CELL_SUB_REGION_2026.c).getValue();
  const data = cacheSheet.getDataRange().getValues();
  
  let pdms = new Set();
  for (let i = 3; i < data.length; i++) {
    const rowRegion = String(data[i][3]).trim(); // D: Sub-Region
    const rowPDM = String(data[i][4]).trim();    // E: PDM
    if (rowPDM && (regionSelection === "All" || rowRegion === regionSelection)) {
       pdms.add(rowPDM);
    }
  }
  sheet.getRange(CELL_PDM_2026.r, CELL_PDM_2026.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All", ...Array.from(pdms).sort()]).build());
}

function updateProductDropdown2026(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE_2026);
  if (!cacheSheet) return;
  
  const solutionSelectionString = String(sheet.getRange(CELL_SOLUTION_2026.r, CELL_SOLUTION_2026.c).getValue());
  if (solutionSelectionString === "All") {
     sheet.getRange(CELL_PRODUCT_2026.r, CELL_PRODUCT_2026.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All"]).build());
     return;
  }
  
  const selectedSolutions = solutionSelectionString.split(',').map(s => s.trim());
  const headers = cacheSheet.getRange(1, 1, 2, cacheSheet.getLastColumn()).getValues();
  const solutionsRow = headers[0];
  const productsRow = headers[1];
  
  let products = new Set();
  for (let c = 7; c < solutionsRow.length; c++) { 
    let effectiveSol = String(solutionsRow[c]).trim();
    if (effectiveSol === "") { 
        for (let k = c - 1; k >= 7; k--) { 
            if (String(solutionsRow[k]).trim() !== "") { effectiveSol = String(solutionsRow[k]).trim(); break; } 
        } 
    }
    if (selectedSolutions.includes(effectiveSol) && productsRow[c]) { 
        products.add(String(productsRow[c]).trim()); 
    }
  }
  sheet.getRange(CELL_PRODUCT_2026.r, CELL_PRODUCT_2026.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All", ...Array.from(products).sort()]).build());
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

  // 1. Get Selection
  const typeSel = String(dashSheet.getRange(CELL_TYPE_2026.r, CELL_TYPE_2026.c).getValue()).trim();
  const regionSel = String(dashSheet.getRange(CELL_SUB_REGION_2026.r, CELL_SUB_REGION_2026.c).getValue()).trim();
  const pdmSel = String(dashSheet.getRange(CELL_PDM_2026.r, CELL_PDM_2026.c).getValue()).trim();
  const solutionSel = String(dashSheet.getRange(CELL_SOLUTION_2026.r, CELL_SOLUTION_2026.c).getValue()).trim();
  const solutionSelArray = solutionSel === "All" ? ["All"] : solutionSel.split(',').map(s => s.trim().toLowerCase());
  const productSel = String(dashSheet.getRange(CELL_PRODUCT_2026.r, CELL_PRODUCT_2026.c).getValue()).trim();

  // 2. Read Cache
  const cacheRange = cacheSheet.getDataRange();
  const cacheValues = cacheRange.getValues();
  const cacheBackgrounds = cacheRange.getBackgrounds();
  const cacheWeights = cacheRange.getFontWeights();
  const cacheFontColors = cacheRange.getFontColors(); 

  if (cacheValues.length < 3) return;

  const rowSol = cacheValues[0];
  const rowProd = cacheValues[1];

  // 3. Columns to Keep (Metadata: Name(C), Sub-Region(D), PDM(E), Type(F), Profiles(G) -> Indexes 2,3,4,5,6)
  // We drop Internal ID and Partner ID from the display.
  const columnsToKeep = [2, 3, 4, 5, 6]; 
  const effectiveHeaders = { sol: {}, prod: {} }; 
  
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

    let keepCol = true;
    if (!solutionSelArray.includes("All") && !solutionSelArray.includes(effectiveSol.toLowerCase())) keepCol = false;
    if (productSel !== "All" && effectiveProd.toLowerCase() !== productSel.toLowerCase()) keepCol = false;
    if (keepCol) columnsToKeep.push(c);
  }

  // 4. Build Output Data
  let outputValues = [], outputBackgrounds = [], outputWeights = [], outputFontColors = [];

  // Headers (3 rows)
  for (let r = 0; r < 3; r++) {
    let rowV = [], rowB = [], rowW = [], rowFC = [];
    columnsToKeep.forEach(idx => {
      let val = cacheValues[r][idx];
      if (idx >= 7) {
        if (r === 0) val = effectiveHeaders.sol[idx];
        if (r === 1) val = effectiveHeaders.prod[idx];
      }
      rowV.push(val);
      rowB.push(cacheBackgrounds[r][idx]);
      rowW.push(cacheWeights[r][idx]);
      rowFC.push(cacheFontColors[r][idx]);
    });
    outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW); outputFontColors.push(rowFC);
  }

  // Data Rows
  for (let r = 3; r < cacheValues.length; r++) {
    const rowType = String(cacheValues[r][5]).trim();
    const rowRegion = String(cacheValues[r][3]).trim();
    const rowPDM = String(cacheValues[r][4]).trim();

    let keepRow = true;
    if (typeSel !== "All" && rowType.toLowerCase() !== typeSel.toLowerCase()) keepRow = false;
    if (regionSel !== "All" && rowRegion.toLowerCase() !== regionSel.toLowerCase()) keepRow = false;
    if (pdmSel !== "All" && rowPDM.toLowerCase() !== pdmSel.toLowerCase()) keepRow = false;

    if (keepRow) {
      let rowV = [], rowB = [], rowW = [], rowFC = [];
      columnsToKeep.forEach(idx => {
        rowV.push(cacheValues[r][idx]);
        rowB.push(cacheBackgrounds[r][idx]);
        rowW.push(cacheWeights[r][idx]);
        rowFC.push(cacheFontColors[r][idx]);
      });
      outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW); outputFontColors.push(rowFC);
    }
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

    targetRange.setHorizontalAlignment("center");
    dashSheet.getRange(DATA_START_ROW_2026, 1, outRows, 1).setHorizontalAlignment("left"); // Partner Name
    dashSheet.getRange(DATA_START_ROW_2026, 1, outRows, outCols).setBorder(true, true, true, true, true, true);
    dashSheet.getRange(DATA_START_ROW_2026, 1, 3, outCols).setBorder(true, true, true, true, true, true);
    
    // Auto Resize width for the first 5 metadata columns
    for(let c=1; c<=5; c++) dashSheet.autoResizeColumn(c);
    
    // Standardize width for score columns
    for (let c = 6; c <= outCols; c++) { dashSheet.setColumnWidth(c, 70); }

    // Re-merge top header blocks matching visually
    const solutionRowIndex = DATA_START_ROW_2026; 
    const productRowIndex = DATA_START_ROW_2026 + 1;
    let solMergeStart = 6; let currentSol = outputValues[0][5];
    let prodMergeStart = 6; let currentProd = outputValues[1][5];

    for (let c = 6; c < outCols; c++) { 
       const nextSol = outputValues[0][c]; const nextProd = outputValues[1][c];
       if (String(nextSol).trim() !== String(currentSol).trim() || String(currentSol).trim() === "") {
          const span = c - (solMergeStart - 1); if (span > 1) dashSheet.getRange(solutionRowIndex, solMergeStart, 1, span).merge();
          solMergeStart = c + 1; currentSol = nextSol;
       }
       if (String(nextProd).trim() !== String(currentProd).trim() || String(currentProd).trim() === "") {
           const span = c - (prodMergeStart - 1); if (span > 1) dashSheet.getRange(productRowIndex, prodMergeStart, 1, span).merge();
           prodMergeStart = c + 1; currentProd = nextProd;
       }
    }
    const solSpan = outCols - (solMergeStart - 1); if (solSpan > 1) dashSheet.getRange(solutionRowIndex, solMergeStart, 1, solSpan).merge();
    const prodSpan = outCols - (prodMergeStart - 1); if (prodSpan > 1) dashSheet.getRange(productRowIndex, prodMergeStart, 1, prodSpan).merge();
    
    // Vertically merge and style the 5 metadata header columns so they span the 3 header rows cleanly
    for (let c = 1; c <= 5; c++) {
      dashSheet.getRange(DATA_START_ROW_2026, c, 3, 1)
        .mergeVertically()
        .setVerticalAlignment("middle")
        .setBackground("#f3f3f3")
        .setFontColor("black")
        .setFontWeight("bold")
        .setWrap(true);
    }

  } else {
    dashSheet.getRange(DATA_START_ROW_2026, 1).setValue("No partners found for this selection.");
  }
}
