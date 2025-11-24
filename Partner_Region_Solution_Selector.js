/**
 * ****************************************
 * Google Apps Script - Partner Dashboard Slicer
 * File: Partner_Region_Solution_Selector.gs
 * Version: 7.2 (Fixed Sorting Error)
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

const CELL_TYPE = {r: 3, c: 2};     
const CELL_REGION = {r: 4, c: 2};   
const CELL_COUNTRY = {r: 5, c: 2};  
const CELL_SOLUTION = {r: 6, c: 2}; 
const CELL_PRODUCT = {r: 7, c: 2};  
const CELL_STATUS = {r: 3, c: 4};   
const DATA_START_ROW = 9;

function setLoadingStatus(sheet, isLoading) {
  const cell = sheet.getRange(CELL_STATUS.r, CELL_STATUS.c);
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

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME_DASHBOARD) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  if (col === 2 && (row >= 3 && row <= 7)) {
    try {
      if (row === CELL_SOLUTION.r) {
        const newValue = e.value; const oldValue = e.oldValue;   
        if (newValue) {
          if (newValue === "All") e.range.setValue("All");
          else if (oldValue && oldValue !== "All") {
            const currentItems = oldValue.split(',').map(s => s.trim());
            const index = currentItems.indexOf(newValue);
            if (index > -1) { currentItems.splice(index, 1); e.range.setValue(currentItems.length === 0 ? "All" : currentItems.join(', ')); } 
            else { e.range.setValue(oldValue + ', ' + newValue); }
          } else e.range.setValue(newValue);
        } else e.range.setValue("All");
        SpreadsheetApp.flush(); 
      }
      
      setLoadingStatus(sheet, true);
      Utilities.sleep(10); 
      
      if (row === CELL_REGION.r) { sheet.getRange(CELL_COUNTRY.r, CELL_COUNTRY.c).setValue("All"); updateCountryDropdown(); }
      if (row === CELL_SOLUTION.r) { sheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).setValue("All"); updateProductDropdown(); }
      
      refreshDashboardData();
      
    } catch (err) {
      e.source.toast("Error: " + err.toString(), "Slicer Failed", 10);
      Logger.log(err);
    } finally {
      setLoadingStatus(sheet, false);
    }
  }
}

function refreshDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  const dbSheet = ss.getSheetByName(SHEET_NAME_DB);
  const scoreSheet = ss.getSheetByName(SHEET_NAME_SCORE);
  
  if (!dbSheet || !scoreSheet) {
    dashSheet.getRange(DATA_START_ROW, 1).setValue("Error: DB or Score Sheets missing. Please run 'Update Partner DB' from menu.");
    return;
  }

  // 1. Get Selection
  const typeSel = dashSheet.getRange(CELL_TYPE.r, CELL_TYPE.c).getValue();
  const regionSel = dashSheet.getRange(CELL_REGION.r, CELL_REGION.c).getValue();
  const countrySel = dashSheet.getRange(CELL_COUNTRY.r, CELL_COUNTRY.c).getValue();
  const solutionSel = String(dashSheet.getRange(CELL_SOLUTION.r, CELL_SOLUTION.c).getValue()).trim();
  const solutionSelArray = solutionSel === "All" ? ["All"] : solutionSel.split(',').map(s => s.trim());
  const productSel = String(dashSheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).getValue()).trim();
  
  // 2. Filter Partners
  const dbData = dbSheet.getDataRange().getValues();
  if (dbData.length < 2) {
    dashSheet.getRange(DATA_START_ROW, 1).setValue("Error: Partner DB is empty. Please run '1️⃣ Update Partner DB' first.");
    return;
  }

  const dbHeaders = dbData[0];
  const partnerMap = new Map(); 
  
  const idxName = 1; const idxCountry = 3; const idxManaged = 5; 
  const idxRegion = regionSel === "LATAM (All)" ? -1 : dbHeaders.indexOf(regionSel);
  
  for (let i = 1; i < dbData.length; i++) {
    const pName = dbData[i][idxName];
    const pCountryString = dbData[i][idxCountry];
    const isManaged = dbData[i][idxManaged] === true;
    const isRegion = idxRegion === -1 ? true : (dbData[i][idxRegion] === true);
    let countryArray = [];
    if (pCountryString) { countryArray = String(pCountryString).split(',').map(s => s.trim()); }

    // Parse Profile Breakdown
    const profileBreakdownStr = dbData[i][18]; // New Column at the end
    const profileMap = new Map();
    if (profileBreakdownStr) {
      profileBreakdownStr.split('|').forEach(pair => {
        const [country, count] = pair.split(':');
        if (country && count) profileMap.set(country.trim(), parseInt(count));
      });
    }
    const totalProfiles = dbData[i][2] || 0; // Total_Profiles column is index 2

    partnerMap.set(pName, {
      countries: countryArray,
      matchesRegion: isRegion,
      isManaged: isManaged,
      profileMap: profileMap,
      totalProfiles: totalProfiles
    });
  }
  
  // 3. Filter Columns
  const scoreRange = scoreSheet.getDataRange();
  const scoreValues = scoreRange.getValues();
  if (scoreValues.length < 3) {
    dashSheet.getRange(DATA_START_ROW, 1).setValue("Error: Score Matrix is empty. Please run '2️⃣ Update Scoring Matrix'.");
    return;
  }

  const scoreBackgrounds = scoreRange.getBackgrounds(); 
  const scoreFontWeights = scoreRange.getFontWeights();
  const rowSol = scoreValues[0];
  const rowProd = scoreValues[1];
  
  const columnsToKeep = [0, 1, 2]; // ID, Name, Total Profiles (will be overwritten)
  const effectiveHeaders = { sol: {}, prod: {} }; 
  
  for (let c = 3; c < rowSol.length; c++) {
    let prod = String(rowProd[c]).trim();
    let sol = String(rowSol[c]).trim(); 
    
    let effectiveSol = sol;
    if (effectiveSol === "") { 
        for (let k = c - 1; k >= 0; k--) { 
            if (String(rowSol[k]).trim() !== "") { effectiveSol = String(rowSol[k]).trim(); break; } 
        } 
    }
    
    let effectiveProd = prod;
    if (effectiveProd === "") { 
        for (let k = c - 1; k >= 0; k--) { 
            if (String(rowProd[k]).trim() !== "") { effectiveProd = String(rowProd[k]).trim(); break; } 
        } 
    }
    
    effectiveHeaders.sol[c] = effectiveSol; 
    effectiveHeaders.prod[c] = effectiveProd;

    let keepCol = true;
    if (!solutionSelArray.includes("All") && !solutionSelArray.includes(effectiveSol)) keepCol = false;
    if (productSel !== "All" && effectiveProd !== productSel) keepCol = false;
    if (keepCol) columnsToKeep.push(c);
  }
  
  // 4. Build Output Rows
  let outputValues = [], outputBackgrounds = [], outputWeights = [];
  
  // Headers
  for (let r = 0; r < 3; r++) {
    let rowV = [], rowB = [], rowW = [];
    columnsToKeep.forEach(idx => {
      if (idx === 2) {
        // Insert 3 columns for ALL rows to maintain column count
        if (r === 0) {
          rowV.push("", "", ""); // Empty for Solution header
        } else if (r === 1) {
          rowV.push("", "", ""); // Empty for Product header
        } else if (r === 2) {
          rowV.push("Total Profiles", "Region Profiles", "Country Profiles");
          }
        rowB.push(scoreBackgrounds[r][idx], scoreBackgrounds[r][idx], scoreBackgrounds[r][idx]);
        rowW.push(scoreFontWeights[r][idx], scoreFontWeights[r][idx], scoreFontWeights[r][idx]);
        return;
      }

      if (r === 0 && idx > 2) rowV.push(effectiveHeaders.sol[idx]);
      else if (r === 1 && idx > 2) rowV.push(effectiveHeaders.prod[idx]);
      else rowV.push(scoreValues[r][idx]);
      rowB.push(scoreBackgrounds[r][idx]);
      rowW.push(scoreFontWeights[r][idx]);
    });
    outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW);
  }
  
  // Data
  for (let r = 3; r < scoreValues.length; r++) {
    const pName = scoreValues[r][1]; 
    const meta = partnerMap.get(pName);
    let keepRow = false;
    if (meta) {
      if (meta.matchesRegion) {
        let countryMatch = false;
        if (countrySel === "All") countryMatch = true; else if (meta.countries.includes(countrySel)) countryMatch = true;
        if (countryMatch) {
            if (typeSel === "All") keepRow = true; else if (typeSel === "Managed" && meta.isManaged) keepRow = true; else if (typeSel === "UnManaged" && !meta.isManaged) keepRow = true;
        }
      }
    }
    if (keepRow) {
      let rowV = [], rowB = [], rowW = [];
      columnsToKeep.forEach(idx => {
        let val = scoreValues[r][idx];
        
        // INSTANT LINKING (VLOOKUP)
        if (idx === 1) { 
           const safeName = String(val).replace(/'/g, "''"); 
           const linkSheet = typeof SHEET_NAME_LINKS !== 'undefined' ? SHEET_NAME_LINKS : "System_Link_Cache";
           val = `=IFNA(HYPERLINK(VLOOKUP("${safeName}", ${linkSheet}!A:B, 2, FALSE), "${safeName}"), "${safeName}")`;
        }
        
        if (idx === 2) {
          // Insert 3 columns data
          const total = meta.totalProfiles;
          let regionTotal = 0;
          let countryTotal = 0;

          // Region Mapping
          const regionMapping = {
            'Brazil': ['Brazil'],
            'Mexico': ['Mexico'],
            'MCO': ['Argentina', 'Bolivia', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela'],
            'GSI': ['GSI'], // Per user request, treat as country
            'PS': ['PS']    // Per user request, treat as country
          };

          const currentRegion = regionSel === "LATAM (All)" ? null : regionSel;
          const currentCountry = countrySel === "All" ? null : countrySel;

          // Calculate Region Total
          if (currentRegion && regionMapping[currentRegion]) {
            regionMapping[currentRegion].forEach(c => {
              regionTotal += (meta.profileMap.get(c) || 0);
            });
          } else {
            regionTotal = total; // Default to total if region not mapped or "All"
          }

          // Calculate Country Total
          if (currentCountry) {
            countryTotal = meta.profileMap.get(currentCountry) || 0;
          } else {
            countryTotal = regionTotal; // Default to region total if country not selected
          }

          rowV.push(total, regionTotal, countryTotal);
          rowB.push(scoreBackgrounds[r][idx], scoreBackgrounds[r][idx], scoreBackgrounds[r][idx]);
          rowW.push(scoreFontWeights[r][idx], scoreFontWeights[r][idx], scoreFontWeights[r][idx]);
          return; // Skip normal push
        }

        rowV.push(val); 
        rowB.push(scoreBackgrounds[r][idx]); 
        rowW.push(scoreFontWeights[r][idx]); 
      });
      outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW);
    }
  }

  // 5. Sorting (FIXED)
  const headerValues = outputValues.slice(0, 3);
  const headerBackgrounds = outputBackgrounds.slice(0, 3);
  const headerWeights = outputWeights.slice(0, 3);
  if (outputValues.length > 3) {
    const dataValues = outputValues.slice(3);
    const dataBackgrounds = outputBackgrounds.slice(3);
    const dataWeights = outputWeights.slice(3);
    
    const combinedData = dataValues.map((val, index) => ({ 
        value: val, 
        background: dataBackgrounds[index], 
        weight: dataWeights[index] 
    }));
    
    // SAFE SORTING LOGIC
    combinedData.sort((a, b) => {
      let nameA = String(a.value[1]); 
      let nameB = String(b.value[1]);

      // Helper to extract clean name from formula: =... "Name"), "Name")
      const extractName = (str) => {
          if (str.startsWith("=IFNA")) {
             const parts = str.split(', "');
             // Safe check: ensure split actually worked
             if (parts.length > 1) {
                 return parts[parts.length - 1].replace('")', '');
             }
          }
          return str;
      };

      nameA = extractName(nameA);
      nameB = extractName(nameB);

      return nameA.toLowerCase().localeCompare(nameB.toLowerCase());
    });
    
    outputValues = [...headerValues, ...combinedData.map(i => i.value)];
    outputBackgrounds = [...headerBackgrounds, ...combinedData.map(i => i.background)];
    outputWeights = [...headerWeights, ...combinedData.map(i => i.weight)];
  }
  
  // 6. Write
  const lastRow = dashSheet.getLastRow(); const lastCol = dashSheet.getLastColumn();
  if (lastRow >= DATA_START_ROW) dashSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol || 1).clear();
  
  if (outputValues.length > 3) {
    const outRows = outputValues.length; const outCols = outputValues[0].length;
    const targetRange = dashSheet.getRange(DATA_START_ROW, 1, outRows, outCols);
    targetRange.setValues(outputValues);
    targetRange.setBackgrounds(outputBackgrounds);
    targetRange.setFontWeights(outputWeights);
    targetRange.setHorizontalAlignment("center");
    dashSheet.getRange(DATA_START_ROW, 2, outRows, 1).setHorizontalAlignment("left");
    dashSheet.getRange(DATA_START_ROW, 3, outRows, 1).setBackground("#d9d9d9");
    dashSheet.getRange(DATA_START_ROW, 1, outRows, outCols).setBorder(true, true, true, true, true, true);
    dashSheet.getRange(DATA_START_ROW, 1, 3, outCols).setBorder(true, true, true, true, true, true);
    for (let c = 6; c <= outCols; c++) { dashSheet.setColumnWidth(c, 50); }
    dashSheet.setColumnWidth(3, 80); // Total
    dashSheet.setColumnWidth(4, 80); // Region
    dashSheet.setColumnWidth(5, 80); // Country

    const solutionRowIndex = DATA_START_ROW; const productRowIndex = DATA_START_ROW + 1;
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
    
  } else {
    dashSheet.getRange(DATA_START_ROW, 1).setValue("No partners found for this selection.");
  }
}

function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_DASHBOARD); }
  sheet.clear();
  setLoadingStatus(sheet, true);
  
  sheet.getRange("A1").setValue("Partner & Solution Slicer").setFontSize(14).setFontWeight("bold");
  
  sheet.getRange("A3").setValue("Select Partner Type:"); 
  sheet.getRange("A4").setValue("Select Region:");
  sheet.getRange("A5").setValue("Select Country:");
  sheet.getRange("A6").setValue("Select Solution (Multi):");
  sheet.getRange("A7").setValue("Select Product:");
  
  const types = ["All", "Managed", "UnManaged"];
  sheet.getRange(CELL_TYPE.r, CELL_TYPE.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(types).build()).setValue(types[0]);

  const regions = ["LATAM (All)", "GSI", "Brazil", "MCO", "Mexico", "PS"];
  sheet.getRange(CELL_REGION.r, CELL_REGION.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(regions).build()).setValue(regions[0]);
  
  sheet.getRange(CELL_COUNTRY.r, CELL_COUNTRY.c).setValue("All");
  updateSolutionDropdown(sheet); 
  sheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).setValue("All");
  
  sheet.getRange("A3:A7").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("B3:B7").setBackground("#fff2cc").setFontWeight("bold");
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 300);
  
  updateCountryDropdown();
  updateProductDropdown();
  refreshDashboardData();
  setLoadingStatus(sheet, false);
}

function updateSolutionDropdown(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoreSheet = ss.getSheetByName(SHEET_NAME_SCORE);
  if (!scoreSheet) return;
  const headers = scoreSheet.getRange(1, 4, 1, scoreSheet.getLastColumn() - 3).getValues()[0];
  let solutions = new Set();
  solutions.add("All");
  headers.forEach(sol => { const cleanSol = String(sol).trim(); if (cleanSol !== "") solutions.add(cleanSol); });
  sheet.getRange(CELL_SOLUTION.r, CELL_SOLUTION.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from(solutions)).build()).setValue("All");
}

function updateCountryDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  const dbSheet = ss.getSheetByName(SHEET_NAME_DB);
  if (!dbSheet) return;
  const regionSelection = dashSheet.getRange(CELL_REGION.r, CELL_REGION.c).getValue();
  const data = dbSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0]; 
  let countries = new Set();
  let filterColIndex = -1;
  if (regionSelection !== "LATAM (All)") filterColIndex = headers.indexOf(regionSelection);
  for (let i = 1; i < data.length; i++) {
    const rawCountries = data[i][3]; 
    let processRow = false;
    if (filterColIndex === -1) processRow = true; else if (data[i][filterColIndex] === true) processRow = true;
    if (processRow && rawCountries) {
        const splitList = String(rawCountries).split(',');
        splitList.forEach(country => { const cleanCountry = country.trim(); if (cleanCountry !== "") countries.add(cleanCountry); });
    }
  }
  dashSheet.getRange(CELL_COUNTRY.r, CELL_COUNTRY.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All", ...Array.from(countries).sort()]).build());
}

function updateProductDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  const scoreSheet = ss.getSheetByName(SHEET_NAME_SCORE);
  if (!scoreSheet) return;
  const solutionSelectionString = String(dashSheet.getRange(CELL_SOLUTION.r, CELL_SOLUTION.c).getValue());
  if (solutionSelectionString === "All") {
     dashSheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All"]).build());
     return;
  }
  const selectedSolutions = solutionSelectionString.split(',').map(s => s.trim());
  const headers = scoreSheet.getRange(1, 1, 2, scoreSheet.getLastColumn()).getValues();
  const solutionsRow = headers[0];
  const productsRow = headers[1];
  let products = new Set();
  for (let c = 3; c < solutionsRow.length; c++) { 
    let effectiveSol = String(solutionsRow[c]).trim();
    if (effectiveSol === "") { for (let k = c - 1; k >= 0; k--) { if (String(solutionsRow[k]).trim() !== "") { effectiveSol = String(solutionsRow[k]).trim(); break; } } }
    if (selectedSolutions.includes(effectiveSol) && productsRow[c]) { products.add(String(productsRow[c]).trim()); }
  }
  dashSheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All", ...Array.from(products).sort()]).build());
}