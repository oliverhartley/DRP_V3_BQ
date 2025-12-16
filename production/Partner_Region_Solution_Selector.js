/**
 * ****************************************
 * Google Apps Script - Partner Dashboard Slicer
 * File: Partner_Region_Solution_Selector.gs
 * Version: 8.1 - Fast Cache (No Rich Text)
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
const SHEET_NAME_CACHE = "CACHE_Dashboard_Data";

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
      try {
        sheet.getRange(1, 5).setValue("Error: " + err.toString()).setBackground("red").setFontColor("white");
      } catch (e2) { }
    } finally {
      setLoadingStatus(sheet, false);
    }
  }
}

/**
 * Updates the Dashboard Cache Sheet.
 * This should be run daily or after data updates.
 */
function updateDashboardCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAME_DB);
  const scoreSheet = ss.getSheetByName(SHEET_NAME_SCORE);
  const baselineSheet = ss.getSheetByName("LATAM_Partner_Score_DRP_Nov1"); // Baseline Sheet
  
  if (!dbSheet || !scoreSheet) {
    throw new Error("DB or Score Sheets missing.");
  }

  ss.toast("Updating Dashboard Cache...", "Processing", 30);

  // 1. Load Data
  const dbData = dbSheet.getDataRange().getValues();
  const scoreRange = scoreSheet.getDataRange();
  const scoreValues = scoreRange.getValues();
  const scoreBackgrounds = scoreRange.getBackgrounds();
  const scoreFontWeights = scoreRange.getFontWeights();

  // 2. Load Baseline & Link Cache
  const baselineMap = new Map();
  if (baselineSheet) {
    const baselineData = baselineSheet.getDataRange().getValues();
    baselineData.slice(3).forEach(row => {
      baselineMap.set(row[0], row); // ID -> Row
    });
  }

  const linkMap = new Map();
  const linkSheetName = typeof SHEET_NAME_LINKS !== 'undefined' ? SHEET_NAME_LINKS : "System_Link_Cache";
  const linkSheet = ss.getSheetByName(linkSheetName);
  if (linkSheet) {
    const linkData = linkSheet.getDataRange().getValues();
    linkData.forEach(row => {
      if (row[0] && row[1]) linkMap.set(String(row[0]).trim(), row[1]); // Name -> URL
    });
  }

  // 3. Map DB Data (Metadata)
  const dbHeaders = dbData[0];
  const idxName = 1;
  const idxCountry = 3;
  const idxManaged = 5;
  const idxTotalProfiles = dbHeaders.indexOf("Total_Profiles");
  const idxProfileBreakdown = dbHeaders.indexOf("Profile_Breakdown");

  // Region Columns
  const regions = ["Brazil", "Mexico", "MCO", "GSI", "PS"];
  const regionIndices = {};
  regions.forEach(r => { regionIndices[r] = dbHeaders.indexOf(r); });

  const partnerMetaMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const pName = dbData[i][idxName];
    const regionFlags = {};
    regions.forEach(r => {
      regionFlags[r] = (regionIndices[r] !== -1 && dbData[i][regionIndices[r]] === true);
    });

    partnerMetaMap.set(pName, {
      countries: String(dbData[i][idxCountry] || ""),
      isManaged: dbData[i][idxManaged] === true,
      totalProfiles: idxTotalProfiles !== -1 ? (dbData[i][idxTotalProfiles] || 0) : 0,
      profileBreakdown: idxProfileBreakdown !== -1 ? String(dbData[i][idxProfileBreakdown] || "") : "",
      regionFlags: regionFlags
    });
  }

  // 4. Build Cache Rows
  // Structure: [Score Columns...] + [Metadata JSON]
  // We will store Metadata in the last column as a JSON string

  const cacheValues = [];
  const cacheBackgrounds = [];
  const cacheWeights = [];
  const cacheFontColors = []; // NEW: Store font colors explicitly

  // Headers (First 3 rows)
  for (let r = 0; r < 3; r++) {
    const rowV = [...scoreValues[r]];
    const rowB = [...scoreBackgrounds[r]];
    const rowW = [...scoreFontWeights[r]];
    const rowFC = rowV.map(() => "#000000");

    // Append Metadata Header
    if (r === 0) rowV.push("METADATA_JSON"); else rowV.push("");
    rowB.push("#ffffff");
    rowW.push("normal");
    rowFC.push("#000000");

    cacheValues.push(rowV);
    cacheBackgrounds.push(rowB);
    cacheWeights.push(rowW);
    cacheFontColors.push(rowFC);
  }

  // Data Rows
  for (let r = 3; r < scoreValues.length; r++) {
    const pId = scoreValues[r][0];
    const pName = scoreValues[r][1];
    const meta = partnerMetaMap.get(pName);
    const baselineRow = baselineMap.get(pId);

    if (!meta) continue; // Skip if not in DB

    const rowV = [];
    const rowB = [];
    const rowW = [];
    const rowFC = [];
    let currentDashboardUrl = null; // Fix: Declare variable

    for (let c = 0; c < scoreValues[r].length; c++) {
      let val = scoreValues[r][c];
      let bg = scoreBackgrounds[r][c];
      let wt = scoreFontWeights[r][c];
      let fc = "#000000"; // Default Black

      // INSTANT LINKING (Col 1)
      if (c === 1) {
        const url = linkMap.get(String(val).trim());
        if (url) {
          currentDashboardUrl = url; // Store for metadata
          val = `=HYPERLINK("${url}", "${val}")`;
          fc = "#1155cc"; // Link Blue
        } else {
          fc = "#000000"; // Normal Black
        }
      }

      // TOTAL PROFILES DELTA (Col 2)
      if (c === 2) {
        const total = meta.totalProfiles;
        let displayTotal = total === 0 ? "-" : total;
        if (baselineRow) {
          const baselineTotal = parseFloat(baselineRow[2]) || 0;
          const delta = total - baselineTotal;
          if (delta !== 0) {
            const deltaStr = delta > 0 ? ` (+${delta})` : ` (${delta})`;
            displayTotal = `${total} /${deltaStr}`;
            fc = delta > 0 ? "#38761d" : "#cc0000"; // Green / Red (Whole Cell)
          }
        }
        val = displayTotal;
      }

      // PRODUCT DELTAS (Col > 2)
      if (c > 2 && baselineRow) {
        const currentVal = parseFloat(val) || 0;
        const baselineVal = parseFloat(baselineRow[c]) || 0;
        const delta = currentVal - baselineVal;
        if (delta !== 0) {
          const deltaStr = delta > 0 ? ` (+${delta})` : ` (${delta})`;
          val = `${currentVal} /${deltaStr}`;
          fc = delta > 0 ? "#38761d" : "#cc0000"; // Green / Red (Whole Cell)
        } else {
          if (currentVal === 0) val = "-";
        }
      } else if (c > 2 && !baselineRow) {
        // No baseline, just format zero
        if (parseFloat(val) === 0) val = "-";
      }

      rowV.push(val);
      rowB.push(bg);
      rowW.push(wt);
      rowFC.push(fc);
    }

    // Append Metadata
    const metadata = {
      countries: meta.countries,
      isManaged: meta.isManaged,
      profileBreakdown: meta.profileBreakdown,
      regionFlags: meta.regionFlags,
      totalProfiles: meta.totalProfiles,
      dashboardUrl: currentDashboardUrl // Persist URL
    };
    rowV.push(JSON.stringify(metadata));
    rowB.push("#ffffff");
    rowW.push("normal");
    rowFC.push("#000000");

    cacheValues.push(rowV);
    cacheBackgrounds.push(rowB);
    cacheWeights.push(rowW);
    cacheFontColors.push(rowFC);
  }

  // 5. Write to Cache Sheet
  let cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(SHEET_NAME_CACHE);
    cacheSheet.hideSheet();
  }
  cacheSheet.clear();

  if (cacheValues.length > 0) {
    const range = cacheSheet.getRange(1, 1, cacheValues.length, cacheValues[0].length);
    range.setValues(cacheValues);
    range.setBackgrounds(cacheBackgrounds);
    range.setFontWeights(cacheWeights);
    range.setFontColors(cacheFontColors); // Fast!
  }

  ss.toast("Dashboard Cache Updated!", "Success", 5);
}

function refreshDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  const cacheSheet = ss.getSheetByName(SHEET_NAME_CACHE);

  if (!cacheSheet) {
    dashSheet.getRange(DATA_START_ROW, 1).setValue("Error: Cache missing. Please run 'Update Dashboard Cache' from menu.");
    return;
  }

  // 1. Get Selection
  const typeSel = dashSheet.getRange(CELL_TYPE.r, CELL_TYPE.c).getValue();
  const regionSel = dashSheet.getRange(CELL_REGION.r, CELL_REGION.c).getValue();
  const countrySel = dashSheet.getRange(CELL_COUNTRY.r, CELL_COUNTRY.c).getValue();
  const solutionSel = String(dashSheet.getRange(CELL_SOLUTION.r, CELL_SOLUTION.c).getValue()).trim();
  const solutionSelArray = solutionSel === "All" ? ["All"] : solutionSel.split(',').map(s => s.trim());
  const productSel = String(dashSheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).getValue()).trim();

  // 2. Read Cache
  const cacheRange = cacheSheet.getDataRange();
  const cacheValues = cacheRange.getValues();
  const cacheBackgrounds = cacheRange.getBackgrounds();
  const cacheWeights = cacheRange.getFontWeights();
  const cacheFontColors = cacheRange.getFontColors(); // Read Colors

  if (cacheValues.length < 3) return;

  const rowSol = cacheValues[0];
  const rowProd = cacheValues[1];
  const metaColIdx = rowSol.length - 1; // Metadata is last column

  // 3. Filter Columns
  const columnsToKeep = [0, 1, 2]; // ID, Name, Total Profiles
  const effectiveHeaders = { sol: {}, prod: {} }; 
  
  for (let c = 3; c < metaColIdx; c++) {
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

  // 4. Build Output
  let outputValues = [], outputBackgrounds = [], outputWeights = [], outputFontColors = [];

  // Headers
  for (let r = 0; r < 3; r++) {
    let rowV = [], rowB = [], rowW = [], rowFC = [];
    columnsToKeep.forEach(idx => {
      if (idx === 2) {
        if (r === 0) rowV.push("", "", "");
        else if (r === 1) rowV.push("", "", "");
        else if (r === 2) rowV.push("Total Profiles", "Region Profiles", "Country Profiles");

        rowB.push("#d9d9d9", "#d9d9d9", "#d9d9d9"); // Force Gray for Headers C, D, E
        rowW.push(cacheWeights[r][idx], cacheWeights[r][idx], cacheWeights[r][idx]);
        rowFC.push(cacheFontColors[r][idx], cacheFontColors[r][idx], cacheFontColors[r][idx]);
        return;
      }

      // FIX: Use Effective Headers for Solution (r=0) and Product (r=1) rows
      // This ensures that even if we pick a column that was originally a "middle" cell of a merge (and thus empty),
      // we populate it with the correct header text so the merge logic below works.
      let val = cacheValues[r][idx];
      if (idx >= 3) {
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

  // Data
  for (let r = 3; r < cacheValues.length; r++) {
    const metaJson = cacheValues[r][metaColIdx];
    let meta = {};
    try { meta = JSON.parse(metaJson); } catch (e) { }

  // Filter Rows
    let keepRow = false;

    // Region Check
    let regionMatch = false;
    if (regionSel === "LATAM (All)") regionMatch = true;
    else if (meta.regionFlags && meta.regionFlags[regionSel] === true) regionMatch = true;

    if (regionMatch) {
      // Country Check
      let countryMatch = false;
      if (countrySel === "All") countryMatch = true;
      else if (meta.countries && meta.countries.includes(countrySel)) countryMatch = true;

      if (countryMatch) {
        // Type Check
        if (typeSel === "All") keepRow = true;
        else if (typeSel === "Managed" && meta.isManaged) keepRow = true;
        else if (typeSel === "UnManaged" && !meta.isManaged) keepRow = true;
      }
    }

    if (keepRow) {
      let rowV = [], rowB = [], rowW = [], rowFC = [];
      columnsToKeep.forEach(idx => {
        if (idx === 2) {
          // Dynamic Profile Calculation
          const totalVal = cacheValues[r][idx];
          const totalFC = cacheFontColors[r][idx];

          let regionTotal = 0;
          let countryTotal = 0;

          // Parse Breakdown
          const breakdownMap = new Map();
          if (meta.profileBreakdown) {
            meta.profileBreakdown.split('|').forEach(pair => {
              const [c, n] = pair.split(':');
              if (c && n) breakdownMap.set(c.trim(), parseInt(n));
            });
          }

          // Region Total
          const regionMapping = {
            'Brazil': ['Brazil'],
            'Mexico': ['Mexico'],
            'MCO': ['Argentina', 'Bolivia', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela'],
            'GSI': ['GSI'],
            'PS': ['PS']
          };

          if (regionSel === "LATAM (All)") {
            regionTotal = meta.totalProfiles; 
          } else {
            const countriesInRegion = regionMapping[regionSel] || [];
            countriesInRegion.forEach(c => regionTotal += (breakdownMap.get(c) || 0));
          }

          // Country Total
          if (countrySel === "All") {
            countryTotal = regionTotal;
          } else {
            countryTotal = breakdownMap.get(countrySel) || 0;
          }

          // Format Zeros as "-"
          const fmt = (v) => v === 0 ? "-" : v;

          rowV.push(totalVal, fmt(regionTotal), fmt(countryTotal));
          rowB.push("#d9d9d9", "#d9d9d9", "#d9d9d9"); // Force Gray for Data C, D, E
          rowW.push(cacheWeights[r][idx], cacheWeights[r][idx], cacheWeights[r][idx]);
          rowFC.push(totalFC, "#000000", "#000000"); // Region/Country totals are black
          return;
        }

        // RECONSTRUCT HYPERLINK FROM METADATA
        let val = cacheValues[r][idx];
        if (idx === 1 && meta.dashboardUrl) {
          val = `=HYPERLINK("${meta.dashboardUrl}", "${val}")`;
        }

        rowV.push(val);
        rowB.push(cacheBackgrounds[r][idx]);
        rowW.push(cacheWeights[r][idx]);
        rowFC.push(cacheFontColors[r][idx]);
      });
      outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW); outputFontColors.push(rowFC);
    }
  }

  // 5. Sorting
  const headerValues = outputValues.slice(0, 3);
  const headerBackgrounds = outputBackgrounds.slice(0, 3);
  const headerWeights = outputWeights.slice(0, 3);
  const headerFontColors = outputFontColors.slice(0, 3);

  if (outputValues.length > 3) {
    const dataValues = outputValues.slice(3);
    const dataBackgrounds = outputBackgrounds.slice(3);
    const dataWeights = outputWeights.slice(3);
    const dataFontColors = outputFontColors.slice(3);
    
    const combinedData = dataValues.map((val, index) => ({ 
        value: val, 
        background: dataBackgrounds[index], 
      weight: dataWeights[index],
      fontColor: dataFontColors[index]
    }));

    combinedData.sort((a, b) => {
      let nameA = String(a.value[1]); 
      let nameB = String(b.value[1]);
      const extractName = (str) => {
          if (str.startsWith("=IFNA")) {
             const parts = str.split(', "');
            if (parts.length > 1) return parts[parts.length - 1].replace('")', '');
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
    outputFontColors = [...headerFontColors, ...combinedData.map(i => i.fontColor)];
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
    targetRange.setFontColors(outputFontColors); // Apply Colors

    targetRange.setHorizontalAlignment("center");
    dashSheet.getRange(DATA_START_ROW, 2, outRows, 1).setHorizontalAlignment("left");
    dashSheet.getRange(DATA_START_ROW, 3, outRows, 1).setBackground("#d9d9d9");
    dashSheet.getRange(DATA_START_ROW, 1, outRows, outCols).setBorder(true, true, true, true, true, true);
    dashSheet.getRange(DATA_START_ROW, 1, 3, outCols).setBorder(true, true, true, true, true, true);
    for (let c = 6; c <= outCols; c++) { dashSheet.setColumnWidth(c, 70); }
    dashSheet.setColumnWidth(3, 80);
    dashSheet.setColumnWidth(4, 80);
    dashSheet.setColumnWidth(5, 80); 

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

  // Try to use cache if available, otherwise warn
  try {
    refreshDashboardData();
  } catch (e) {
    sheet.getRange(DATA_START_ROW, 1).setValue("Please run 'Update Dashboard Cache' to initialize data.");
  }
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