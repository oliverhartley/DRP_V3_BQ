/**
 * ****************************************
 * Google Apps Script - Database & Persistence
 * File: Database.js
 * Description: Management of DB_Partners (Managed), DB_Reference, and Migration.
 * ****************************************
 */

/**
 * Initializes the full database structure:
 * 1. DB_Managed_Context (Manual Source - Minimal)
 * 2. DB_Reference (Product/Solution Mappings)
 * 3. CACHE_Partner_Landscape (Automated View)
 */
function initSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Initialize DB_Managed_Context
  let sheetContext = ss.getSheetByName(SHEETS.DB_MANAGED_CONTEXT);
  if (!sheetContext) {
    sheetContext = ss.insertSheet(SHEETS.DB_MANAGED_CONTEXT);
    sheetContext.setTabColor("16a765"); // Green for "Source"
  }

  const headersContext = [
    "Partner Name", "Domain", 
    "Email To", "Email CC"
  ];

  if (sheetContext.getLastRow() === 0) {
    sheetContext.getRange(1, 1, 1, headersContext.length)
      .setValues([headersContext])
      .setBackground("#16a765")
      .setFontColor("white")
      .setFontWeight("bold");
    sheetContext.setFrozenRows(1);
    sheetContext.setFrozenColumns(2);
  }

  // 2. Initialize DB_Reference
  let sheetRef = ss.getSheetByName(SHEETS.DB_REFERENCE);
  if (!sheetRef) {
    sheetRef = ss.insertSheet(SHEETS.DB_REFERENCE);
    sheetRef.setTabColor("ea4335"); // Red for "Reference/Admin"
  }

  const headersRef = ["Product Name", "Solution Category", "BigQuery Key"];
  if (sheetRef.getLastRow() === 0) {
    sheetRef.getRange(1, 1, 1, headersRef.length)
      .setValues([headersRef])
      .setBackground("#ea4335")
      .setFontColor("white")
      .setFontWeight("bold");
  }

  // 3. Initialize CACHE_Partner_Landscape
  let sheetLandscape = ss.getSheetByName(SHEETS.CACHE_PARTNER_LANDSCAPE);
  if (!sheetLandscape) {
    sheetLandscape = ss.insertSheet(SHEETS.CACHE_PARTNER_LANDSCAPE);
    sheetLandscape.setTabColor("4285f4"); // Blue for "Automated View"
  }
  // Headers for Landscape are dynamic/rebuilt daily, but we can init blank.
  
  Logger.log("System initialized: Managed Context & Reference ready.");
}

/**
 * Migrates data from Master Source to DB_Managed_Context.
 * Minimal Import: Name, Domain, Emails.
 */
function runMigration() {
  initSystem(); 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(SHEETS.DB_MANAGED_CONTEXT);

  // Safety check
  if (targetSheet.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Migration Warning', 'DB_Managed_Context already has data. Append?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  const masterSS = SpreadsheetApp.openById(MASTER_SOURCE_SS_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_SHEET_NAME);
  const data = masterSheet.getDataRange().getValues();

  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partnerName = row[33]; // AH
    const domain = row[34];      // AI
    
    if (!partnerName || !domain) continue;

    const newRow = [
      partnerName,
      domain,
      row[35], // Email To
      row[36]  // Email CC
    ];
    
    results.push(newRow);
  }

  if (results.length > 0) {
    const startRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startRow, 1, results.length, results[0].length).setValues(results);
  }
  
  Logger.log(`Migration complete. ${results.length} partners imported into DB_Managed_Context.`);
}

/**
 * Rebuilds the Partner Landscape (CACHE_Partner_Landscape).
 * - FEEDS inputs for "Managed" Partners from DB_Managed_Context.
 * - JOINS with BigQuery for "Unmanaged" and Presence Flags.
 * - OUTPUTS a single comprehensive table.
 */
function rebuildPartnerLandscape() {
  initSystem();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Read Managed Context
  const sheetContext = ss.getSheetByName(SHEETS.DB_MANAGED_CONTEXT);
  const contextData = sheetContext.getDataRange().getValues();
  const managedMap = new Map(); // Domain -> {Name, EmailTo, EmailCC}
  
  // Skip header, store Managed info
  for (let i = 1; i < contextData.length; i++) {
    const domain = String(contextData[i][1]).toLowerCase().trim();
    if(domain) {
      managedMap.set(domain, {
        name: contextData[i][0],
        emailTo: contextData[i][2],
        emailCC: contextData[i][3]
      });
    }
  }

  // 2. Fetch BigQuery Data (All Latam + Countries)
  // We need a combined query: Get All Partners + Their Countries
  const sql = getPartnerCountryPresenceSql(); 
  // NOTE: getPartnerCountryPresenceSql returns [Domain, CountryArray]. 
  // We also need Partner Name for unmanaged ones. 
  // Let's rely on BQ for unmanaged names, but prefer Managed Context name if exists.
  
  // Update: We need modified SQL or logic to separate "Name" fetching if not in presence SQL.
  // Actually, let's use `getAllLatamPartnersSql` which gives [Domain, Name, Country] (flattened).
  // We can aggregate locally to avoid complex array logic in SQL.
  
  const rawBqData = executeBigQuery(getAllLatamPartnersSql()); // [Domain, Name, Country]
  if (!rawBqData) return;

  // 3. Process & Merge
  // Map: Domain -> { BQ_Name, Set(Countries) }
  const bqMap = new Map();
  
  for (let i = 1; i < rawBqData.length; i++) {
    const domain = String(rawBqData[i][0]).toLowerCase().trim();
    const name = rawBqData[i][1];
    const country = rawBqData[i][2];
    
    if (!bqMap.has(domain)) {
      bqMap.set(domain, { name: name, countries: new Set() });
    }
    bqMap.get(domain).countries.add(country);
  }
  
  // 4. Construct Landscape Table
  // Schema: Partner Name, Domain, Managed(T/F), EmailTo, EmailCC, [Countries...]
  
  // Countries Header Map
  const countryList = [
    'Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 
    'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 
    'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 
    'Nicaragua', 'Panama', 'Paraguay', 'Peru', 
    'Uruguay', 'Venezuela'
  ];
  
  const headers = [
    "Partner Name", "Domain", "Managed Partner", "Email To", "Email CC", ...countryList
  ];
  
  const finalRows = [];
  
  // Set of all domains (Managed + BQ)
  const allDomains = new Set([...managedMap.keys(), ...bqMap.keys()]);
  
  const sortedDomains = Array.from(allDomains).sort();
  
  for (const domain of sortedDomains) {
    const isManaged = managedMap.has(domain);
    const managedInfo = managedMap.get(domain) || {};
    const bqInfo = bqMap.get(domain) || { name: "Unknown", countries: new Set() };
    
    // Prefer Managed Name, fallback to BQ Name
    const displayName = isManaged ? managedInfo.name : bqInfo.name;
    
    const row = [
      displayName,
      domain,
      isManaged, // Managed Partner Flag
      managedInfo.emailTo || "",
      managedInfo.emailCC || ""
    ];
    
    // Append Country Flags
    for (const country of countryList) {
      row.push(bqInfo.countries.has(country));
    }
    
    finalRows.push(row);
  }
  
  // 5. Write to CACHE_Partner_Landscape
  const sheetLandscape = ss.getSheetByName(SHEETS.CACHE_PARTNER_LANDSCAPE);
  sheetLandscape.clear(); // Complete overwrite
  
  if (finalRows.length > 0) {
    sheetLandscape.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#4285f4")
      .setFontColor("white")
      .setFontWeight("bold");
      
    sheetLandscape.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
    
    // Format Checkboxes (Managed + Countries)
    // Managed is Col 3. Countries start at Col 6.
    // Let's just blindly format known boolean columns.
    
    // Col 3 (Managed)
    sheetLandscape.getRange(2, 3, finalRows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    
    // Cols 6 to End (Countries)
    sheetLandscape.getRange(2, 6, finalRows.length, countryList.length).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    
    sheetLandscape.setFrozenRows(1);
    sheetLandscape.setFrozenColumns(2);
  }
  
  Logger.log(`[Landscape] Rebuilt with ${finalRows.length} partners.`);
}
