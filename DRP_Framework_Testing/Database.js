/**
 * ****************************************
 * Google Apps Script - Database & Persistence
 * File: Database.js
 * Description: Management of DB_Partners (Managed), DB_Reference, and Migration.
 * ****************************************
 */

/**
 * Initializes the full database structure:
 * 1. DB_Partners (Target for Managed Partners)
 * 2. DB_Reference (Product/Solution Mappings)
 */
function initSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Initialize DB_Partners
  let sheetPartners = ss.getSheetByName(SHEETS.DB_PARTNERS);
  if (!sheetPartners) {
    sheetPartners = ss.insertSheet(SHEETS.DB_PARTNERS);
    sheetPartners.setTabColor("16a765"); // Green for "Source"
  }

  const headersPartners = [
    "Partner Name", "Domain",
    "Managed Partner", // Explicit Flag
    "Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Costa Rica", "Cuba",
    "Dominican Republic", "Ecuador", "El Salvador", "Guatemala", "Honduras",
    "Mexico", "Nicaragua", "Panama", "Paraguay", "Peru", "Uruguay", "Venezuela",
    "MCO", "GSI", "PS",
    "Google Compute Engine", "Google Cloud Networking", "SAP on Google Cloud",
    "Google Cloud VMware Engine", "Google Distributed Cloud",
    "Google Kubernetes Engine", "Apigee API Management",
    "Cloud SQL", "AlloyDB for PostgreSQL", "Spanner", "Cloud Run", "Oracle",
    "BigQuery", "Looker", "Dataflow", "Dataproc",
    "Vertex AI Platform", "AI Applications", "Gemini Enterprise", "Customer Engagement Suite",
    "Cloud Security", "Security Command Center", "Security Operations", "Google Threat Intelligence",
    "Workspace",
    "Email To", "Email CC"
  ];

  // Only set headers if empty to avoid overwriting user data
  if (sheetPartners.getLastRow() === 0) {
    sheetPartners.getRange(1, 1, 1, headersPartners.length)
      .setValues([headersPartners])
      .setBackground("#16a765")
      .setFontColor("white")
      .setFontWeight("bold");
    sheetPartners.setFrozenRows(1);
    sheetPartners.setFrozenColumns(2);
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
    // Populate Default Reference Data? (Can be done later)
  }

  Logger.log("System initialized: DB_Partners and DB_Reference ready.");
}

/**
 * Migrates data from Master Source to DB_Partners.
 * Only runs if DB_Partners is empty or explicit reset requested.
 */
function runMigration() {
  initSystem(); // Ensure sheets exist

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(SHEETS.DB_PARTNERS);

  // Safety check: Don't overwrite if data exists (unless we add a force flag later)
  if (targetSheet.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Migration Warning', 'DB_Partners already has data. This will APPEND data. Continue?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  const masterSS = SpreadsheetApp.openById(MASTER_SOURCE_SS_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_SHEET_NAME);
  const data = masterSheet.getDataRange().getValues();

  const results = [];

  // Skip header, map rows
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partnerName = row[33]; // AH
    const domain = row[34];      // AI

    if (!partnerName || !domain) continue;

    const newRow = new Array(53).fill(false); // 53 cols match header count
    newRow[0] = partnerName;
    newRow[1] = domain;
    newRow[2] = true; // Managed Partner = TRUE (by default from Master)

    // Country Logic (Indices sifted by +1 due to new Managed Flag)
    // Legacy Mapping Logic
    const regions = String(row[8] || "");
    const solutions = String(row[12] || "");

    // MCO/GSI/PS (Indices 22, 23, 24)
    if (regions.includes("MCO")) newRow[22] = true;
    if (regions.includes("GSI")) newRow[23] = true;
    if (regions.includes("PS")) newRow[24] = true;

    // Solution Aggregation mapping (Indices shifted +1 from previous V3)
    if (solutions.includes("Infra")) [25, 26, 27, 28, 29].forEach(idx => newRow[idx] = true);
    if (solutions.includes("App_Mod")) [30, 31].forEach(idx => newRow[idx] = true);
    if (solutions.includes("DB")) [32, 33, 34, 35, 36].forEach(idx => newRow[idx] = true);
    if (solutions.includes("Analytics")) [37, 38, 39, 40].forEach(idx => newRow[idx] = true);
    if (solutions.includes("AI_ML")) [41, 42, 43, 44].forEach(idx => newRow[idx] = true);
    if (solutions.includes("Security")) [45, 46, 47, 48].forEach(idx => newRow[idx] = true);
    if (solutions.includes("GWS")) newRow[49] = true;

    newRow[50] = row[35]; // AJ - Email To
    newRow[51] = row[36]; // AK - Email CC

    results.push(newRow);
  }

  if (results.length > 0) {
    // Write starting at next available row
    const startRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startRow, 1, results.length, results[0].length).setValues(results);
  }

  Logger.log(`Migration complete. ${results.length} partners imported into DB_Partners.`);
}
