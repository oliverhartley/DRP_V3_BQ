/**
 * ****************************************
 * Google Apps Script - Source Data Migration
 * File: Migrate_Source_Data.gs
 * Description: One-time migration script to move external partner data to a local sheet.
 * ****************************************
 */

function migrateDataToLocal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSsId = "1XUVbK_VsV-9SsUzfp8YwUF2zJr3rMQ1ANJyQWdtagos";
  const sourceSheetName = "Consolidate by Partner";
  const localSheetName = "Local_Partner_DB";

  let localSheet = ss.getSheetByName(localSheetName);
  if (!localSheet) {
    localSheet = ss.insertSheet(localSheetName);
  } else {
    localSheet.clear();
  }

  const sourceSs = SpreadsheetApp.openById(sourceSsId);
  const sourceSheet = sourceSs.getSheetByName(sourceSheetName);
  if (!sourceSheet) throw new Error("Source sheet not found.");

  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const rows = sourceData.slice(2); // Assuming data starts on row 3

  // Define Target Schema
  const countries = [
    "Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Costa Rica", "Cuba", 
    "Dominican Republic", "Ecuador", "El Salvador", "Guatemala", "Honduras", 
    "Mexico", "Nicaragua", "Panama", "Paraguay", "Peru", "Uruguay", "Venezuela"
  ];
  const regions = ["MCO", "GSI", "PS"];
  const products = [
    // Infra
    "Google Compute Engine", "Google Cloud Networking", "SAP on Google Cloud", "Google Cloud VMware Engine", "Google Distributed Cloud",
    // AppMod
    "Google Kubernetes Engine", "Apigee API Management",
    // DB
    "Cloud SQL", "AlloyDB for PostgreSQL", "Spanner", "Cloud Run", "Oracle",
    // Data
    "BigQuery", "Looker", "Dataflow", "Dataproc",
    // AI
    "Vertex AI Platform", "AI Applications", "Gemini Enterprise", "Customer Engagement Suite",
    // Security
    "Cloud Security", "Security Command Center", "Security Operations", "Google Threat Intelligence",
    // Workspace
    "Workspace"
  ];

  const headers = [
    "Partner Name", "Domain",
    ...countries,
    ...regions,
    ...products,
    "Email To (AJ)", "Email CC (AK)"
  ];

  const output = [headers];

  // Mapping from Solution to Products
  const solutionToProducts = {
    "GSI": ["GSI"],
    "Brazil": ["Brazil"],
    "MCO": ["MCO"],
    "Mexico": ["Mexico"], // Map Mexico to MCO logic or keep separate? Usually Mexico is its own but user said regions: MCO, GSI, PS. 
    "PS": ["PS"],
    "AI_ML": ["Vertex AI Platform", "AI Applications", "Gemini Enterprise", "Customer Engagement Suite"],
    "GWS": ["Workspace"],
    "SECURITY": ["Cloud Security", "Security Command Center", "Security Operations", "Google Threat Intelligence"],
    "DB": ["Cloud SQL", "AlloyDB for PostgreSQL", "Spanner", "Cloud Run", "Oracle"],
    "ANALYTICS": ["BigQuery", "Looker", "Dataflow", "Dataproc"],
    "INFRA": ["Google Compute Engine", "Google Cloud Networking", "SAP on Google Cloud", "Google Cloud VMware Engine", "Google Distributed Cloud"],
    "APP_MOD": ["Google Kubernetes Engine", "Apigee API Management"]
  };

  const oldColMap = {
    PARTNER_NAME: 33, DOMAIN: 34,
    GSI: 7, BRAZIL: 8, MCO: 9, MEXICO: 10, PS: 11,
    AI_ML: 13, GWS: 14, SECURITY: 15, DB: 16, ANALYTICS: 17, INFRA: 18, APP_MOD: 19,
    EMAIL_TO: 35, EMAIL_CC: 36 // AJ is 35, AK is 36 (0-indexed)
  };

  rows.forEach(row => {
    let domain = String(row[oldColMap.DOMAIN] || "").trim();
    if (!domain || domain.includes("#N/A")) return;

    const rowOut = [];
    rowOut.push(row[oldColMap.PARTNER_NAME], domain);

    // Countries (Initialize all to false, check BQ match later or use existing flags)
    // For now, we'll mark the specific region country as true
    countries.forEach(c => {
      let isCountry = false;
      if (c === "Brazil" && row[oldColMap.BRAZIL] === true) isCountry = true;
      if (c === "Mexico" && row[oldColMap.MEXICO] === true) isCountry = true;
      // MCO countries are harder to map from old flags without deeper logic
      rowOut.push(isCountry);
    });

    // Regions
    rowOut.push(row[oldColMap.MCO] === true, row[oldColMap.GSI] === true, row[oldColMap.PS] === true);

    // Products (Map old Solution flags to all Products in that solution)
    products.forEach(p => {
      let active = false;
      for (const [sol, prodList] of Object.entries(solutionToProducts)) {
        if (prodList.includes(p) && row[oldColMap[sol]] === true) {
          active = true;
          break;
        }
      }
      rowOut.push(active);
    });

    // Metadata
    rowOut.push(row[oldColMap.EMAIL_TO], row[oldColMap.EMAIL_CC]);

    output.push(rowOut);
  });

  localSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  localSheet.setFrozenRows(1);
  localSheet.getRange(1, 1, 1, output[0].length).setBackground("#d9ead3").setFontWeight("bold");
  
  Logger.log("Migration complete. " + (output.length - 1) + " partners migrated.");
}
