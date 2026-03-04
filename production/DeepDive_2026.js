/**
 * ****************************************
 * Google Apps Script - Profile Deep Dive
 * File: DeepDive_2026.js
 * Description: 2026 Migration version of Deep Dive, utilizing Internal ID and row-level sub-region filtering.
 * ****************************************
 */

const DEEPDIVE_START_ROW_2026 = 2; // Data starts on row 2

function getDeepDiveSpreadsheetData2026() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_2026);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME_2026}" not found.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < DEEPDIVE_START_ROW_2026) return ""; 

  // Reading Columns A to F
  // A: Partner Name, B: Sub-Region, C: PDM, D: Type, E: Internal ID, F: Partner ID
  const range = sheet.getRange(DEEPDIVE_START_ROW_2026, 1, lastRow - DEEPDIVE_START_ROW_2026 + 1, 6);
  const values = range.getValues();

  let structList = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    let partnerName = String(row[0] || "").trim().replace(/[\x00-\x1F\x7F-\x9F\u200B]/g, "");
    let subRegion   = String(row[1] || "").trim();
    let pdm         = String(row[2] || "").trim();
    let type        = String(row[3] || "").trim();
    let internalId  = String(row[4] || "").trim();
    let partnerId   = String(row[5] || "").trim();

    if (internalId && partnerName) {
      structList.push(`STRUCT('${internalId}' AS internal_id, '${partnerId.replace(/'/g, "\\'")}' AS partner_id, '${partnerName.replace(/'/g, "\\'")}' AS partner_name, '${subRegion.replace(/'/g, "\\'")}' AS sub_region, '${pdm.replace(/'/g, "\\'")}' AS pdm, '${type.replace(/'/g, "\\'")}' AS partner_type)`);
    }
  }
  return structList.join(',\n');
}

function getDeepDiveSql2026(virtualTableData) {
  const sqlParts = [
    "WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([",
    `    ${virtualTableData}`,
    "  ]) ),",
    "RawProfileData AS (",
    "    SELECT",
    "        sheet.internal_id,",
    "        sheet.partner_id,",
    "        sheet.partner_name,",
    "        sheet.sub_region,",
    "        sheet.pdm,",
    "        sheet.partner_type,",
    "        t1.profile_details.profile_id,",
    "        t1.profile_details.residing_country,",
    "        t1.profile_details.job_title,",
    "        scores.scored_product,",
    "        scores.score,",
    "",
    "        CASE",
    "          WHEN scores.score >= 50 THEN 'Tier 1'",
    "          WHEN scores.score BETWEEN 35 AND 49 THEN 'Tier 2'",
    "          WHEN scores.score BETWEEN 20 AND 34 THEN 'Tier 3'",
    "          WHEN scores.score < 20 THEN 'Tier 4'",
    "          ELSE 'No Tier'",
    "        END AS practitioner_tier,",
    "",
    "        CASE",
    "          WHEN scores.scored_product IN ('Google Compute Engine', 'Google Cloud Networking', 'SAP on Google Cloud', 'Google Cloud VMware Engine', 'Google Distributed Cloud') THEN 'Infrastructure Modernization'",
    "          WHEN scores.scored_product IN ('Google Kubernetes Engine', 'Apigee API Management') THEN 'Application Modernization'",
    "          WHEN scores.scored_product IN ('Cloud SQL', 'AlloyDB for PostgreSQL', 'Spanner', 'Cloud Run', 'Oracle') THEN 'Databases'",
    "          WHEN scores.scored_product IN ('BigQuery', 'Looker', 'Dataflow', 'Dataproc') THEN 'Data & Analytics'",
    "          WHEN scores.scored_product IN ('Vertex AI Platform', 'AI Applications', 'Gemini Enterprise', 'Customer Engagement Suite') THEN 'Artificial Intelligence'",
    "          WHEN scores.scored_product IN ('Cloud Security', 'Security Command Center', 'Security Operations', 'Google Threat Intelligence') THEN 'Security'",
    "          WHEN scores.scored_product = 'Workspace' THEN 'Workspace'",
    "          ELSE 'Other'",
    "        END AS scored_solution",
    "",
    "    FROM",
    "        Spreadsheet_Data AS sheet",
    "    LEFT JOIN `concord-prod.service_partnercoe.drp_partner_master` AS t1",
    "        ON t1.partner_id = sheet.partner_id",
    "        AND (",
    "            (sheet.sub_region = 'Brazil' AND t1.profile_details.residing_country = 'Brazil') OR",
    "            (sheet.sub_region = 'Mexico' AND t1.profile_details.residing_country = 'Mexico') OR",
    "            (sheet.sub_region = 'MCO' AND t1.profile_details.residing_country NOT IN ('Brazil', 'Mexico') AND t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela'))",
    "        )",
    "    LEFT JOIN UNNEST(t1.profile_details.score_details) AS scores",
    "    WHERE scores.scored_product IS NOT NULL", // Ensure we only get actual scored profiles
    ")",
    "SELECT *",
    "FROM RawProfileData",
    "ORDER BY partner_name, profile_id, scored_solution, score DESC",
    "LIMIT 50000"
  ];
  
  return sqlParts.join('\n');
}

function runDeepDive2026() {
  try {
    Logger.log("1. Generating Virtual Table for 2026 Deep Dive...");
    const VIRTUAL_TABLE_DATA = getDeepDiveSpreadsheetData2026(); 
    if (!VIRTUAL_TABLE_DATA) return;

    Logger.log("2. Constructing Deep Dive SQL...");
    const SQL_QUERY = getDeepDiveSql2026(VIRTUAL_TABLE_DATA);

    Logger.log("3. Executing BigQuery Job...");
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    if (!queryResults.rows || queryResults.rows.length === 0) { 
      Logger.log("No data found."); 
      return; 
    }

    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_DEEPDIVE_2026);
    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_DEEPDIVE_2026); }
    sheet.clear();

    const headers = queryResults.schema.fields.map(f => f.name);
    // Properly format headers with nice capitalization
    const prettyHeaders = headers.map(h => {
        return h.split('_').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
    });

    const rows = queryResults.rows.map(row => row.f.map(cell => (cell.v === null) ? "" : cell.v));

    sheet.getRange(1, 1, 1, prettyHeaders.length).setValues([prettyHeaders]).setFontWeight("bold").setBackground("#d9d9d9");
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Freeze and format
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(6); // Freeze the metadata colums up to Partner Type

    // AutoResize columns A through M (Internal ID to Practitioner Tier)
    for (let c=1; c<=13; c++) {
      sheet.autoResizeColumn(c);
    }
    
    sheet.hideColumns(1, 2); // Hide Internal ID (A) and Partner ID (B)

    Logger.log(`SUCCESS! Loaded ${rows.length} rows to ${SHEET_NAME_DEEPDIVE_2026}.`);
  } catch (e) {
    console.error(`ERROR in runDeepDive2026: ${e.toString()}`);
  }
}
