/**
 * ****************************************
 * Google Apps Script - Profile Deep Dive (SQL Source)
 * File: Profile_DeepDive.gs
 * Version: 2.0 (Fetch ALL Data for Batch Processing)
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs
const SHEET_NAME_DEEPDIVE_SOURCE = "TEST_DeepDive_Data";

function runDeepDiveQuerySource() {
  Logger.log(`1. Generating Virtual Table for ALL Partners...`);
  const VIRTUAL_TABLE_DATA = getScoringSpreadsheetData(); 
  if (!VIRTUAL_TABLE_DATA) return;

  Logger.log("2. Constructing Deep Dive SQL (Batch Mode)...");

  const SQL_QUERY = `
    WITH Spreadsheet_Data AS (
      SELECT * FROM UNNEST([ ${VIRTUAL_TABLE_DATA} ])
    ),
    RawProfileData AS (
      SELECT
        t1.partner_name,
        t1.profile_details.profile_id,
        t1.profile_details.residing_country,
        t1.profile_details.job_title,
        scores.scored_product,
        scores.score,
        
        CASE
          WHEN scores.score >= 50 THEN 'Tier 1'
          WHEN scores.score BETWEEN 35 AND 49 THEN 'Tier 2'
          WHEN scores.score BETWEEN 20 AND 34 THEN 'Tier 3'
          WHEN scores.score < 20 THEN 'Tier 4'
          ELSE 'No Tier'
        END AS practitioner_tier,

        CASE
          WHEN scores.scored_product IN ('Google Compute Engine', 'Google Cloud Networking', 'SAP on Google Cloud', 'Google Cloud VMware Engine', 'Google Distributed Cloud') THEN 'Infrastructure Modernization'
          WHEN scores.scored_product IN ('Google Kubernetes Engine', 'Apigee API Management') THEN 'Application Modernization'
          WHEN scores.scored_product IN ('Cloud SQL', 'AlloyDB for PostgreSQL', 'Spanner', 'Cloud Run', 'Oracle') THEN 'Databases'
          WHEN scores.scored_product IN ('BigQuery', 'Looker', 'Dataflow', 'Dataproc') THEN 'Data & Analytics'
          WHEN scores.scored_product IN ('Vertex AI Platform', 'AI Applications', 'Gemini Enterprise', 'Customer Engagement Suite') THEN 'Artificial Intelligence'
          WHEN scores.scored_product IN ('Cloud Security', 'Security Command Center', 'Security Operations', 'Google Threat Intelligence') THEN 'Security'
          WHEN scores.scored_product = 'Workspace' THEN 'Workspace'
          ELSE 'Other'
        END AS scored_solution

      FROM
        \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.profile_details.score_details) AS scores
      INNER JOIN Spreadsheet_Data AS sheet
        ON TRIM(LOWER(t1.partner_details.email_domain[OFFSET(0)])) = sheet.domain
      
      WHERE
        t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
        AND scores.scored_product IS NOT NULL
    )
    SELECT * FROM RawProfileData
    ORDER BY partner_name, profile_id, scored_solution, score DESC
    LIMIT 50000 
  `;

  Logger.log("3. Executing BigQuery Job...");
  const request = { query: SQL_QUERY, useLegacySql: false };
  const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

  if (!queryResults.rows || queryResults.rows.length === 0) { Logger.log("No data found."); return; }

  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  let sheet = ss.getSheetByName(SHEET_NAME_DEEPDIVE_SOURCE);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_DEEPDIVE_SOURCE); }
  sheet.clear();

  const headers = queryResults.schema.fields.map(f => f.name);
  const rows = queryResults.rows.map(row => row.f.map(cell => (cell.v === null) ? "" : cell.v));

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log(`SUCCESS! Loaded ${rows.length} rows.`);
}