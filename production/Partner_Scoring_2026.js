/**
 * ****************************************
 * Google Apps Script - BigQuery Scored Partner Pivot
 * File: Partner_Scoring_2026.js
 * Description: 2026 migration version with Row-Level Granularity based on internal_id.
 * ****************************************
 */

const SCORING_START_ROW_2026 = 2; // Data starts on row 2

function getScoringSpreadsheetData2026() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_2026);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME_2026}" not found.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < SCORING_START_ROW_2026) return ""; 

  // Reading Columns A to F
  // A: Partner Name, B: Sub-Region, C: PDM, D: Type, E: Internal ID, F: Partner ID
  const range = sheet.getRange(SCORING_START_ROW_2026, 1, lastRow - SCORING_START_ROW_2026 + 1, 6);
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

function getScorePivotSql2026(virtualTableData) {
  const sqlParts = [
    "WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([",
    `    ${virtualTableData}`,
    "  ]) ),",
    "PivotData AS (",
    "    SELECT",
    "        t2.internal_id,",
    "        t2.partner_id,",
    "        t2.partner_name,",
    "        t2.sub_region,",
    "        t2.pdm,",
    "        t2.partner_type,",
    "        COUNT(DISTINCT t2.profile_id) AS Profiles_per_Region,",
    "",
    "        -- INFRASTRUCTURE MODERNIZATION",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_GCE_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_GCE_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_GCE_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_GCE_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_GCN_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_GCN_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_GCN_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_GCN_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_SAP_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_SAP_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_SAP_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_SAP_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_VME_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_VME_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_VME_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_VME_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_GDC_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_GDC_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_GDC_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_GDC_Tier4,",
    "        -- APPLICATION MODERNIZATION",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AppMod_GKE_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AppMod_GKE_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AppMod_GKE_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AppMod_GKE_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AppMod_Apigee_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AppMod_Apigee_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AppMod_Apigee_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AppMod_Apigee_Tier4,",
    "        -- DATABASES",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_CloudSQL_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_CloudSQL_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_CloudSQL_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_CloudSQL_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_AlloyDB_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_AlloyDB_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_AlloyDB_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_AlloyDB_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_Spanner_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_Spanner_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_Spanner_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_Spanner_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_CloudRun_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_CloudRun_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_CloudRun_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_CloudRun_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_Oracle_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_Oracle_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_Oracle_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_Oracle_Tier4,",
    "        -- DATA & ANALYTICS",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_BQ_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_BQ_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_BQ_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_BQ_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_Looker_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_Looker_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_Looker_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_Looker_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_Dataflow_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_Dataflow_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_Dataflow_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_Dataflow_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_Dataproc_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_Dataproc_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_Dataproc_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_Dataproc_Tier4,",
    "        -- ARTIFICIAL INTELLIGENCE",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_Vertex_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_Vertex_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_Vertex_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_Vertex_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_Applications_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_Applications_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_Applications_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_Applications_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_Gemini_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_Gemini_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_Gemini_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_Gemini_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_CES_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_CES_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_CES_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_CES_Tier4,",
    "        -- SECURITY",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_Cloud_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_Cloud_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_Cloud_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_Cloud_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_SCC_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_SCC_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_SCC_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_SCC_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_Ops_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_Ops_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_Ops_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_Ops_Tier4,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_GTI_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_GTI_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_GTI_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_GTI_Tier4,",
    "        -- WORKSPACE",
    "        COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS WS_Tier1,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS WS_Tier2,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS WS_Tier3,",
    "        COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS WS_Tier4"
  ];
  
  sqlParts.push("    FROM");
  sqlParts.push("        (");
  sqlParts.push("            SELECT");
  sqlParts.push("                sheet.internal_id,");
  sqlParts.push("                sheet.partner_id,");
  sqlParts.push("                sheet.partner_name,");
  sqlParts.push("                sheet.sub_region,");
  sqlParts.push("                sheet.pdm,");
  sqlParts.push("                sheet.partner_type,");
  sqlParts.push("                -- Pass profile_id to outer query for counting");
  sqlParts.push("                t1.profile_details.profile_id,");
  sqlParts.push("                scores.scored_product,");
  sqlParts.push("                CASE");
  sqlParts.push("                    WHEN scores.score >= 50 THEN 'Tier 1'");
  sqlParts.push("                    WHEN scores.score BETWEEN 35 AND 49 THEN 'Tier 2'");
  sqlParts.push("                    WHEN scores.score BETWEEN 20 AND 34 THEN 'Tier 3'");
  sqlParts.push("                    WHEN scores.score < 20 THEN 'Tier 4'");
  sqlParts.push("                    ELSE 'No Tier'");
  sqlParts.push("                END AS practitioner_tier,");
  sqlParts.push("                CASE");
  sqlParts.push("                    WHEN scores.scored_product IN ('Google Compute Engine', 'Google Cloud Networking', 'SAP on Google Cloud', 'Google Cloud VMware Engine', 'Google Distributed Cloud') THEN 'Infrastructure Modernization'");
  sqlParts.push("                    WHEN scores.scored_product IN ('Google Kubernetes Engine', 'Apigee API Management') THEN 'Application Modernization'");
  sqlParts.push("                    WHEN scores.scored_product IN ('Cloud SQL', 'AlloyDB for PostgreSQL', 'Spanner', 'Cloud Run', 'Oracle') THEN 'Databases'");
  sqlParts.push("                    WHEN scores.scored_product IN ('BigQuery', 'Looker', 'Dataflow', 'Dataproc') THEN 'Data & Analytics'");
  sqlParts.push("                    WHEN scores.scored_product IN ('Vertex AI Platform', 'AI Applications', 'Gemini Enterprise', 'Customer Engagement Suite') THEN 'Artificial Intelligence'");
  sqlParts.push("                    WHEN scores.scored_product IN ('Cloud Security', 'Security Command Center', 'Security Operations', 'Google Threat Intelligence') THEN 'Security'");
  sqlParts.push("                    WHEN scores.scored_product = 'Workspace' THEN 'Workspace'");
  sqlParts.push("                    ELSE 'Other' ");
  sqlParts.push("                END AS scored_solution");
  sqlParts.push("            FROM");
  sqlParts.push("                Spreadsheet_Data AS sheet");
  sqlParts.push("            LEFT JOIN `concord-prod.service_partnercoe.drp_partner_master` AS t1");
  sqlParts.push("              ON t1.partner_id = sheet.partner_id");
  sqlParts.push("              AND (");
  sqlParts.push("                  (sheet.sub_region = 'Brazil' AND t1.profile_details.residing_country = 'Brazil') OR");
  sqlParts.push("                  (sheet.sub_region = 'Mexico' AND t1.profile_details.residing_country = 'Mexico') OR");
  sqlParts.push("                  (sheet.sub_region = 'MCO' AND t1.profile_details.residing_country NOT IN ('Brazil', 'Mexico') AND t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela'))");
  sqlParts.push("              )");
  sqlParts.push("            LEFT JOIN UNNEST(t1.profile_details.score_details) AS scores");
  sqlParts.push("        ) AS t2");
  sqlParts.push("    GROUP BY");
  sqlParts.push("        t2.internal_id,");
  sqlParts.push("        t2.partner_id,");
  sqlParts.push("        t2.partner_name,");
  sqlParts.push("        t2.sub_region,");
  sqlParts.push("        t2.pdm,");
  sqlParts.push("        t2.partner_type");
  sqlParts.push(")");
  sqlParts.push("SELECT * FROM PivotData");
  
  return sqlParts.join('\n');
}

function formatScorePivotSheet2026(sheet) {
  if (!sheet) return;

  try {
    sheet.insertRowsBefore(1, 3);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // The order now is:
    // A: internal_id (can be hidden)
    // B: partner_id
    // C: partner_name
    // D: sub_region
    // E: pdm
    // F: partner_type
    // G: Profiles_per_Region

    // Freeze Cols and Rows
    sheet.setFrozenColumns(7); // Freeze up to Profiles_per_Region
    sheet.setFrozenRows(3);

    // Setup headers for Metadata columns
    sheet.getRange("A1").setValue("Internal ID");
    sheet.getRange("B1").setValue("Partner ID");
    sheet.getRange("C1").setValue("Partner Name");
    sheet.getRange("D1").setValue("Sub Region");
    sheet.getRange("E1").setValue("PDM");
    sheet.getRange("F1").setValue("Type of Partner");
    
    // FORMAT COLUMN G (Profiles per Region)
    sheet.getRange("G1").setValue("Profiles per Region").setBackground("#d9d9d9").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("G1:G3").merge(); // Merge the header rows 1-3 for this column
    sheet.setColumnWidth(7, 80);
    if (lastRow > 3) {
        sheet.getRange(4, 7, lastRow - 3, 1).setHorizontalAlignment("center");
    }

    const headerConfig = [
        {
            solution: 'Infrastructure Modernization',
            color: '#fce5cd',
            products: [
                { name: 'Google Compute Engine', span: 4, color: '#f4cccc' },
                { name: 'Google Cloud Networking', span: 4, color: '#fff2cc' },
                { name: 'SAP on Google Cloud', span: 4, color: '#d9ead3' },
                { name: 'Google Cloud VMware Engine', span: 4, color: '#d0e0e3' },
                { name: 'Google Distributed Cloud', span: 4, color: '#c9daf8' }
            ]
        },
        {
            solution: 'Application Modernization',
            color: '#d9d2e9',
            products: [
                { name: 'Google Kubernetes Engine', span: 4, color: '#f4cccc' },
                { name: 'Apigee API Management', span: 4, color: '#fff2cc' }
            ]
        },
        {
            solution: 'Databases',
            color: '#fce5cd',
            products: [
                { name: 'Cloud SQL', span: 4, color: '#f4cccc' },
                { name: 'AlloyDB for PostgreSQL', span: 4, color: '#fff2cc' },
                { name: 'Spanner', span: 4, color: '#d9ead3' },
                { name: 'Cloud Run', span: 4, color: '#d0e0e3' },
                { name: 'Oracle', span: 4, color: '#c9daf8' }
            ]
        },
        {
            solution: 'Data & Analytics',
            color: '#d9ead3',
            products: [
                { name: 'BigQuery', span: 4, color: '#f4cccc' },
                { name: 'Looker', span: 4, color: '#fff2cc' },
                { name: 'Dataflow', span: 4, color: '#d9ead3' },
                { name: 'Dataproc', span: 4, color: '#d0e0e3' }
            ]
        },
        {
            solution: 'Artificial Intelligence',
            color: '#c9daf8',
            products: [
                { name: 'Vertex AI Platform', span: 4, color: '#f4cccc' },
                { name: 'AI Applications', span: 4, color: '#fff2cc' },
                { name: 'Gemini Enterprise', span: 4, color: '#d9ead3' },
                { name: 'Customer Engagement Suite', span: 4, color: '#d0e0e3' }
            ]
        },
        {
            solution: 'Security',
            color: '#f4cccc',
            products: [
                { name: 'Cloud Security', span: 4, color: '#f4cccc' },
                { name: 'Security Command Center', span: 4, color: '#fff2cc' },
                { name: 'Security Operations', span: 4, color: '#d9ead3' },
                { name: 'Google Threat Intelligence', span: 4, color: '#d0e0e3' }
            ]
        },
        {
            solution: 'Workspace',
            color: '#fff2cc',
            products: [
                { name: 'Workspace', span: 4, color: '#f4cccc' }
            ]
        }
    ];

    // START FORMATTING AT COLUMN 8 (H)
    let currentCol = 8; 
    headerConfig.forEach(solutionGroup => {
        const solutionWidth = solutionGroup.products.reduce((acc, p) => acc + p.span, 0);
        if (solutionWidth > 0) {
            const solutionRange = sheet.getRange(1, currentCol, 1, solutionWidth);
            solutionRange.merge().setValue(solutionGroup.solution).setBackground(solutionGroup.color).setHorizontalAlignment('center').setFontWeight('bold');
        }
        solutionGroup.products.forEach(product => {
            if (product.span > 0) {
                const productRange = sheet.getRange(2, currentCol, 1, product.span);
                productRange.merge().setValue(product.name).setBackground(product.color).setHorizontalAlignment('center').setFontWeight('bold');
                const tierRange = sheet.getRange(3, currentCol, 1, product.span);
                const tierHeaders = [['Tier 1', 'Tier 2', 'Tier 3', 'Tier 4']];
                tierRange.setValues(tierHeaders).setHorizontalAlignment('center').setFontWeight('bold').setBackground(product.color);
                if (lastRow > 3) {
                    const dataRows = lastRow - 3; 
                    const dataRange = sheet.getRange(4, currentCol, dataRows, product.span);
                    dataRange.setBackground(product.color).setHorizontalAlignment('center'); 
                }
                // Setting widths correctly inside loop could be slow but keeping the original logic
                for(let i=0; i < product.span; i++){
                  sheet.setColumnWidth(currentCol + i, 50);
                }
                currentCol += product.span;
            }
        });
    });
    
    const headerBlock = sheet.getRange(1, 1, 3, lastCol);
    headerBlock.setBorder(true, true, true, true, true, true);
    
    // AutoResize columns A through F
    for (let c=1; c<=6; c++) {
      sheet.autoResizeColumn(c);
    }
    
  } catch (e) {
    console.error("Error during formatting: " + e.toString());
  }
}

function runPartnerScorePivot2026() {
  try {
    Logger.log("Generating virtual table for 2026 Scoring...");
    const VIRTUAL_TABLE_DATA = getScoringSpreadsheetData2026();
    if (!VIRTUAL_TABLE_DATA) { console.error("Error: No data."); return; }
    
    const SQL_QUERY = getScorePivotSql2026(VIRTUAL_TABLE_DATA);
    
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_SCORE_2026);
    if (!sheet) { 
      sheet = ss.insertSheet(SHEET_NAME_SCORE_2026); 
    } else { 
      sheet.clear(); 
    }
    
    Logger.log("Running 2026 Score Pivot query...");
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    if (!queryResults.rows || queryResults.rows.length === 0) { 
      sheet.getRange('A1').setValue("No results."); 
      return; 
    }

    const data = [];
    if (queryResults.rows) {
      queryResults.rows.forEach(row => {
        const rowData = row.f.map(field => field.v);
        data.push(rowData);
      });
    }
    
    // OVERWRITE ENTIRE SHEET
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
    
    formatScorePivotSheet2026(sheet);
    Logger.log("Success. Rows: " + data.length);
  } catch (e) { 
    console.error("CRITICAL ERROR: " + e.toString()); 
  }
}
