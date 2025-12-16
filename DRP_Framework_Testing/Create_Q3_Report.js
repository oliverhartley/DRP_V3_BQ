/**
 * ****************************************
 * Google Apps Script - Q3 Report Generator
 * File: Create_Q3_Report.gs
 * Description: Generates a static snapshot of Partner Scores for Q3 (up to Sep 30).
 * ****************************************
 */

const SHEET_NAME_Q3_DATA = "LATAM_Partner_Score_DRP_Nov01";
const SHEET_NAME_Q3_DASHBOARD = "DRP Scores Nov 1";
const Q3_DATE_LIMIT = '2025-11-01'; // Updated to Nov 1

function runQ3Report() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  try { ss.toast("Generating Q3 Report...", "Started", 10); } catch (e) { }
  Logger.log("Generating Q3 Report...");

  // 1. Generate Q3 Data
  const success = generateQ3Data();
  if (!success) {
    try { ss.toast("Failed to generate Q3 Data. Check logs.", "Error", 20); } catch (e) { }
    Logger.log("Failed to generate Q3 Data.");
    return;
  }

  // 2. Create/Update Dashboard
  createQ3Dashboard();
  try { ss.toast("Q3 Report Ready!", "Success", 10); } catch (e) { }
  Logger.log("Q3 Report Ready!");
}

function generateQ3Data() {
  try {
    Logger.log("Generating Q3 Data...");

    // We need to replicate the logic from Partner_Scoring.gs but using score_history_details
    // and filtering by date.

    // First, get the virtual table data (same as normal scoring)
    const VIRTUAL_TABLE_DATA = getScoringSpreadsheetData(); // Reuse from Partner_Scoring.gs
    if (!VIRTUAL_TABLE_DATA) { console.error("Error: No virtual table data."); return false; }
    Logger.log("Virtual Table Data Length: " + VIRTUAL_TABLE_DATA.length);

    const SQL_QUERY = `
      WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([
          ${VIRTUAL_TABLE_DATA}
        ]) ),
      History AS (
        SELECT 
          t1.partner_id,
          t1.partner_name,
          t1.profile_details.profile_id,
          -- Map Security Foundation to Cloud Security to fix zeros
          CASE 
            WHEN TRIM(h.name) = 'Security Foundation' THEN 'Cloud Security'
            ELSE h.name 
          END AS scored_product,
          h.score,
          h.update_date
        FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.profile_details.score_history_details) AS h
        WHERE h.update_date <= '${Q3_DATE_LIMIT}'
      ),
      LatestHistory AS (
        SELECT * FROM History
        QUALIFY ROW_NUMBER() OVER (PARTITION BY partner_id, profile_id, scored_product ORDER BY update_date DESC) = 1
      ),
      PivotData AS (
          SELECT
              t2.partner_id,
              t2.partner_name,
              COUNT(DISTINCT t2.profile_id) AS Total_Profiles,

              -- INFRASTRUCTURE MODERNIZATION
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_GCE_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_GCE_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_GCE_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Compute Engine' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_GCE_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_GCN_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_GCN_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_GCN_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud Networking' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_GCN_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_SAP_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_SAP_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_SAP_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'SAP on Google Cloud' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_SAP_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_VME_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_VME_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_VME_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Cloud VMware Engine' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_VME_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Infra_GDC_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Infra_GDC_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Infra_GDC_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Infrastructure Modernization' AND t2.scored_product = 'Google Distributed Cloud' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Infra_GDC_Tier4,

              -- APPLICATION MODERNIZATION
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AppMod_GKE_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AppMod_GKE_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AppMod_GKE_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Google Kubernetes Engine' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AppMod_GKE_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AppMod_Apigee_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AppMod_Apigee_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AppMod_Apigee_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Application Modernization' AND t2.scored_product = 'Apigee API Management' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AppMod_Apigee_Tier4,

              -- DATABASES
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_CloudSQL_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_CloudSQL_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_CloudSQL_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud SQL' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_CloudSQL_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_AlloyDB_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_AlloyDB_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_AlloyDB_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'AlloyDB for PostgreSQL' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_AlloyDB_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_Spanner_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_Spanner_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_Spanner_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Spanner' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_Spanner_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_CloudRun_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_CloudRun_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_CloudRun_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Cloud Run' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_CloudRun_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DB_Oracle_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DB_Oracle_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DB_Oracle_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Databases' AND t2.scored_product = 'Oracle' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DB_Oracle_Tier4,

              -- DATA & ANALYTICS
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_BQ_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_BQ_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_BQ_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'BigQuery' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_BQ_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_Looker_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_Looker_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_Looker_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Looker' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_Looker_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_Dataflow_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_Dataflow_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_Dataflow_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataflow' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_Dataflow_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS DA_Dataproc_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS DA_Dataproc_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS DA_Dataproc_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Data & Analytics' AND t2.scored_product = 'Dataproc' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS DA_Dataproc_Tier4,

              -- ARTIFICIAL INTELLIGENCE
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_Vertex_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_Vertex_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_Vertex_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Vertex AI Platform' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_Vertex_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_Applications_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_Applications_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_Applications_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'AI Applications' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_Applications_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_Gemini_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_Gemini_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_Gemini_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Gemini Enterprise' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_Gemini_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS AI_CES_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS AI_CES_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS AI_CES_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Artificial Intelligence' AND t2.scored_product = 'Customer Engagement Suite' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS AI_CES_Tier4,

              -- SECURITY
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_Cloud_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_Cloud_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_Cloud_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Cloud Security' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_Cloud_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_SCC_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_SCC_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_SCC_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Command Center' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_SCC_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_Ops_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_Ops_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_Ops_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Security Operations' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_Ops_Tier4,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS Security_GTI_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS Security_GTI_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS Security_GTI_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Security' AND t2.scored_product = 'Google Threat Intelligence' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS Security_GTI_Tier4,

              -- WORKSPACE
              COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 1' THEN 1 END) AS WS_Tier1,
              COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 2' THEN 1 END) AS WS_Tier2,
              COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 3' THEN 1 END) AS WS_Tier3,
              COUNT(CASE WHEN t2.scored_solution = 'Workspace' AND t2.scored_product = 'Workspace' AND t2.practitioner_tier = 'Tier 4' THEN 1 END) AS WS_Tier4

          FROM
              (
                  SELECT
                      t1.partner_id,
                      t1.partner_name,
                      scores.profile_id,
                      scores.scored_product,
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
                      \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
                  JOIN LatestHistory AS scores 
                    ON t1.partner_id = scores.partner_id 
                    AND t1.profile_details.profile_id = scores.profile_id
                  LEFT JOIN Spreadsheet_Data AS sheet
                    ON TRIM(LOWER(t1.partner_details.email_domain[SAFE_OFFSET(0)])) = sheet.domain
                  WHERE
                      t1.profile_details.residing_country IN (
                          'Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica',
                          'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala',
                          'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru',
                          'Uruguay', 'Venezuela'
                      )
              ) AS t2
          GROUP BY
              t2.partner_id,
              t2.partner_name
      )
      SELECT * FROM PivotData
    `;

    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    // Save to Sheet
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_Q3_DATA);
    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_Q3_DATA); } else { sheet.clear(); }

    if (!queryResults.rows || queryResults.rows.length === 0) {
      console.error("Error: Query returned no rows.");
      sheet.getRange('A1').setValue("No results.");
      return false;
    }

    const data = [];
    queryResults.rows.forEach(row => {
      data.push(row.f.map(field => field.v));
    });

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();

    // We reuse formatScorePivotSheet from Partner_Scoring.gs to ensure headers are correct
    // This is important because the dashboard logic relies on specific column positions and headers
    formatScorePivotSheet(sheet);

    return true;
  } catch (e) {
    console.error("Q3 Data Generation Error: " + e.toString());
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    try { ss.toast("Error: " + e.toString(), "Failed", 20); } catch (e2) { }
    return false;
  }
}

function debugQ3Diagnostics() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName("DEBUG_DIAGNOSTICS");
  if (!sheet) ss.insertSheet("DEBUG_DIAGNOSTICS");
  else sheet.clear();

  const queries = [
    {
      name: "1. Master Table Count",
      sql: `SELECT COUNT(*) FROM \`concord-prod.service_partnercoe.drp_partner_master\``
    },
    {
      name: "2. Master Valid Profile_IDs",
      sql: `SELECT COUNT(*) FROM \`concord-prod.service_partnercoe.drp_partner_master\` WHERE profile_details.profile_id IS NOT NULL`
    },
    {
      name: "3. LatestHistory Count (Known Good CTE)",
      sql: `
        WITH History AS (
          SELECT 
            t1.partner_id,
            t1.profile_details.profile_id,
            CASE 
              WHEN TRIM(h.name) = 'Security Foundation' THEN 'Cloud Security'
              ELSE h.name 
            END AS scored_product,
            h.update_date
          FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
          UNNEST(t1.profile_details.score_history_details) AS h
          WHERE h.update_date <= '${Q3_DATE_LIMIT}'
        ),
        LatestHistory AS (
          SELECT * FROM History
          QUALIFY ROW_NUMBER() OVER (PARTITION BY partner_id, profile_id, scored_product ORDER BY update_date DESC) = 1
        )
        SELECT COUNT(*) FROM LatestHistory
      `
    },
    {
      name: "4. Join Check (PartnerID + ProfileID)",
      sql: `
        WITH History AS (
          SELECT 
            t1.partner_id,
            t1.profile_details.profile_id,
            CASE 
              WHEN TRIM(h.name) = 'Security Foundation' THEN 'Cloud Security'
              ELSE h.name 
            END AS scored_product,
            h.update_date
          FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
          UNNEST(t1.profile_details.score_history_details) AS h
          WHERE h.update_date <= '${Q3_DATE_LIMIT}'
        ),
        LatestHistory AS (
          SELECT * FROM History
          QUALIFY ROW_NUMBER() OVER (PARTITION BY partner_id, profile_id, scored_product ORDER BY update_date DESC) = 1
        )
        SELECT COUNT(*) 
        FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
        JOIN LatestHistory AS scores 
          ON t1.partner_id = scores.partner_id 
          AND t1.profile_details.profile_id = scores.profile_id
      `
    },
    {
      name: "5. Pivot Check (With Fix)",
      sql: `
        WITH History AS (
          SELECT 
            t1.partner_id,
            t1.partner_name,
            t1.profile_details.profile_id,
            CASE 
              WHEN TRIM(h.name) = 'Security Foundation' THEN 'Cloud Security'
              ELSE h.name 
            END AS scored_product,
            h.score,
            h.update_date
          FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
          UNNEST(t1.profile_details.score_history_details) AS h
          WHERE h.update_date <= '${Q3_DATE_LIMIT}'
        ),
        LatestHistory AS (
          SELECT * FROM History
          QUALIFY ROW_NUMBER() OVER (PARTITION BY partner_id, profile_id, scored_product ORDER BY update_date DESC) = 1
        ),
        PivotData AS (
            SELECT
                t2.partner_id,
                t2.partner_name,
                COUNT(DISTINCT t2.profile_id) AS Total_Profiles
            FROM
                (
                    SELECT
                        t1.partner_id,
                        t1.partner_name,
                        scores.profile_id,
                        scores.scored_product
                    FROM
                        \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
                    JOIN LatestHistory AS scores 
                      ON t1.partner_id = scores.partner_id
                      AND t1.profile_details.profile_id = scores.profile_id
                    WHERE
                        t1.profile_details.residing_country IN (
                            'Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica',
                            'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala',
                            'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru',
                            'Uruguay', 'Venezuela'
                        )
                ) AS t2
            GROUP BY
                t2.partner_id,
                t2.partner_name
        )
        SELECT COUNT(*) as count FROM PivotData
      `
    }
  ];

  let row = 1;
  queries.forEach(q => {
    try {
      const request = { query: q.sql, useLegacySql: false };
      const result = BigQuery.Jobs.query(request, PROJECT_ID);
      const count = result.rows ? result.rows[0].f[0].v : "0";
      sheet.getRange(row, 1).setValue(q.name);
      sheet.getRange(row, 2).setValue(count);
      Logger.log(`${q.name}: ${count}`);
    } catch (e) {
      sheet.getRange(row, 1).setValue(q.name);
      sheet.getRange(row, 2).setValue("Error: " + e.toString());
      Logger.log(`${q.name} Error: ${e.toString()}`);
    }
    row++;
  });

  // 6. Test Full Pivot (No Spreadsheet Data)
  try {
    const pivotQuery = `
      WITH History AS (
        SELECT 
          t1.partner_id,
          t1.partner_name,
          t1.profile_details.profile_id,
          CASE 
            WHEN TRIM(h.name) = 'Security Foundation' THEN 'Cloud Security'
            ELSE h.name 
          END AS scored_product,
          h.score,
          h.update_date
        FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.profile_details.score_history_details) AS h
        WHERE h.update_date <= '${Q3_DATE_LIMIT}'
      ),
      LatestHistory AS (
        SELECT * FROM History
        QUALIFY ROW_NUMBER() OVER (PARTITION BY partner_id, profile_id, scored_product ORDER BY update_date DESC) = 1
      ),
      PivotData AS (
          SELECT
              t2.partner_id,
              t2.partner_name,
              COUNT(DISTINCT t2.profile_id) AS Total_Profiles
          FROM
              (
                  SELECT
                      t1.partner_id,
                      t1.partner_name,
                      scores.profile_id,
                      scores.scored_product
                  FROM
                      \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
                  JOIN LatestHistory AS scores 
                    ON t1.partner_id = scores.partner_id
                    AND t1.profile_details.profile_id = scores.profile_id
                  WHERE
                      t1.profile_details.residing_country IN (
                          'Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica',
                          'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala',
                          'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru',
                          'Uruguay', 'Venezuela'
                      )
              ) AS t2
          GROUP BY
              t2.partner_id,
              t2.partner_name
      )
      SELECT COUNT(*) as count FROM PivotData
    `;
    const request = { query: pivotQuery, useLegacySql: false };
    const result = BigQuery.Jobs.query(request, PROJECT_ID);
    const count = result.rows ? result.rows[0].f[0].v : "0";
    sheet.getRange(row, 1).setValue("6. Pivot Count (No Sheet Join)");
    sheet.getRange(row, 2).setValue(count);
    Logger.log(`6. Pivot Count (No Sheet Join): ${count}`);
  } catch (e) {
    sheet.getRange(row, 1).setValue("6. Pivot Count Error");
    sheet.getRange(row, 2).setValue(e.toString());
    Logger.log(`6. Pivot Count Error: ${e.toString()}`);
  }
  row++; // Increment row after the first new query block

  // 7. Test Full Pivot WITH Spreadsheet Data (The Real Deal)
  try {
    const VIRTUAL_TABLE_DATA = getScoringSpreadsheetData();
    if (!VIRTUAL_TABLE_DATA) {
      Logger.log("7. Pivot Count (With Sheet Join): Error - No Virtual Table Data");
      sheet.getRange(row, 1).setValue("7. Pivot Count (With Sheet Join)");
      sheet.getRange(row, 2).setValue("Error: No Virtual Table Data");
    } else {
      const fullQuery = `
          WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([
              ${VIRTUAL_TABLE_DATA}
            ]) ),
          History AS (
            SELECT 
              t1.partner_id,
              t1.partner_name,
              t1.profile_details.profile_id,
              CASE 
                WHEN TRIM(h.name) = 'Security Foundation' THEN 'Cloud Security'
                ELSE h.name 
              END AS scored_product,
              h.score,
              h.update_date
            FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
            UNNEST(t1.profile_details.score_history_details) AS h
            WHERE h.update_date <= '${Q3_DATE_LIMIT}'
          ),
          LatestHistory AS (
            SELECT * FROM History
            QUALIFY ROW_NUMBER() OVER (PARTITION BY partner_id, profile_id, scored_product ORDER BY update_date DESC) = 1
          ),
          PivotData AS (
              SELECT
                  t2.partner_id,
                  t2.partner_name,
                  COUNT(DISTINCT t2.profile_id) AS Total_Profiles
              FROM
                  (
                      SELECT
                          t1.partner_id,
                          t1.partner_name,
                          scores.profile_id,
                          scores.scored_product
                      FROM
                          \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
                      JOIN LatestHistory AS scores 
                        ON t1.partner_id = scores.partner_id
                        AND t1.profile_details.profile_id = scores.profile_id
                      LEFT JOIN Spreadsheet_Data AS sheet
                        ON TRIM(LOWER(t1.partner_details.email_domain[SAFE_OFFSET(0)])) = sheet.domain
                      WHERE
                          t1.profile_details.residing_country IN (
                              'Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica',
                              'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala',
                              'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru',
                              'Uruguay', 'Venezuela'
                          )
                  ) AS t2
              GROUP BY
                  t2.partner_id,
                  t2.partner_name
          )
          SELECT COUNT(*) as count FROM PivotData
        `;
      const request = { query: fullQuery, useLegacySql: false };
      const result = BigQuery.Jobs.query(request, PROJECT_ID);
      if (result.errorResult) {
        Logger.log(`7. Pivot Count (With Sheet Join): Error - ${JSON.stringify(result.errorResult)}`);
        sheet.getRange(row, 1).setValue("7. Pivot Count (With Sheet Join)");
        sheet.getRange(row, 2).setValue(`Error: ${result.errorResult.message}`);
      } else {
        const count = result.rows ? result.rows[0].f[0].v : "0";
        sheet.getRange(row, 1).setValue("7. Pivot Count (With Sheet Join)");
        sheet.getRange(row, 2).setValue(count);
        Logger.log(`7. Pivot Count (With Sheet Join): ${count}`);
      }
    }
  } catch (e) {
    sheet.getRange(row, 1).setValue("7. Pivot Count Error");
    sheet.getRange(row, 2).setValue(e.toString());
    Logger.log(`7. Pivot Count Error: ${e.toString()}`);
  }
  row++; // Increment row after the second new query block
}

function createQ3Dashboard() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);

  // 1. Copy Dashboard
  let dashboardSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!dashboardSheet) { console.error("Dashboard sheet not found"); return; }

  let q3Sheet = ss.getSheetByName(SHEET_NAME_Q3_DASHBOARD);
  if (q3Sheet) {
    // If it exists, we clear the data area but keep the sheet
    // Actually, easier to delete and recreate to ensure fresh copy of slicers/layout
    ss.deleteSheet(q3Sheet);
  }

  q3Sheet = dashboardSheet.copyTo(ss).setName(SHEET_NAME_Q3_DASHBOARD);

  // 2. Populate with Q3 Data
  // We need to run the dashboard logic but pointing to Q3 Data Sheet
  populateQ3DashboardData(q3Sheet);
}

function populateQ3DashboardData(dashSheet) {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const dbSheet = ss.getSheetByName(SHEET_NAME_DB);
  const scoreSheet = ss.getSheetByName(SHEET_NAME_Q3_DATA); // USE Q3 DATA

  if (!dbSheet || !scoreSheet) {
    dashSheet.getRange(DATA_START_ROW, 1).setValue("Error: DB or Q3 Score Sheets missing.");
    return;
  }

  // 1. Get Selection (Default to All for static report, or read from the copied sheet)
  // Since we copied the sheet, it has the values from the original dashboard.
  // We should probably reset them to "All" to show full report?
  // Or keep them as is?
  // The user said "populate with the numbers... so we can compare".
  // Usually reports are "All".
  // Let's reset to "All" for consistency.
  dashSheet.getRange(CELL_TYPE.r, CELL_TYPE.c).setValue("All");
  dashSheet.getRange(CELL_REGION.r, CELL_REGION.c).setValue("LATAM (All)");
  dashSheet.getRange(CELL_COUNTRY.r, CELL_COUNTRY.c).setValue("All");
  dashSheet.getRange(CELL_SOLUTION.r, CELL_SOLUTION.c).setValue("All");
  dashSheet.getRange(CELL_PRODUCT.r, CELL_PRODUCT.c).setValue("All");

  const typeSel = "All";
  const regionSel = "LATAM (All)";
  const countrySel = "All";
  const solutionSel = "All";
  const solutionSelArray = ["All"];
  const productSel = "All";

  // 2. Filter Partners (Logic from Partner_Region_Solution_Selector.gs)
  const dbData = dbSheet.getDataRange().getValues();
  if (dbData.length < 2) return;

  const dbHeaders = dbData[0];
  const partnerMap = new Map();

  const idxName = 1; const idxCountry = 3; const idxManaged = 5;
  const idxRegion = -1; // LATAM (All)

  const idxTotalProfiles = dbHeaders.indexOf("Total_Profiles");
  const idxProfileBreakdown = dbHeaders.indexOf("Profile_Breakdown");

  for (let i = 1; i < dbData.length; i++) {
    const pName = dbData[i][idxName];
    const pCountryString = dbData[i][idxCountry];
    const isManaged = dbData[i][idxManaged] === true;
    const isRegion = true; // LATAM (All)
    let countryArray = [];
    if (pCountryString) { countryArray = String(pCountryString).split(',').map(s => s.trim()); }

    const profileMap = new Map();
    let totalProfiles = 0;

    if (idxProfileBreakdown !== -1 && dbData[i][idxProfileBreakdown]) {
      const profileBreakdownStr = String(dbData[i][idxProfileBreakdown]);
      profileBreakdownStr.split('|').forEach(pair => {
        const [country, count] = pair.split(':');
        if (country && count) profileMap.set(country.trim(), parseInt(count));
      });
    }
    if (idxTotalProfiles !== -1) {
      totalProfiles = dbData[i][idxTotalProfiles] || 0;
    }

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
  if (scoreValues.length < 3) return;

  const scoreBackgrounds = scoreRange.getBackgrounds();
  const scoreFontWeights = scoreRange.getFontWeights();
  const rowSol = scoreValues[0];
  const rowProd = scoreValues[1];

  const columnsToKeep = [0, 1, 2];
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
    columnsToKeep.push(c); // Keep all for "All" selection
  }

  // 4. Build Output Rows
  let outputValues = [], outputBackgrounds = [], outputWeights = [];

  // Headers
  for (let r = 0; r < 3; r++) {
    let rowV = [], rowB = [], rowW = [];
    columnsToKeep.forEach(idx => {
      if (idx === 2) {
        if (r === 0) rowV.push("", "", "");
        else if (r === 1) rowV.push("", "", "");
        else if (r === 2) rowV.push("Total Profiles", "Region Profiles", "Country Profiles");
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
      keepRow = true; // Keep all for now since we selected All
    }
    if (keepRow) {
      let rowV = [], rowB = [], rowW = [];
      columnsToKeep.forEach(idx => {
        let val = scoreValues[r][idx];

        if (idx === 1) {
          const safeName = String(val).replace(/'/g, "''");
          const linkSheet = typeof SHEET_NAME_LINKS !== 'undefined' ? SHEET_NAME_LINKS : "System_Link_Cache";
          val = `=IFNA(HYPERLINK(VLOOKUP("${safeName}", ${linkSheet}!A:B, 2, FALSE), "${safeName}"), "${safeName}")`;
        }

        if (idx === 2) {
          const total = meta ? meta.totalProfiles : 0;
          rowV.push(total, total, total); // All same for "All" selection
          rowB.push(scoreBackgrounds[r][idx], scoreBackgrounds[r][idx], scoreBackgrounds[r][idx]);
          rowW.push(scoreFontWeights[r][idx], scoreFontWeights[r][idx], scoreFontWeights[r][idx]);
          return;
        }

        rowV.push(val);
        rowB.push(scoreBackgrounds[r][idx]);
        rowW.push(scoreFontWeights[r][idx]);
      });
      outputValues.push(rowV); outputBackgrounds.push(rowB); outputWeights.push(rowW);
    }
  }

  // 5. Sorting
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

    // Merging logic
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

function checkDataAvailabilityForDates() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  let sheet = ss.getSheetByName("DATA_AVAILABILITY_CHECK");
  if (!sheet) {
    sheet = ss.insertSheet("DATA_AVAILABILITY_CHECK");
  } else {
    sheet.clear();
  }

  const datesToCheck = ['2025-11-01', '2025-10-01'];
  let row = 1;

  sheet.getRange(row, 1, 1, 2).setValues([['Date', 'Row Count']]).setFontWeight('bold');
  row++;

  datesToCheck.forEach(date => {
    const tableName = '`concord-prod.service_partnercoe.drp_partner_master`';
    const sql = 'SELECT COUNT(*) as count FROM ' + tableName + ' AS t1, UNNEST(t1.profile_details.score_history_details) AS h WHERE h.update_date <= \'' + date + '\'';

    try {
      const request = { query: sql, useLegacySql: false };
      const result = BigQuery.Jobs.query(request, PROJECT_ID);
      const count = result.rows ? result.rows[0].f[0].v : "0";
      sheet.getRange(row, 1).setValue(date);
      sheet.getRange(row, 2).setValue(count);
      Logger.log('Data count for ' + date + ': ' + count);
    } catch (e) {
      sheet.getRange(row, 1).setValue(date);
      sheet.getRange(row, 2).setValue('Error: ' + e.toString());
      Logger.log('Error checking data for ' + date + ': ' + e.toString());
    }
    row++;
  });

  ss.toast('Data availability check complete. See DATA_AVAILABILITY_CHECK sheet.', 'Complete', 10);
}
