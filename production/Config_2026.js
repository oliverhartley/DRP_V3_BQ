/**
 * ****************************************
 * Google Apps Script - Global Configuration
 * File: Config.gs
 * Description: Centralized variables used across all scripts.
 * ****************************************
 */

const PROJECT_ID = "concord-prod"; 

// The Spreadsheet where the scripts run (Destination)
const DESTINATION_SS_ID = "1i_C2AdhnxqPqEAQrr3thhJGK_3cwXJhRrmnXQpPvtD0"; 

// The Master Source Spreadsheet (Where the original domains come from)
const SOURCE_SS_ID = "1XUVbK_VsV-9SsUzfp8YwUF2zJr3rMQ1ANJyQWdtagos"; // <--- THIS WAS MISSING

// Capacity Gap Spreadsheet
const CAPACITY_GAP_SS_ID = "15iyKfWZmce97cnxlZeryxF9ASalWbfpRjqND70Yp7Kw";

// The Folder where individual decks are saved
const PARTNER_FOLDER_ID = "1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC"; 

// Sheet Names
const SHEET_NAME_DB = "LATAM_Partner_DB";
const SHEET_NAME_SCORE = "LATAM_Partner_Score_DRP";
const SHEET_NAME_2026 = "LATAM_Partner_DB_2026";
const SHEET_NAME_SCORE_2026 = "LATAM_Partner_Score_2026";
const SHEET_NAME_DEEPDIVE_2026 = "LATAM_DeepDive_2026";
const SHEET_NAME_SOURCE = "Consolidate by Partner"; // Name of the tab in the SOURCE_SS_ID
const SHEET_NAME_DASHBOARD = "Partner / Region / Solution Selector";
const SHEET_NAME_DASHBOARD_2026 = "LATAM_Partner_Dashboard_2026";
const SHEET_NAME_CACHE_2026 = "CACHE_Dashboard_2026";
const SHEET_NAME_LINKS = "System_Link_Cache";

const PRODUCT_SCHEMA = [
  { solution: 'Infrastructure Modernization', color: '#fce5cd', products: ['Google Compute Engine', 'Google Cloud Networking', 'SAP on Google Cloud', 'Google Cloud VMware Engine', 'Google Distributed Cloud'] },
  { solution: 'Application Modernization', color: '#d9d2e9', products: ['Google Kubernetes Engine', 'Apigee API Management'] },
  { solution: 'Databases', color: '#fce5cd', products: ['Cloud SQL', 'AlloyDB for PostgreSQL', 'Spanner', 'Cloud Run', 'Oracle'] },
  { solution: 'Data & Analytics', color: '#d9ead3', products: ['BigQuery', 'Looker', 'Dataflow', 'Dataproc'] },
  { solution: 'Artificial Intelligence', color: '#c9daf8', products: ['Vertex AI Platform', 'AI Applications', 'Gemini Enterprise', 'Customer Engagement Suite'] },
  { solution: 'Security', color: '#f4cccc', products: ['Cloud Security', 'Security Command Center', 'Security Operations', 'Google Threat Intelligence'] },
  { solution: 'Workspace', color: '#fff2cc', products: ['Workspace'] }
];