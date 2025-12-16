/**
 * ****************************************
 * Google Apps Script - Global Configuration
 * File: Config.js
 * Version: 6.0 (V3 Reset)
 * ****************************************
 */

// Deployment Environment
const ENV = 'TESTING'; // 'PRODUCTION' or 'TESTING'

// Project & Spreadsheet IDs
const PROJECT_ID = "concord-prod";

// Testing IDs
const DESTINATION_SS_ID = "1MVVcY_mXlz5mn_6UqptWJRpm08Sd5odS5qy5NOfse_o";
const SOURCE_SS_ID = "1MVVcY_mXlz5mn_6UqptWJRpm08Sd5odS5qy5NOfse_o"; // Self-contained

// Master Source (For Migration Only)
const MASTER_SOURCE_SS_ID = "1XUVbK_VsV-9SsUzfp8YwUF2zJr3rMQ1ANJyQWdtagos";
const MASTER_SHEET_NAME = "Consolidate by Partner";

// Capacity Gap & Folders
const CAPACITY_GAP_SS_ID = "15iyKfWZmce97cnxlZeryxF9ASalWbfpRjqND70Yp7Kw";
const PARTNER_FOLDER_ID = "1d8b3BxpFl79BoeriES2MSsS1GAmY1STq";

// Sheet Names (Internal)
const SHEETS = {
  // Managed Sources (User Edits These)
  DB_PARTNERS: "DB_Partners",
  DB_REFERENCE: "DB_Reference",

  // Automated Caches (Read-Only, Rebuilt Daily)
  CACHE_SCORING: "CACHE_Scoring",
  CACHE_DEEPDIVE: "CACHE_DeepDive",

  // Legacy / Temp definitions (to be phased out or mapped)
  SOURCE: "DB_Partners", // Mapping old 'SOURCE' to new 'DB_PARTNERS' for compatibility
};

// Column Mappings (Refined for Local_Partner_DB)
const COL_INDEX = {
  PARTNER_NAME: 0,
  DOMAIN: 1,
  COUNTRIES_START: 2,
  REGIONS_START: 21,
  PRODUCTS_START: 24,
  GWS_PRODUCT: 48,
  EMAIL_TO: 49,
  EMAIL_CC: 50,
  STATUS_BATCH: 51
};
