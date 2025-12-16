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
const SHEET_NAME_SOURCE = "Consolidate by Partner"; // Name of the tab in the SOURCE_SS_ID
const SHEET_NAME_DASHBOARD = "Partner / Region / Solution Selector";
const SHEET_NAME_LINKS = "System_Link_Cache";