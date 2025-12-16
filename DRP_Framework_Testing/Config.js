/**
 * ****************************************
 * Google Apps Script - Global Configuration
 * File: Config.gs
 * Description: Centralized variables used across all scripts.
 * ****************************************
 */

const PROJECT_ID = "concord-prod";

// The Spreadsheet where the scripts run (Destination)
const DESTINATION_SS_ID = "1MVVcY_mXlz5mn_6UqptWJRpm08Sd5odS5qy5NOfse_o";

// The Master Source Spreadsheet (Where the original domains come from)
const SOURCE_SS_ID = "1MVVcY_mXlz5mn_6UqptWJRpm08Sd5odS5qy5NOfse_o";

// Capacity Gap Spreadsheet
const CAPACITY_GAP_SS_ID = "15iyKfWZmce97cnxlZeryxF9ASalWbfpRjqND70Yp7Kw";

// The Folder where individual decks are saved
const PARTNER_FOLDER_ID = "1d8b3BxpFl79BoeriES2MSsS1GAmY1STq";

// Sheet Names
const SHEET_NAME_DB = "LATAM_Partner_DB";
const SHEET_NAME_SCORE = "LATAM_Partner_Score_DRP";
const SHEET_NAME_SOURCE = "Local_Partner_DB"; // Points to the new local sheet
const SHEET_NAME_DASHBOARD = "Partner / Region / Solution Selector";
const SHEET_NAME_LINKS = "System_Link_Cache";