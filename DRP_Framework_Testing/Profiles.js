/**
 * ****************************************
 * Google Apps Script - Profiles DB
 * File: Profiles.js
 * Description: Generates a flat "Profile DB" for all LATAM partners.
 * ****************************************
 */

/**
 * Main execution function for the Profiles DB Loader.
 */
function runProfilesLoader() {
  try {
    const sql = getProfilesSql(); // No virtual table needed, we want ALL LATAM
    const data = executeBigQuery(sql); // Reusing generic executor form BigQuery_Core.js
    
    if (data) {
      persistToSheet(data, SHEETS.CACHE_PROFILES);
      formatProfilesSheet(SHEETS.CACHE_PROFILES);
    }
  } catch (e) {
    Logger.log(`[Profiles] ERROR: ${e.toString()}`);
  }
}

/**
 * Builds the SQL for the Profiles DB.
 * Fetches: Name, ProfileID, Country, JobTitle, Product, Score, Solution
 */
function getProfilesSql() {
  return `
            ELSE 'Other'
        END AS scored_solution
    */
    SELECT
        t1.partner_name,
        t1.profile_details.profile_id,
        t1.profile_details.residing_country
    FROM
        \`${PROJECT_ID}.service_partnercoe.drp_partner_master\` AS t1
    WHERE
        t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
    ORDER BY 1, 3, 2
  `;
}

/**
 * Applies formatting to the Profiles sheet.
 */
function formatProfilesSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground("#fbbc04") // Yellow for "Data"
    .setFontColor("black")
    .setFontWeight("bold");
    
  Logger.log("[Profiles] Formatting complete.");
}
