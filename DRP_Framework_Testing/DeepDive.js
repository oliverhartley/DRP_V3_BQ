/**
 * ****************************************
 * Google Apps Script - Deep Dive Profiles
 * File: DeepDive.js
 * Description: BigQuery Profile Deep Dive Loader for CACHE_DeepDive.
 * ****************************************
 */

/**
 * Main execution function for the Deep Dive Loader.
 */
function runDeepDiveLoader() {
  try {
    const virtualTableData = getScoringSqlStruct(); // Reuse structure from Scoring (Domain list)
    if (!virtualTableData) return;

    const sql = getDeepDiveSql(virtualTableData);
    const data = executeBigQuery(sql); // Reusing generic executor
    
    if (data) {
      persistToSheet(data, SHEETS.CACHE_DEEPDIVE);
    }
  } catch (e) {
    Logger.log(`[DeepDive] ERROR: ${e.toString()}`);
  }
}

/**
 * Builds the SQL for the Deep Dive Profile cache.
 * Fetches granular profile data for all Managed Partners.
 */
function getDeepDiveSql(virtualTableData) {
  return `
    WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${virtualTableData} ]) )
    SELECT
        t1.partner_name,
        t1.partner_id,
        t1.profile_details.profile_id,
        t1.profile_details.residing_country,
        p.name AS specialization_name,
        p.category AS specialization_category,
        p.status AS specialization_status
    FROM
        \`${PROJECT_ID}.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.profile_details.specializations) AS p
    JOIN Spreadsheet_Data AS sheet
      ON REGEXP_REPLACE(TRIM(LOWER(t1.partner_details.email_domain[OFFSET(0)])), r'^@', '') = REGEXP_REPLACE(TRIM(LOWER(sheet.domain)), r'^@', '')
    WHERE
        t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
  `;
}
