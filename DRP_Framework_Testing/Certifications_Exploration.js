/**
 * ****************************************
 * Google Apps Script - Certification Exploration
 * File: Certifications_Exploration.gs
 * Description: Dumps the raw structure of the 'certification_details' nested array.
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

const SHEET_NAME_CERTS_TEST = "TEST_Certifications";

function runCertificationDiscovery() {
  const targetPartner = "Xertica"; 

  try {
    Logger.log(`1. Starting Certification Discovery for: ${targetPartner}`);

    // We construct a query that simply UNNESTS certification_details
    // and selects STAR (*) to see all available columns inside it.
    const SQL_QUERY = `
      SELECT
        t1.partner_name,
        t1.profile_details.profile_id,
        t1.profile_details.residing_country,
        
        -- THIS IS THE KEY PART: 
        -- We select everything inside the nested certification struct
        cert.* 

      FROM
        \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.profile_details.certification_details) AS cert
      
      WHERE
        t1.partner_name = '${targetPartner}'
        AND t1.profile_details.residing_country IS NOT NULL
      
      -- Order by profile so we can see if one person has multiple certs
      ORDER BY 
        t1.profile_details.profile_id, 
        cert.name
      LIMIT 1000
    `;

    Logger.log("2. Executing BigQuery Job...");
    const request = {
      query: SQL_QUERY,
      useLegacySql: false
    };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    if (!queryResults.rows || queryResults.rows.length === 0) {
      Logger.log("No certification data found for " + targetPartner);
      return;
    }

    Logger.log(`3. Found ${queryResults.rows.length} certifications. Writing to Sheet...`);
    
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_CERTS_TEST);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_CERTS_TEST);
    }
    sheet.clear();

    // Parse and Write Headers & Data
    const headers = queryResults.schema.fields.map(f => f.name);
    const rows = queryResults.rows.map(row => {
      return row.f.map(cell => (cell.v === null) ? "" : cell.v);
    });

    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setBackground("#fff2cc") // Light Yellow to distinguish from other tests
         .setFontWeight("bold");
         
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Auto-resize for readability
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log("SUCCESS! Check tab: " + SHEET_NAME_CERTS_TEST);

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    Browser.msgBox("Error", e.toString(), Browser.Buttons.OK);
  }
}
