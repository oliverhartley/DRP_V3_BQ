/**
 * ****************************************
 * Google Apps Script - Projects Exploration
 * File: Projects_Exploration.gs
 * Description: Dumps the raw structure of the 'project_details' nested array.
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs

const SHEET_NAME_PROJECTS_TEST = "TEST_Projects";

function runProjectDiscovery() {
  const targetPartner = "Xertica"; 

  try {
    Logger.log(`1. Starting Project Discovery for: ${targetPartner}`);

    // We construct a query that UNNESTS project_details
    // and selects STAR (*) to see all available columns inside it.
    const SQL_QUERY = `
      SELECT
        t1.partner_name,
        t1.profile_details.profile_id,
        t1.profile_details.residing_country,
        
        -- Select everything inside the nested project struct
        proj.* 

      FROM
        \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.profile_details.project_details) AS proj
      
      WHERE
        t1.partner_name = '${targetPartner}'
        AND t1.profile_details.residing_country IS NOT NULL
      
      -- Order by profile so we can see if one person has multiple projects
      ORDER BY 
        t1.profile_details.profile_id, 
        proj.customer_name
      LIMIT 1000
    `;

    Logger.log("2. Executing BigQuery Job...");
    const request = {
      query: SQL_QUERY,
      useLegacySql: false
    };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    if (!queryResults.rows || queryResults.rows.length === 0) {
      Logger.log("No project data found for " + targetPartner);
      return;
    }

    Logger.log(`3. Found ${queryResults.rows.length} projects. Writing to Sheet...`);
    
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_PROJECTS_TEST);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_PROJECTS_TEST);
    }
    sheet.clear();

    // Parse and Write Headers & Data
    const headers = queryResults.schema.fields.map(f => f.name);
    const rows = queryResults.rows.map(row => {
      return row.f.map(cell => (cell.v === null) ? "" : cell.v);
    });

    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setBackground("#d9ead3") // Light Green to distinguish from other tests
         .setFontWeight("bold");
         
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Auto-resize for readability
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log("SUCCESS! Check tab: " + SHEET_NAME_PROJECTS_TEST);

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    Browser.msgBox("Error", e.toString(), Browser.Buttons.OK);
  }
}