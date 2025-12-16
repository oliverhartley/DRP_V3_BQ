/**
 * ****************************************
 * Google Apps Script - BigQuery Playground
 * File: bigquery_tests.gs
 * Version: 1.4 - Cleaned (Uses Global Config)
 * ****************************************
 */

// NOTE: PROJECT_ID and DESTINATION_SS_ID come from Config.gs

const TEST_SHEET_NAME = "TEST_OUTPUT"; 

// Column Indices (0-based)
const COL_INDEX = {
  BRAZIL: 7,
  MCO: 8,
  MEXICO: 9,
  PS: 10,
  DOMAIN: 33
};

/**
 * 1. HELPER FUNCTION
 */
function getSpreadsheetDataAsSqlStruct_Test() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName("Consolidate by Partner");
  
  if (!sheet) {
    throw new Error('Sheet "Consolidate by Partner" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return ""; 

  const range = sheet.getRange(3, 1, lastRow - 2, 34); 
  const values = range.getValues();

  let structList = [];

  values.forEach(row => {
    let domain = String(row[COL_INDEX.DOMAIN]);
    domain = domain.toLowerCase().trim().replace(/[\x00-\x1F\x7F-\x9F\u200B]/g, "");

    if (domain && domain.includes('@')) {
      let isBrazil = row[COL_INDEX.BRAZIL] === true;
      let isMCO = row[COL_INDEX.MCO] === true;
      let isMexico = row[COL_INDEX.MEXICO] === true;
      let isPS = row[COL_INDEX.PS] === true;

      let sqlLine = `STRUCT('${domain}' AS domain, ${isBrazil} AS is_brazil, ${isMCO} AS is_mco, ${isMexico} AS is_mexico, ${isPS} AS is_ps)`;
      structList.push(sqlLine);
    }
  });

  return structList.join(',\n');
}

/**
 * 2. MAIN TEST FUNCTION
 */
function runTestQuery() {
  try {
    Logger.log("1. Building Virtual Table...");
    const VIRTUAL_TABLE_DATA = getSpreadsheetDataAsSqlStruct_Test(); // Use local helper
    
    if (!VIRTUAL_TABLE_DATA) {
      Logger.log("No data found.");
      return;
    }

    Logger.log("2. Constructing SQL...");

    const SQL_QUERY = `
      WITH Spreadsheet_Data AS (
        SELECT * FROM UNNEST([
          ${VIRTUAL_TABLE_DATA}
        ])
      ),
      JoinedData AS (
        SELECT
          bq.partner_id,
          bq.partner_name,
          LOGICAL_OR(sheet.is_brazil) AS Flag_Brazil,
          LOGICAL_OR(sheet.is_mco) AS Flag_MCO,
          LOGICAL_OR(sheet.is_mexico) AS Flag_Mexico,
          LOGICAL_OR(sheet.is_ps) AS Flag_PS,
          STRING_AGG(DISTINCT sheet.domain, ', ') AS Matched_Sheet_Domains,
          STRING_AGG(DISTINCT bq_domain, ', ') AS Matched_BQ_Domains
        FROM
          \`concord-prod.service_partnercoe.drp_partner_master\` AS bq,
          UNNEST(bq.partner_details.email_domain) AS bq_domain
        INNER JOIN Spreadsheet_Data AS sheet
          ON TRIM(LOWER(bq_domain)) = sheet.domain
        GROUP BY
          1, 2
      )
      SELECT * FROM JoinedData
      ORDER BY partner_name
      LIMIT 1000
    `;

    Logger.log("3. Executing BigQuery Job...");
    const request = {
      query: SQL_QUERY,
      useLegacySql: false
    };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID); // Uses Global

    if (!queryResults.rows || queryResults.rows.length === 0) {
      Logger.log("0 rows returned.");
      return;
    }

    Logger.log("4. Writing results...");
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(TEST_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(TEST_SHEET_NAME);
    }
    sheet.clear();

    const headers = queryResults.schema.fields.map(f => f.name);
    const rows = queryResults.rows.map(row => {
      return row.f.map(cell => (cell.v === null) ? "" : cell.v);
    });

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    Logger.log("SUCCESS!");

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    Browser.msgBox("Error", e.toString(), Browser.Buttons.OK);
  }
}
