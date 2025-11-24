
function testSimpleQuery() {
  const PROJECT_ID = 'concord-prod'; // Assuming this is correct from context
  const SQL_QUERY = `SELECT 1 as test_col`;
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    Logger.log("Test Query Success: " + JSON.stringify(queryResults.rows));
  } catch (e) {
    Logger.log("Test Query Failed: " + e.toString());
  }
}

function testVirtualTableQuery() {
  const PROJECT_ID = 'concord-prod';
  const VIRTUAL_TABLE_DATA = "STRUCT('@test.com' AS domain, true AS is_gsi)";
  const SQL_QUERY = `
    WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${VIRTUAL_TABLE_DATA} ]) )
    SELECT * FROM Spreadsheet_Data
  `;
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    Logger.log("Virtual Table Query Success: " + JSON.stringify(queryResults.rows));
  } catch (e) {
    Logger.log("Virtual Table Query Failed: " + e.toString());
  }
}

function testComplexQuery() {
  const PROJECT_ID = 'concord-prod';
  const SQL_QUERY = `
    WITH DummyData AS (
      SELECT 'p1' as partner_id, 'c1' as country, 'prof1' as profile_id
      UNION ALL SELECT 'p1', 'c1', 'prof2'
      UNION ALL SELECT 'p1', 'c2', 'prof3'
      UNION ALL SELECT 'p2', 'c1', 'prof4'
    ),
    ProfileBreakdown AS (
      SELECT 
        partner_id,
        STRING_AGG(CONCAT(country, ':', CAST(count AS STRING)), '|') as breakdown
      FROM (
        SELECT partner_id, country, COUNT(DISTINCT profile_id) as count
        FROM DummyData
        GROUP BY partner_id, country
      ) AS sub
      GROUP BY partner_id
    )
    SELECT * FROM ProfileBreakdown
  `;
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    Logger.log("Complex Query Success: " + JSON.stringify(queryResults.rows));
  } catch (e) {
    Logger.log("Complex Query Failed: " + e.toString());
  }
}

function testCteStructure() {
  const PROJECT_ID = 'concord-prod';
  const SQL_QUERY = `
    WITH DummyData AS (
      SELECT 'p1' as partner_id, 'c1' as country, 'prof1' as profile_id
      UNION ALL SELECT 'p1', 'c1', 'prof2'
      UNION ALL SELECT 'p1', 'c2', 'prof3'
      UNION ALL SELECT 'p2', 'c1', 'prof4'
    ),
    Prep AS (
      SELECT partner_id, country, COUNT(DISTINCT profile_id) as count
      FROM DummyData
      GROUP BY partner_id, country
    ),
    Breakdown AS (
      SELECT partner_id, STRING_AGG(CONCAT(country, ':', CAST(count AS STRING)), '|') as breakdown
      FROM Prep
      GROUP BY partner_id
    )
    SELECT * FROM Breakdown
  `;
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    Logger.log("CTE Structure Success: " + JSON.stringify(queryResults.rows));
  } catch (e) {
    Logger.log("CTE Structure Failed: " + e.toString());
  }
}

function testDomainMatching() {
  const PROJECT_ID = 'concord-prod';
  const SQL_QUERY = `
    WITH Spreadsheet_Data AS (
      SELECT * FROM UNNEST([
        STRUCT('accenture.com' AS domain),
        STRUCT('capgemini.com' AS domain),
        STRUCT('deloitte.com' AS domain)
      ])
    )
    SELECT 
      (SELECT COUNT(*) FROM \`concord-prod.service_partnercoe.drp_partner_master\`) as total_rows,
      t1.partner_name,
      bq_domain
    FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
    CROSS JOIN UNNEST(t1.partner_details.email_domain) AS bq_domain
    LIMIT 10
  `;
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    Logger.log("Domain Matching Success: " + JSON.stringify(queryResults.rows));
  } catch (e) {
    Logger.log("Domain Matching Failed: " + e.toString());
  }
}

function testSpreadsheetData() {
  const SOURCE_SS_ID = "1XUVbK_VsV-9SsUzfp8YwUF2zJr3rMQ1ANJyQWdtagos";
  const SHEET_NAME_SOURCE = "Consolidate by Partner";
  const DOMAIN_START_ROW = 3;
  const COL_MAP_DOMAIN = 33; // Column AH

  try {
    const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_SOURCE);
    if (!sheet) {
      Logger.log("Error: Sheet not found: " + SHEET_NAME_SOURCE);
      return;
    }
    const lastRow = sheet.getLastRow();
    Logger.log("Last Row: " + lastRow);
    if (lastRow < DOMAIN_START_ROW) {
      Logger.log("No data rows.");
      return;
    }
    const range = sheet.getRange(DOMAIN_START_ROW, 1, Math.min(10, lastRow - DOMAIN_START_ROW + 1), 35);
    const values = range.getValues();
    Logger.log("Read " + values.length + " rows.");
    for (let i = 0; i < values.length; i++) {
      const domain = values[i][COL_MAP_DOMAIN];
      Logger.log(`Row ${i + DOMAIN_START_ROW}: Domain='${domain}'`);
    }
  } catch (e) {
    Logger.log("Error reading spreadsheet: " + e.toString());
  }
}
