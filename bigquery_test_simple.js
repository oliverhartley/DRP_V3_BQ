
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
