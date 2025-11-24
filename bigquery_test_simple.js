
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
