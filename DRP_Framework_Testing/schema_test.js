function testScoreHistorySchema() {
  const PROJECT_ID = "concord-prod";
  const SQL_QUERY = `
    SELECT 
      h.name, h.type, h.score, h.update_date 
    FROM \`concord-prod.service_partnercoe.drp_partner_master\` t,
    UNNEST(t.profile_details.score_history_details) h
    LIMIT 10
  `;
  
  const request = { query: SQL_QUERY, useLegacySql: false };
  try {
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    if (queryResults.rows) {
      Logger.log("Schema Test Results:");
      queryResults.rows.forEach(row => {
        Logger.log(row.f.map(f => f.v).join(", "));
      });
    } else {
      Logger.log("No rows returned.");
    }
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}
