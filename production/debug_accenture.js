function debugOrionInBQ() {
  const keyword = "Orion";
  const domain = "orion2000.com";
  
  const sql = `
    SELECT 
      t1.partner_id,
      t1.partner_name,
      t1.profile_details.residing_country,
      d as domain
    FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
    UNNEST(t1.partner_details.email_domain) as d
    WHERE LOWER(t1.partner_name) LIKE '%${keyword.toLowerCase()}%'
       OR LOWER(d) LIKE '%${domain.toLowerCase()}%'
  `;
  
  Logger.log(`Running Debug Search for '${keyword}' or '${domain}'...`);

  try {
    const request = { query: sql, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    if (!queryResults.rows || queryResults.rows.length === 0) {
      Logger.log("No results found in BigQuery.");
      return;
    }

    Logger.log(`Found ${queryResults.rows.length} matches:`);
    queryResults.rows.forEach(row => {
      const f = row.f;
      Logger.log(`[${f[0].v}] ${f[1].v} | Country: ${f[2].v} | Domain: ${f[3].v}`);
    });

  } catch (e) {
    Logger.log("Query Failed: " + e.toString());
  }
}
