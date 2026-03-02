function check2026PartnerMatches() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName('LATAM_Partner_DB_2026');
  if (!sheet) {
    Logger.log("Error: LATAM_Partner_DB_2026 sheet not found.");
    return;
  }
  
  // 1. Get the 82 names from the spreadsheet
  const lastRow = sheet.getLastRow();
  // Assuming headers so we get data from A2 onwards
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
  
  let validNames = [];
  data.forEach(row => {
    let name = String(row[0]).trim();
    if (name) {
      // Escape for SQL
      validNames.push(`'${name.replace(/'/g, "\\'")}'`);
    }
  });

  if (validNames.length === 0) {
    Logger.log("No valid names found in column A.");
    return;
  }

  Logger.log(`Found ${validNames.length} names to check.`);

  // 2. Query BigQuery for matches and non-matches
  // We want to see how many of THESE names match ANY variation in the Master DB
  const SQL_QUERY = `
    SELECT DISTINCT
      t1.partner_id,
      t1.partner_name,
      t1.partner_details.vector_details.partner_group_name,
      bq_domain
    FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
    LEFT JOIN UNNEST(t1.partner_details.email_domain) AS bq_domain
    WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
    AND (
      LOWER(t1.partner_name) LIKE '%dev %' OR
      LOWER(t1.partner_name) LIKE 'dev-%' OR
      LOWER(t1.partner_name) LIKE '%forticus%' OR
      LOWER(t1.partner_name) LIKE '%codes%' OR
      LOWER(t1.partner_details.vector_details.partner_group_name) LIKE '%dev %' OR
      LOWER(t1.partner_details.vector_details.partner_group_name) LIKE 'dev-%' OR
      LOWER(t1.partner_details.vector_details.partner_group_name) LIKE '%forticus%' OR
      LOWER(t1.partner_details.vector_details.partner_group_name) LIKE '%codes%'
    )
    ORDER BY partner_name
  `;

  // 3. Execute and Write Results to a Temp Sheet
  Logger.log("Executing Match Check Query...");
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    let resultSheet = ss.getSheetByName('Temp_Match_Check_2026');
    if (!resultSheet) {
       resultSheet = ss.insertSheet('Temp_Match_Check_2026');
    } else {
       resultSheet.clear();
    }

    if (!queryResults.rows || queryResults.rows.length === 0) { 
      resultSheet.getRange('A1').setValue("0 Query results."); 
      Logger.log("No query results return from BigQuery match check.");
      return; 
    }
    
    const outputData = [];
    const headers = queryResults.schema.fields.map(field => field.name);
    outputData.push(headers); 
    
    queryResults.rows.forEach(row => { 
      const rowData = row.f.map(field => field.v === null ? "" : field.v); 
      outputData.push(rowData); 
    });
    
    resultSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // Formatting
    resultSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");
    resultSheet.autoResizeColumns(1, 4);

    Logger.log("Match check complete. See 'Temp_Match_Check_2026' tab in your Destination Sheet.");

  } catch (e) {
    Logger.log("ERROR in Match Check: " + e.toString());
  }
}
