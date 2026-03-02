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
    WITH SheetNames AS (
      SELECT name FROM UNNEST([${validNames.join(', ')}]) AS name
    ),
    BQ_Names AS (
      SELECT DISTINCT 
        t1.partner_id,
        t1.partner_name AS bq_primary_name,
        LOWER(TRIM(t1.partner_name)) AS bq_primary_name_clean,
        alias
      FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
      UNNEST(CASE WHEN ARRAY_LENGTH(t1.partner_details.alias) > 0 THEN t1.partner_details.alias ELSE [t1.partner_name] END) as alias
      WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
    )
    SELECT 
      sn.name AS Spreadsheet_Name,
      CASE 
        WHEN bq.partner_id IS NOT NULL THEN 'MATCHED (Exact or Alias)'
        ELSE 'NOT FOUND'
      END AS Match_Status,
      bq.partner_id AS Matched_BQ_ID,
      bq.bq_primary_name AS Matched_BQ_Primary_Name
    FROM SheetNames sn
    LEFT JOIN BQ_Names bq 
      ON LOWER(TRIM(sn.name)) = LOWER(TRIM(bq.alias)) 
      OR LOWER(TRIM(sn.name)) = bq.bq_primary_name_clean
    ORDER BY Match_Status DESC, sn.name
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
