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

function backfillSpreadsheetLinks2026() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const dbSheet = ss.getSheetByName('LATAM_Partner_DB_2026');

  if (!dbSheet) {
    Logger.log("Error: LATAM_Partner_DB_2026 sheet not found.");
    return;
  }

  const lastRow = dbSheet.getLastRow();
  // Assume Partner Name is in Col A (1)
  // We want to write ID to Col I (9) and URL to Col J (10)
  const dataRange = dbSheet.getRange(2, 1, lastRow - 1, 1);
  const partnerNames = dataRange.getValues();

  // Track which rows belong to which partner (1-based index)
  const rowMap = {};
  for (let i = 0; i < partnerNames.length; i++) {
    const pName = String(partnerNames[i][0]).trim();
    if (!pName) continue;
    if (!rowMap[pName]) rowMap[pName] = [];
    rowMap[pName].push(i + 2);
  }

  Logger.log(`Found ${Object.keys(rowMap).length} unique partners in DB.`);

  // Create header if missing
  dbSheet.getRange("I1").setValue("Spreadsheet_ID").setBackground("#d9d9d9").setFontWeight("bold");
  dbSheet.getRange("J1").setValue("Spreadsheet_URL").setBackground("#d9d9d9").setFontWeight("bold");

  // Scan the target folder for existing decks
  // Assuming DESTINATION_FOLDER_ID is defined in Config.js
  if (typeof DESTINATION_FOLDER_ID === 'undefined') {
    Logger.log("Error: DESTINATION_FOLDER_ID not found in Config. Please define it or hardcode the folder ID here.");
    return;
  }

  const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let matchCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    // Decks are usually named "Partner Name - 2026 Deck" or similar.
    // We try to find a DB partner name that is completely contained within the file name.
    let matchedPartner = null;
    for (const pName of Object.keys(rowMap)) {
      // Very basic inclusion check. Might need refinement based on exact naming convention.
      if (fileName.toLowerCase().includes(pName.toLowerCase())) {
        matchedPartner = pName;
        break;
      }
    }

    if (matchedPartner) {
      const fileId = file.getId();
      const fileUrl = file.getUrl();

      const rowsToUpdate = rowMap[matchedPartner];
      for (const r of rowsToUpdate) {
        dbSheet.getRange(r, 9).setValue(fileId);
        dbSheet.getRange(r, 10).setValue(fileUrl);
      }
      matchCount++;
      Logger.log(`Mapped: ${matchedPartner} -> ${fileUrl}`);
    }
  }

  Logger.log(`Backfill Complete! Updated ${matchCount} partner spreadsheets.`);
}
