/**
 * ****************************************
 * Google Apps Script - BigQuery Loader (2026 V1)
 * File: LATAM_Partner_DB_2026.js
 * Description: Reads directly from the Destination Sheet (LATAM_Partner_DB_2026)
 * and matches against BQ using partner_name or partner_group_name.
 * ****************************************
 */

const SHEET_NAME_2026 = "LATAM_Partner_DB_2026";
const START_ROW_2026 = 2; // Assuming row 1 is headers

const COL_MAP_2026 = {
  NAME: 0,        // Column A: partner_group_name
  SUBREGION: 1,   // Column B: Sub-region
  PDM: 2,         // Column C: PDM
  TYPE: 3         // Column D: Type of Partner
};

function getSpreadsheetData2026AsSqlStruct() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_2026);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME_2026}" not found in Destination Spreadsheet.`);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW_2026) return ""; 
  
  const range = sheet.getRange(START_ROW_2026, 1, lastRow - START_ROW_2026 + 1, 4); 
  const values = range.getValues();
  
  let structList = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    let partnerName = String(row[COL_MAP_2026.NAME] || "").trim().replace(/[\x00-\x1F\x7F-\x9F\u200B]/g, "");
    let subRegion = String(row[COL_MAP_2026.SUBREGION] || "").trim().replace(/'/g, "\\'");
    let pdm = String(row[COL_MAP_2026.PDM] || "").trim().replace(/'/g, "\\'");
    let type = String(row[COL_MAP_2026.TYPE] || "").trim().replace(/'/g, "\\'");

    if (partnerName) {
      const escapedName = partnerName.replace(/'/g, "\\'"); 
      const internalId = `ROW_${i + START_ROW_2026}`;
      let sqlLine = `STRUCT('${internalId}' AS internal_id, '${escapedName}' AS partner_name, '${subRegion}' AS sub_region, '${pdm}' AS pdm, '${type}' AS partner_type)`;
      structList.push(sqlLine);
    }
  }
  return structList.join(',\n');
}

function runBigQueryQuery2026() {
  try {
    Logger.log("Generating virtual table from 2026 spreadsheet...");
    const VIRTUAL_TABLE_DATA = getSpreadsheetData2026AsSqlStruct();
    if (!VIRTUAL_TABLE_DATA) { Logger.log("Error: No data found in Column A."); return; }

    const SQL_QUERY = `
      -- Query Version: ${new Date().toISOString()}
      WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${VIRTUAL_TABLE_DATA} ]) ),
      
      -- 1. Flatten BigQuery Data
      BQ_Flattened AS (
        SELECT DISTINCT
           t1.partner_id,
           t1.partner_name,
           t1.profile_details.residing_country,
           t1.profile_details.profile_id,
           LOWER(TRIM(t1.partner_name)) AS bq_primary_name_clean,
           LOWER(TRIM(t1.partner_details.vector_details.partner_group_name)) as bq_group_name_clean,
           LOWER(bq_domain) as bq_domain_flat,
           REPLACE(LOWER(bq_domain), '@', '') as bq_domain_clean
        FROM \\\`concord-prod.service_partnercoe.drp_partner_master\\\` AS t1
        LEFT JOIN UNNEST(t1.partner_details.email_domain) AS bq_domain
        WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
      ),

      -- 1.5 Manual Domain Overrides for Missing Partners
      Manual_Overrides AS (
          SELECT 'AI Inversiones' as sheet_name, 'canvia.com' as domain UNION ALL
          SELECT 'CoreBi' as sheet_name, 'corebi.com.ar' as domain UNION ALL
          SELECT 'ENGINEERING DO BRASIL S/A' as sheet_name, 'engdb.com.br' as domain UNION ALL
          SELECT 'SERPRO' as sheet_name, 'serpro.gov.br' as domain UNION ALL
          SELECT 'DEV- CODES SAS' as sheet_name, 'dev-codes.com' as domain
      ),

      -- 2. Join Sheet (Left) -> BQ (Right) using Exact Name, Group Name, OR Manual Domain Override
      RawData AS (
          SELECT
              sheet.internal_id,
              bq.partner_id,
              bq.partner_name as bq_partner_name,
              bq.profile_id,
              bq.residing_country,
              bq.bq_domain_flat,
              bq.bq_domain_clean,
              CASE WHEN bq.partner_id IS NOT NULL THEN TRUE ELSE FALSE END as is_matched,
              sheet.partner_name as sheet_partner_name,
              sheet.sub_region,
              sheet.pdm,
              sheet.partner_type
          FROM Spreadsheet_Data AS sheet
          LEFT JOIN Manual_Overrides mo ON LOWER(TRIM(sheet.partner_name)) = LOWER(TRIM(mo.sheet_name))
          LEFT JOIN BQ_Flattened AS bq
            ON LOWER(TRIM(sheet.partner_name)) = bq.bq_primary_name_clean
            OR LOWER(TRIM(sheet.partner_name)) = bq.bq_group_name_clean
            OR (mo.domain IS NOT NULL AND bq.bq_domain_clean = mo.domain)
      ),
      
      -- 3. Get Unique Profiles (per row)
      UniqueProfiles AS (
          SELECT DISTINCT
              internal_id,
              IFNULL(partner_id, CONCAT('MISSING_BQ_', REGEXP_REPLACE(sheet_partner_name, ' ', '_'))) as partner_id, 
              sheet_partner_name as partner_name, 
              profile_id,
              residing_country,
              bq_domain_clean,
              sub_region,
              pdm,
              partner_type,
              is_matched
          FROM RawData
      ),
      
      -- 5. Profile Breakdown
      ProfileBreakdown_Prep AS (
          SELECT internal_id, residing_country, COUNT(DISTINCT profile_id) as count
          FROM UniqueProfiles
          WHERE profile_id IS NOT NULL 
          GROUP BY internal_id, residing_country
      ),
      
      ProfileBreakdown AS (
          SELECT 
              internal_id, 
              STRING_AGG(CONCAT(residing_country, ':', CAST(count AS STRING)), '|') as breakdown
          FROM ProfileBreakdown_Prep
          GROUP BY internal_id
      ),
      
      -- 6. Final Aggregation
      PartnerAggregation AS (
          SELECT
              up.internal_id,
              MAX(up.partner_id) as partner_id,
              MAX(up.partner_name) as Partner_Name, 
              COUNT(DISTINCT up.profile_id) AS Total_Profiles,
              MAX(up.is_matched) AS Matched_In_BQ,
              MAX(up.sub_region) AS Sub_Region,
              MAX(up.pdm) AS PDM,
              MAX(up.partner_type) AS Type_of_Partner,
              ARRAY_AGG(DISTINCT up.bq_domain_clean IGNORE NULLS) as domains
          FROM UniqueProfiles up
          GROUP BY up.internal_id
      )
      SELECT 
          pa.Partner_Name,
          pa.Sub_Region,
          pa.PDM,
          pa.Type_of_Partner,
          pa.internal_id AS Internal_ID,
          pa.partner_id AS Partner_ID,
          pa.Total_Profiles,
          pa.Matched_In_BQ,
          pb.breakdown AS Profile_Breakdown,
          (SELECT STRING_AGG(DISTINCT domain, ', ') FROM UNNEST(pa.domains) AS domain WHERE domain IS NOT NULL) AS Discovered_Domains
      FROM PartnerAggregation AS pa
      LEFT JOIN ProfileBreakdown AS pb ON pa.internal_id = pb.internal_id
      ORDER BY pa.Partner_Name;
    `;

    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_2026);
    if (!sheet) { Logger.log("Error: Sheet not found."); return; }
    
    Logger.log("Executing 2026 query...");
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    if (!queryResults.rows || queryResults.rows.length === 0) { 
      sheet.getRange('A2').setValue("No results returned."); 
      return; 
    }
    
    const data = [];
    const headers = queryResults.schema.fields.map(field => field.name);
    data.push(headers); 
    queryResults.rows.forEach(row => { 
      const rowData = row.f.map(field => field.v === null ? "" : field.v); 
      data.push(rowData); 
    });
    
    // OVERWRITE ENTIRE SHEET (Clean slate architecture)
    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Formatting
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    if (data.length > 1) {
      sheet.getRange(2, 7, data.length - 1, 1).setNumberFormat("#,##0"); // Profiles column is now 7th
    }
    
    Logger.log("2026 Data load complete!");
  } catch (e) { 
    Logger.log("ERROR 2026: " + e.toString()); 
  }
}
