/**
 * ****************************************
 * Google Apps Script - BigQuery Infrastructure
 * File: BigQuery_Core.js
 * Description: Modular BigQuery query builder and execution engine.
 * ****************************************
 */

/**
 * Main execution function for the Partner DB BigQuery Query.
 */
function runBigQueryLoader() {
  try {
    const virtualTableData = getPartnerSqlStruct();
    if (!virtualTableData) return;

    const sql = getPartnerLoaderSql(virtualTableData);
    const results = executeBigQuery(sql);
    
    if (results) {
      persistToSheet(results, SHEETS.DB_LOAD);
    }
  } catch (e) {
    Logger.log(`[BigQuery_Core] ERROR: ${e.toString()}`);
  }
}

/**
 * Builds the SQL query for the Partner DB Loader.
 */
function getPartnerLoaderSql(virtualTableData) {
  return `
    WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${virtualTableData} ]) ),
    
    BQ_Flattened AS (
      SELECT 
         t1.partner_id,
         t1.partner_name,
         t1.profile_details.profile_id,
         t1.profile_details.residing_country,
         LOWER(bq_domain) as bq_domain_flat
      FROM \`${PROJECT_ID}.service_partnercoe.drp_partner_master\` AS t1,
      UNNEST(t1.partner_details.email_domain) AS bq_domain
      WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
    ),

    RawData AS (
        SELECT
            bq.partner_id,
            bq.partner_name,
            bq.profile_id,
            bq.residing_country,
            bq.partner_id IS NOT NULL as is_matched, 
            sheet.domain as sheet_domain, 
            sheet.partner_name as sheet_partner_name,
            IFNULL(sheet.is_gsi, FALSE) as is_gsi,
            IFNULL(sheet.is_mco, FALSE) as is_mco,
            IFNULL(sheet.is_ps, FALSE) as is_ps,
            IFNULL(sheet.is_ai_ml, FALSE) as is_ai_ml,
            IFNULL(sheet.is_gws, FALSE) as is_gws,
            IFNULL(sheet.is_security, FALSE) as is_security,
            IFNULL(sheet.is_db, FALSE) as is_db,
            IFNULL(sheet.is_analytics, FALSE) as is_analytics,
            IFNULL(sheet.is_infra, FALSE) as is_infra,
            IFNULL(sheet.is_app_mod, FALSE) as is_app_mod
        FROM Spreadsheet_Data AS sheet
        LEFT JOIN BQ_Flattened AS bq
          ON REGEXP_REPLACE(TRIM(LOWER(bq.bq_domain_flat)), r'^@', '') = REGEXP_REPLACE(TRIM(LOWER(sheet.domain)), r'^@', '')
          OR TRIM(LOWER(bq.partner_name)) = TRIM(LOWER(sheet.partner_name))
    ),
    
    PartnerAggregation AS (
        SELECT
            IFNULL(partner_id, CONCAT('MISSING_BQ_', REGEXP_REPLACE(sheet_partner_name, ' ', '_'))) as partner_id,
            MAX(IFNULL(partner_name, sheet_partner_name)) as partner_name, 
            COUNT(DISTINCT profile_id) AS Total_Profiles,
            STRING_AGG(DISTINCT residing_country, ', ') AS Operating_Countries,
            (APPROX_TOP_COUNT(residing_country, 1))[OFFSET(0)].value AS Top_Operating_Country,
            LOGICAL_OR(is_matched) AS Managed_Partners,
            LOGICAL_OR(is_gsi) AS GSI, 
            LOGICAL_OR(is_mco) AS MCO, 
            LOGICAL_OR(is_ps) AS PS,
            LOGICAL_OR(is_ai_ml) as is_ai_ml, 
            LOGICAL_OR(is_gws) as is_gws, 
            LOGICAL_OR(is_security) as is_security, 
            LOGICAL_OR(is_db) as is_db, 
            LOGICAL_OR(is_analytics) as is_analytics, 
            LOGICAL_OR(is_infra) as is_infra, 
            LOGICAL_OR(is_app_mod) as is_app_mod,
            ARRAY_AGG(DISTINCT sheet_domain IGNORE NULLS) as domains
        FROM RawData
        GROUP BY 1
    )
    
    SELECT 
        pa.* EXCEPT (domains), 
        (SELECT STRING_AGG(DISTINCT domain, ', ') FROM UNNEST(pa.domains) AS domain) AS Partner_Domains
    FROM PartnerAggregation AS pa;
  `;
}

/**
 * Generic BigQuery Executor.
 */
function executeBigQuery(sql) {
  Logger.log(`[BigQuery] Executing Query: ${sql.substring(0, 100)}...`);
  const request = { query: sql, useLegacySql: false };
  const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
  
  if (!queryResults.rows || queryResults.rows.length === 0) {
    Logger.log("[BigQuery] No results.");
    return null;
  }

  const data = [];
  const headers = queryResults.schema.fields.map(field => field.name);
  data.push(headers);
  
  queryResults.rows.forEach(row => {
    const rowData = row.f.map(field => field.v === null ? "" : field.v);
    data.push(rowData);
  });
  
  return data;
}

/**
 * Persists BigQuery Results to a Sheet.
 */
function persistToSheet(data, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  Logger.log(`[BigQuery] Successfully loaded ${data.length} rows into ${sheetName}.`);
}

/**
 * Helper to build the Virtual Table from local spreadsheet.
 */
function getPartnerSqlStruct() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SOURCE);
  if (!sheet) return null;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  const range = sheet.getRange(2, 1, lastRow - 1, 51);
  const values = range.getValues();
  const structs = [];
  
  for (const row of values) {
    const name = String(row[COL_INDEX.PARTNER_NAME]).replace(/'/g, "\\'");
    const domain = String(row[COL_INDEX.DOMAIN]).replace(/'/g, "\\'");
    
    if (!domain) continue;
    
    const isTrue = (val) => val === true || String(val).toUpperCase() === 'TRUE';
    
    const mco = isTrue(row[COL_INDEX.REGIONS_START]);
    const gsi = isTrue(row[COL_INDEX.REGIONS_START+1]);
    const ps = isTrue(row[COL_INDEX.REGIONS_START+2]);
    
    const is_infra = [row[24], row[25], row[26], row[27], row[28]].some(isTrue);
    const is_app_mod = [row[29], row[30]].some(isTrue);
    const is_db = [row[31], row[32], row[33], row[34], row[35]].some(isTrue);
    const is_analytics = [row[36], row[37], row[38], row[39]].some(isTrue);
    const is_ai_ml = [row[40], row[41], row[42], row[43]].some(isTrue);
    const is_security = [row[44], row[45], row[46], row[47]].some(isTrue);
    const is_gws = isTrue(row[48]);

    structs.push(`STRUCT('${domain}' AS domain, '${name}' AS partner_name, ${gsi} AS is_gsi, ${mco} AS is_mco, ${ps} AS is_ps, ${is_ai_ml} AS is_ai_ml, ${is_gws} AS is_gws, ${is_security} AS is_security, ${is_db} AS is_db, ${is_analytics} AS is_analytics, ${is_infra} AS is_infra, ${is_app_mod} AS is_app_mod)`);
  }
  
  return structs.join(',\n');
}

/**
 * Builds SQL to fetch ALL LATAM partners (Managed + Unmanaged).
 */
function getAllLatamPartnersSql() {
  return `
    SELECT DISTINCT
        REGEXP_REPLACE(TRIM(LOWER(t1.partner_details.email_domain[OFFSET(0)])), r'^@', '') AS domain,
        t1.partner_name,
        t1.profile_details.residing_country
    FROM \`${PROJECT_ID}.service_partnercoe.drp_partner_master\` AS t1
    WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
    AND t1.partner_details.email_domain IS NOT NULL
    AND ARRAY_LENGTH(t1.partner_details.email_domain) > 0
  `;
}

/**
 * Builds SQL to fetch distinct countries for each partner domain.
 */
function getPartnerCountryPresenceSql() {
  return `
    SELECT
        REGEXP_REPLACE(TRIM(LOWER(t1.partner_details.email_domain[OFFSET(0)])), r'^@', '') AS domain,
        ARRAY_AGG(DISTINCT t1.profile_details.residing_country IGNORE NULLS) AS operating_countries
    FROM \`${PROJECT_ID}.service_partnercoe.drp_partner_master\` AS t1
    WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
    AND t1.partner_details.email_domain IS NOT NULL
    AND ARRAY_LENGTH(t1.partner_details.email_domain) > 0
    GROUP BY 1
  `;
}
