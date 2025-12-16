/**
 * ****************************************
 * Google Apps Script - BigQuery Loader
 * File: LATAM_Partner_DB.gs
 * Version: V 5.9 - Fixed Filter Column Names
 * ****************************************
 */

// NOTE: PROJECT_ID, DESTINATION_SS_ID, and SOURCE_SS_ID are defined in Config.gs

const DESTINATION_SHEET_NAME = "LATAM_Partner_DB";
const DOMAIN_START_ROW = 2;

const COL_MAP = {
  PARTNER_NAME: 0,
  DOMAIN: 1,
  COUNTRIES_START: 2,
  REGIONS_START: 21,
  PRODUCTS_START: 24,
  EMAIL_TO: 49,
  EMAIL_CC: 50
};

function getSpreadsheetDataAsSqlStruct() {
  const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_SOURCE);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME_SOURCE}" not found in Source Spreadsheet.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < DOMAIN_START_ROW) return "";
  const range = sheet.getRange(DOMAIN_START_ROW, 1, lastRow - DOMAIN_START_ROW + 1, 51); // 51 columns total
  const values = range.getValues();
  const textStyles = sheet.getRange(DOMAIN_START_ROW, 1, lastRow - DOMAIN_START_ROW + 1, 1).getTextStyles();
  let structList = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (textStyles[i][0].isStrikethrough()) continue;
    let domain = String(row[COL_MAP.DOMAIN]).toLowerCase().trim().replace(/[\x00-\x1F\x7F-\x9F\u200B]/g, "");
    let partnerName = String(row[COL_MAP.PARTNER_NAME] || "").trim().replace(/[\x00-\x1F\x7F-\x9F\u200B]/g, "");

    if (domain && !domain.includes('#n/a')) {
      const escapedDomain = domain.replace(/'/g, "\\'"); // Escape single quotes for SQL
      const escapedName = partnerName.replace(/'/g, "\\'"); // Escape single quotes for SQL
      const isTrue = (val) => val === true || String(val).toUpperCase() === 'TRUE';

      // Regions
      const mco = isTrue(row[COL_MAP.REGIONS_START]);
      const gsi = isTrue(row[COL_MAP.REGIONS_START + 1]);
      const ps = isTrue(row[COL_MAP.REGIONS_START + 2]);

      // Solutions (Aggregated from products for BigQuery)
      const is_infra = [row[24], row[25], row[26], row[27], row[28]].some(isTrue);
      const is_app_mod = [row[29], row[30]].some(isTrue);
      const is_db = [row[31], row[32], row[33], row[34], row[35]].some(isTrue);
      const is_analytics = [row[36], row[37], row[38], row[39]].some(isTrue);
      const is_ai_ml = [row[40], row[41], row[42], row[43]].some(isTrue);
      const is_security = [row[44], row[45], row[46], row[47]].some(isTrue);
      const is_gws = isTrue(row[48]);

      let sqlLine = `STRUCT('${escapedDomain}' AS domain, '${escapedName}' AS partner_name, ${gsi} AS is_gsi, false AS is_brazil, ${mco} AS is_mco, false AS is_mexico, ${ps} AS is_ps, ${is_ai_ml} AS is_ai_ml, ${is_gws} AS is_gws, ${is_security} AS is_security, ${is_db} AS is_db, ${is_analytics} AS is_analytics, ${is_infra} AS is_infra, ${is_app_mod} AS is_app_mod)`;
      structList.push(sqlLine);
    }
  }
  return structList.join(',\n');
}

function runBigQueryQuery() {
  try {
    Logger.log("Generando tabla virtual desde spreadsheet...");
    const VIRTUAL_TABLE_DATA = getSpreadsheetDataAsSqlStruct();
    if (!VIRTUAL_TABLE_DATA) { Logger.log("Error: No se encontraron datos."); return; }

    const SQL_QUERY = `
      -- Query Version: ${new Date().toISOString()}
      WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${VIRTUAL_TABLE_DATA} ]) ),
      
      -- 1. Flatten BigQuery Data (Safe Pre-processing)
      BQ_Flattened AS (
        SELECT 
           t1.partner_id,
           t1.partner_name,
           t1.profile_details.profile_id,
           t1.profile_details.residing_country,
           LOWER(bq_domain) as bq_domain_flat
        FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
        UNNEST(t1.partner_details.email_domain) AS bq_domain
        WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
      ),

      -- 2. Join Sheet (Left) -> BQ (Right) using the flattened table
      -- This ensures ALL sheet rows are kept, even if no BQ match found
      RawData AS (
          SELECT
              bq.partner_id,
              bq.partner_name,
              bq.profile_id,
              bq.residing_country,
              sheet.domain IS NOT NULL as is_matched, 
              sheet.domain as sheet_domain, 
              -- Keep Sheet Name for Fallback
              sheet.partner_name as sheet_partner_name,
              IFNULL(sheet.is_gsi, FALSE) as is_gsi,
              IFNULL(sheet.is_brazil, FALSE) as is_brazil,
              IFNULL(sheet.is_mco, FALSE) as is_mco,
              IFNULL(sheet.is_mexico, FALSE) as is_mexico,
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
            -- FALLBACK MATCH: If Domain fails, try Exact Name Match
            OR TRIM(LOWER(bq.partner_name)) = TRIM(LOWER(sheet.partner_name))
      ),
      
      -- 3. Get Unique Profiles
      UniqueProfiles AS (
          SELECT DISTINCT
              -- Use BQ ID if matched, otherwise generate a placeholder using Sheet Name
              IFNULL(partner_id, CONCAT('MISSING_BQ_', REGEXP_REPLACE(sheet_partner_name, ' ', '_'))) as partner_id, 
              IFNULL(partner_name, sheet_partner_name) as partner_name, 
              profile_id,
              residing_country,
              sheet_domain
          FROM RawData
      ),
      
      -- 4. Aggregate Partner Flags
      PartnerFlags AS (
          SELECT
              IFNULL(partner_id, CONCAT('MISSING_BQ_', REGEXP_REPLACE(sheet_partner_name, ' ', '_'))) as partner_id,
              TRUE as is_matched, 
              LOGICAL_OR(is_gsi) as is_gsi,
              LOGICAL_OR(is_brazil) as is_brazil,
              LOGICAL_OR(is_mco) as is_mco,
              LOGICAL_OR(is_mexico) as is_mexico,
              LOGICAL_OR(is_ps) as is_ps,
              LOGICAL_OR(is_ai_ml) as is_ai_ml,
              LOGICAL_OR(is_gws) as is_gws,
              LOGICAL_OR(is_security) as is_security,
              LOGICAL_OR(is_db) as is_db,
              LOGICAL_OR(is_analytics) as is_analytics,
              LOGICAL_OR(is_infra) as is_infra,
              LOGICAL_OR(is_app_mod) as is_app_mod,
              ARRAY_AGG(DISTINCT sheet_domain) as domains
          FROM RawData
          GROUP BY partner_id
      ),
      
      -- 5. Profile Breakdown
      ProfileBreakdown_Prep AS (
          SELECT partner_id, residing_country, COUNT(DISTINCT profile_id) as count
          FROM UniqueProfiles
          WHERE profile_id IS NOT NULL 
          GROUP BY partner_id, residing_country
      ),
      
      ProfileBreakdown AS (
          SELECT 
              partner_id, 
              STRING_AGG(CONCAT(residing_country, ':', CAST(count AS STRING)), '|') as breakdown
          FROM ProfileBreakdown_Prep
          GROUP BY partner_id
      ),
      
      -- 6. Final Aggregation
      PartnerAggregation AS (
          SELECT
              up.partner_id,
              MAX(up.partner_name) as partner_name, 
              COUNT(DISTINCT up.profile_id) AS Total_Profiles,
              STRING_AGG(DISTINCT up.residing_country, ', ') AS Operating_Countries,
              (APPROX_TOP_COUNT(up.residing_country, 1))[OFFSET(0)].value AS Top_Operating_Country,
              pf.is_matched AS Managed_Partners,
              pf.is_gsi AS GSI, pf.is_brazil AS Brazil, pf.is_mco AS MCO, pf.is_mexico AS Mexico, pf.is_ps AS PS,
              pf.is_ai_ml, pf.is_gws, pf.is_security, pf.is_db, pf.is_analytics, pf.is_infra, pf.is_app_mod,
              pf.domains
          FROM UniqueProfiles up
          JOIN PartnerFlags pf ON up.partner_id = pf.partner_id
          GROUP BY up.partner_id, pf.is_matched, pf.is_gsi, pf.is_brazil, pf.is_mco, pf.is_mexico, pf.is_ps, pf.is_ai_ml, pf.is_gws, pf.is_security, pf.is_db, pf.is_analytics, pf.is_infra, pf.is_app_mod, pf.domains
      )
      SELECT 
          pa.* EXCEPT (domains), 
          pb.breakdown AS Profile_Breakdown,
          (SELECT STRING_AGG(DISTINCT domain, ', ') FROM UNNEST(pa.domains) AS domain WHERE domain IS NOT NULL) AS Partner_Domains
      FROM PartnerAggregation AS pa
      LEFT JOIN ProfileBreakdown AS pb ON pa.partner_id = pb.partner_id;
    `;

    // ... (Rest of the execution code is standard) ...
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    let sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(DESTINATION_SHEET_NAME);
      Logger.log("Sheet '" + DESTINATION_SHEET_NAME + "' created.");
    }
    Logger.log("Iniciando consulta...");
    Logger.log("SQL Query: " + SQL_QUERY); // Added logging
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    if (!queryResults.rows || queryResults.rows.length === 0) { sheet.clearContents(); sheet.getRange('A1').setValue("0 resultados."); return; }
    const data = [];
    const headers = queryResults.schema.fields.map(field => field.name);
    data.push(headers);
    queryResults.rows.forEach(row => { const rowData = row.f.map(field => field.v === null ? "" : field.v); data.push(rowData); });
    sheet.clearContents();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    Logger.log("Carga completa.");
  } catch (e) { Logger.log("ERROR: " + e.toString()); }
}