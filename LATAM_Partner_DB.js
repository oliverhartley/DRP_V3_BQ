/**
 * ****************************************
 * Google Apps Script - BigQuery Loader
 * File: LATAM_Partner_DB.gs
 * Version: V 5.3 - Fixed Join Syntax & Single Quotes
 * ****************************************
 */

// NOTE: PROJECT_ID, DESTINATION_SS_ID, and SOURCE_SS_ID are defined in Config.gs

const DESTINATION_SHEET_NAME = "LATAM_Partner_DB";
const DOMAIN_START_ROW = 3;

const COL_MAP = {
  GSI: 7, BRAZIL: 8, MCO: 9, MEXICO: 10, PS: 11, 
  AI_ML: 13, GWS: 14, SECURITY: 15, DB: 16, ANALYTICS: 17, INFRA: 18, APP_MOD: 19, 
  DOMAIN: 34
};

// ... [getSpreadsheetDataAsSqlStruct function remains EXACTLY THE SAME as V4.1] ...
// (I am omitting the helper function here to save space, please keep it from the previous version)
function getSpreadsheetDataAsSqlStruct() {
  const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_SOURCE);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME_SOURCE}" not found in Source Spreadsheet.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < DOMAIN_START_ROW) return ""; 
  const range = sheet.getRange(DOMAIN_START_ROW, 1, lastRow - DOMAIN_START_ROW + 1, 35); 
  const values = range.getValues();
  const textStyles = sheet.getRange(DOMAIN_START_ROW, 1, lastRow - DOMAIN_START_ROW + 1, 1).getTextStyles();
  let structList = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (textStyles[i][0].isStrikethrough()) continue; 
    let domain = String(row[COL_MAP.DOMAIN]).toLowerCase().trim().replace(/[\x00-\x1F\x7F-\x9F\u200B]/g, "");
    if (domain && !domain.startsWith('@')) domain = '@' + domain;
    if (domain && domain.includes('@') && !domain.includes('#n/a')) {
      const escapedDomain = domain.replace(/'/g, "\\'"); // Escape single quotes for SQL
      const isTrue = (val) => val === true;
      let sqlLine = `STRUCT('${escapedDomain}' AS domain, ${isTrue(row[COL_MAP.GSI])} AS is_gsi, ${isTrue(row[COL_MAP.BRAZIL])} AS is_brazil, ${isTrue(row[COL_MAP.MCO])} AS is_mco, ${isTrue(row[COL_MAP.MEXICO])} AS is_mexico, ${isTrue(row[COL_MAP.PS])} AS is_ps, ${isTrue(row[COL_MAP.AI_ML])} AS is_ai_ml, ${isTrue(row[COL_MAP.GWS])} AS is_gws, ${isTrue(row[COL_MAP.SECURITY])} AS is_security, ${isTrue(row[COL_MAP.DB])} AS is_db, ${isTrue(row[COL_MAP.ANALYTICS])} AS is_analytics, ${isTrue(row[COL_MAP.INFRA])} AS is_infra, ${isTrue(row[COL_MAP.APP_MOD])} AS is_app_mod)`;
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
      
      -- 1. Get Raw Data with Join to Spreadsheet
      RawData AS (
          SELECT
              t1.partner_id,
              t1.partner_name,
              t1.profile_details.profile_id,
              t1.profile_details.residing_country,
              sheet.domain IS NOT NULL as is_matched,
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
              IFNULL(sheet.is_app_mod, FALSE) as is_app_mod,
              bq_domain
          FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
          CROSS JOIN UNNEST(t1.partner_details.email_domain) AS bq_domain
          LEFT JOIN Spreadsheet_Data AS sheet ON TRIM(LOWER(bq_domain)) = sheet.domain
          WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
      ),
      
      -- 2. Filter to Matched Partners & Get Unique Profiles
      UniqueProfiles AS (
          SELECT DISTINCT
              partner_id,
              partner_name,
              profile_id,
              residing_country
          FROM RawData
          WHERE is_matched = true
      ),
      
      -- 3. Aggregate Partner Flags (Handle multiple matched domains)
      PartnerFlags AS (
          SELECT
              partner_id,
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
              ARRAY_AGG(DISTINCT bq_domain) as domains
          FROM RawData
          WHERE is_matched = true
          GROUP BY partner_id
      ),
      
      -- 4. Profile Breakdown Prep
      ProfileBreakdown_Prep AS (
          SELECT partner_id, residing_country, COUNT(DISTINCT profile_id) as count
          FROM UniqueProfiles
          GROUP BY partner_id, residing_country
      ),
      
      -- 5. Profile Breakdown Aggregation
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
              up.partner_name,
              COUNT(DISTINCT up.profile_id) AS Total_Profiles,
              STRING_AGG(DISTINCT up.residing_country, ', ') AS Operating_Countries,
              (APPROX_TOP_COUNT(up.residing_country, 1))[OFFSET(0)].value AS Top_Operating_Country,
              TRUE AS Managed_Partners,
              pf.is_gsi, pf.is_brazil, pf.is_mco, pf.is_mexico, pf.is_ps,
              pf.is_ai_ml, pf.is_gws, pf.is_security, pf.is_db, pf.is_analytics, pf.is_infra, pf.is_app_mod,
              pf.domains
          FROM UniqueProfiles up
          JOIN PartnerFlags pf ON up.partner_id = pf.partner_id
          GROUP BY up.partner_id, up.partner_name, pf.is_gsi, pf.is_brazil, pf.is_mco, pf.is_mexico, pf.is_ps, pf.is_ai_ml, pf.is_gws, pf.is_security, pf.is_db, pf.is_analytics, pf.is_infra, pf.is_app_mod, pf.domains
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
    const sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
    if (!sheet) { Logger.log("Error: Hoja no encontrada."); return; }
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