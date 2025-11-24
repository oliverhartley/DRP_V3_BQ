/**
 * ****************************************
 * Google Apps Script - BigQuery Loader
 * File: LATAM_Partner_DB.gs
 * Version: V 4.2 - Added Total Profiles Count
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
    if (domain && domain.includes('@')) {
      const isTrue = (val) => val === true;
      let sqlLine = `STRUCT('${domain}' AS domain, ${isTrue(row[COL_MAP.GSI])} AS is_gsi, ${isTrue(row[COL_MAP.BRAZIL])} AS is_brazil, ${isTrue(row[COL_MAP.MCO])} AS is_mco, ${isTrue(row[COL_MAP.MEXICO])} AS is_mexico, ${isTrue(row[COL_MAP.PS])} AS is_ps, ${isTrue(row[COL_MAP.AI_ML])} AS is_ai_ml, ${isTrue(row[COL_MAP.GWS])} AS is_gws, ${isTrue(row[COL_MAP.SECURITY])} AS is_security, ${isTrue(row[COL_MAP.DB])} AS is_db, ${isTrue(row[COL_MAP.ANALYTICS])} AS is_analytics, ${isTrue(row[COL_MAP.INFRA])} AS is_infra, ${isTrue(row[COL_MAP.APP_MOD])} AS is_app_mod)`;
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
      WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${VIRTUAL_TABLE_DATA} ]) ),
      
      -- 1. Find Matched Partners via Domain
      MatchedPartners AS (
          SELECT DISTINCT
              t1.partner_id,
              LOGICAL_OR(IFNULL(sheet.is_gsi, FALSE)) AS is_gsi,
              LOGICAL_OR(IFNULL(sheet.is_brazil, FALSE)) AS is_brazil,
              LOGICAL_OR(IFNULL(sheet.is_mco, FALSE)) AS is_mco,
              LOGICAL_OR(IFNULL(sheet.is_mexico, FALSE)) AS is_mexico,
              LOGICAL_OR(IFNULL(sheet.is_ps, FALSE)) AS is_ps,
              LOGICAL_OR(IFNULL(sheet.is_ai_ml, FALSE)) AS is_ai_ml,
              LOGICAL_OR(IFNULL(sheet.is_gws, FALSE)) AS is_gws,
              LOGICAL_OR(IFNULL(sheet.is_security, FALSE)) AS is_security,
              LOGICAL_OR(IFNULL(sheet.is_db, FALSE)) AS is_db,
              LOGICAL_OR(IFNULL(sheet.is_analytics, FALSE)) AS is_analytics,
              LOGICAL_OR(IFNULL(sheet.is_infra, FALSE)) AS is_infra,
              LOGICAL_OR(IFNULL(sheet.is_app_mod, FALSE)) AS is_app_mod
          FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1,
          UNNEST(t1.partner_details.email_domain) AS bq_domain
          INNER JOIN Spreadsheet_Data AS sheet ON TRIM(LOWER(bq_domain)) = sheet.domain
          GROUP BY t1.partner_id
      ),
      
      -- 2. Get Profile Data for Matched Partners
      PartnerData AS (
          SELECT 
              t1.partner_id,
              t1.partner_name,
              t1.profile_details.profile_id,
              t1.profile_details.residing_country,
              mp.is_gsi, mp.is_brazil, mp.is_mco, mp.is_mexico, mp.is_ps,
              mp.is_ai_ml, mp.is_gws, mp.is_security, mp.is_db, mp.is_analytics, mp.is_infra, mp.is_app_mod,
              t1.partner_details.email_domain as domains
          FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
          JOIN MatchedPartners mp ON t1.partner_id = mp.partner_id
          WHERE t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
      ),
      
      -- 3. Aggregate Profiles by Country
      ProfileBreakdown AS (
          SELECT 
              partner_id, 
              STRING_AGG(CONCAT(residing_country, ':', count), '|') as breakdown
          FROM (
              SELECT partner_id, residing_country, COUNT(DISTINCT profile_id) as count
              FROM PartnerData
              GROUP BY partner_id, residing_country
          )
          GROUP BY partner_id
      ),
      
      -- 4. Main Aggregation
      PartnerAggregation AS (
          SELECT
              partner_id,
              partner_name,
              COUNT(DISTINCT profile_id) AS Total_Profiles,
              STRING_AGG(DISTINCT residing_country, ', ') AS Operating_Countries,
              (APPROX_TOP_COUNT(residing_country, 1))[OFFSET(0)].value AS Top_Operating_Country,
              TRUE AS Managed_Partners, -- If they are in MatchedPartners, they are managed
              LOGICAL_OR(is_gsi) AS GSI,
              LOGICAL_OR(is_brazil) AS Brazil,
              LOGICAL_OR(is_mco) AS MCO,
              LOGICAL_OR(is_mexico) AS Mexico,
              LOGICAL_OR(is_ps) AS PS,
              LOGICAL_OR(is_ai_ml) AS AI_ML,
              LOGICAL_OR(is_gws) AS GWS,
              LOGICAL_OR(is_security) AS Security,
              LOGICAL_OR(is_db) AS DB,
              LOGICAL_OR(is_analytics) AS Analytics,
              LOGICAL_OR(is_infra) AS Infra,
              LOGICAL_OR(is_app_mod) AS App_Mod,
              ARRAY_CONCAT_AGG(domains) AS raw_partner_domains
          FROM PartnerData
          GROUP BY partner_id, partner_name
      )
      SELECT 
          pa.* EXCEPT (raw_partner_domains), 
          pb.breakdown AS Profile_Breakdown,
          (SELECT STRING_AGG(DISTINCT domain, ', ') FROM UNNEST(pa.raw_partner_domains) AS domain WHERE domain IS NOT NULL) AS Partner_Domains
      FROM PartnerAggregation AS pa
      LEFT JOIN ProfileBreakdown AS pb ON pa.partner_id = pb.partner_id;
    `;

    // ... (Rest of the execution code is standard) ...
    const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
    const sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
    if (!sheet) { Logger.log("Error: Hoja no encontrada."); return; }
    Logger.log("Iniciando consulta...");
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