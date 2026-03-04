/**
 * ****************************************
 * Google Apps Script - Web App Dashboard
 * File: WebApp.js
 * Description: Serves the 2026 Collaborative Web App Dashboard.
 * ****************************************
 */

function doGet(e) {
  const htmlOutput = HtmlService.createTemplateFromFile('Dashboard_UI');
  return htmlOutput.evaluate()
    .setTitle('DRP Partner Dashboard 2026')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Fetches data from the LATAM_Partner_Score_2026 sheet for the web app.
 * Returns a JSON object with metadata for slicers and the raw row data.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_SCORE_2026);
  if (!sheet) throw new Error("2026 Score Sheet not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 3) return { success: false, message: "No data available." };

  const headersMap = {
    solution: data[0],
    product: data[1],
    tier: data[2]
  };

  const rows = [];
  const slicerOptions = {
    types: new Set(),
    subRegions: new Set(),
    pdms: new Set(),
    solutions: new Set(),
    products: new Set()
  };

  for (let i = 3; i < data.length; i++) {
    const row = data[i];
    
    const internalId = row[0];
    const partnerId = row[1];
    const partnerName = row[2];
    const subRegion = row[3];
    const pdm = row[4];
    const type = row[5];
    const profiles = row[6];

    if (!internalId || !partnerName) continue;

    slicerOptions.types.add(type);
    slicerOptions.subRegions.add(subRegion);
    slicerOptions.pdms.add(pdm);

    const scores = [];
    for (let c = 7; c < row.length; c++) {
      if (row[c] !== "" && row[c] !== 0 && row[c] !== "-") {
         
         // Fix merged headers
         let sol = String(headersMap.solution[c]).trim();
         if (sol === "") {
             for (let k = c - 1; k >= 7; k--) {
                 if (String(headersMap.solution[k]).trim() !== "") { sol = String(headersMap.solution[k]).trim(); break; }
             }
         }
         
         let prod = String(headersMap.product[c]).trim();
         if (prod === "") {
             for (let k = c - 1; k >= 7; k--) {
                 if (String(headersMap.product[k]).trim() !== "") { prod = String(headersMap.product[k]).trim(); break; }
             }
         }

         const tier = headersMap.tier[c];

         slicerOptions.solutions.add(sol);
         slicerOptions.products.add(prod);

         scores.push({
           solution: sol,
           product: prod,
           tier: tier,
           count: row[c]
         });
      }
    }

    rows.push({
      internalId,
      partnerId,
      partnerName,
      subRegion,
      pdm,
      type,
      profiles,
      scores
    });
  }

  return {
    success: true,
    data: rows,
    slicers: {
      types: Array.from(slicerOptions.types).sort(),
      subRegions: Array.from(slicerOptions.subRegions).sort(),
      pdms: Array.from(slicerOptions.pdms).sort(),
      solutions: Array.from(slicerOptions.solutions).sort(),
      products: Array.from(slicerOptions.products).sort()
    }
  };
}
