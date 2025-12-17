/**
 * ****************************************
 * Google Apps Script - Selector View (Local Analytics)
 * File: Selector_View.js
 * Description: Generates a consolidated "Selector" view by aggregating local CACHE data.
 *              Replaces the need for complex BQ pivots for the end-user.
 * ****************************************
 */

/**
 * Main execution function for the Selector View.
 */
function runSelectorBuilder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Read Data Sources
  const landscapeSheet = ss.getSheetByName(SHEETS.CACHE_PARTNER_LANDSCAPE);
  const profilesSheet = ss.getSheetByName(SHEETS.CACHE_PROFILES);

  if (!landscapeSheet || !profilesSheet) {
    Logger.log("[Selector] Missing CACHE sheets. Run Landscape & Profiles rebuild first.");
    return;
  }

  const landscapeData = landscapeSheet.getDataRange().getValues();
  const profilesData = profilesSheet.getDataRange().getValues();

  // 2. Index Partner Landscape (Domain -> Metadata)
  // Header: Name(0), Domain(1), Managed(2), EmailTo(3), EmailCC(4), Countries(5+)...
  // We need to know which Countries a partner is active in (from Landscape view)
  const partnerMap = new Map();
  const countryHeaders = landscapeData[0].slice(5); // Columns F onwards are countries

  for (let i = 1; i < landscapeData.length; i++) {
    const row = landscapeData[i];
    const domain = String(row[1]).toLowerCase().trim();
    const name = row[0];
    const isManaged = row[2];

    // Capture active countries
    const activeCountries = [];
    for (let c = 0; c < countryHeaders.length; c++) {
      if (row[5 + c] === true) {
        activeCountries.push(countryHeaders[c]);
      }
    }

    if (domain) {
      partnerMap.set(domain, {
        name: name,
        isManaged: isManaged,
        activeCountries: new Set(activeCountries)
      });
    }
  }

  // 3. Process Profiles & Calculate Tiers
  // Schema: Partner Name(0), ID(1), Country(2), Job(3), Product(4), Score(5), Solution(6)
  // We want to aggregate by Partner + Solution + Product
  // And output: Partner, Domain, Solution, Product, Tier 1 Count, Tier 2 Count, ...

  const aggregation = new Map(); // Key: "Domain|Solution|Product" -> Counts

  // Helper to getKey
  const getAggKey = (dom, sol, prod) => `${dom}|${sol}|${prod}`;

  for (let i = 1; i < profilesData.length; i++) {
    const row = profilesData[i];
    // We need to match Profile Domain to Partner Map
    // Profiles.js now explicitly returns domain in Column 2 (Index 1)

    const pName = row[0];
    const pDomain = String(row[1]).toLowerCase().trim(); // Domain
    const pProfileId = row[2];
    const pCountry = row[3];
    const pJob = row[4];
    const pProduct = row[5]; // Product is Col F (Index 5)
    const pScore = Number(row[6]) || 0; // Score is Col G (Index 6)
    const pSolution = row[7]; // Solution is Col H (Index 7)

    if (!pProduct || !pSolution) continue;

    // Determine Tier
    let tier = 'No Tier';
    if (pScore >= 50) tier = 'Tier 1';
    else if (pScore >= 35) tier = 'Tier 2';
    else if (pScore >= 20) tier = 'Tier 3';
    else tier = 'Tier 4';

    const key = getAggKey(pDomain, pSolution, pProduct); // Key by Domain!

    if (!aggregation.has(key)) {
      aggregation.set(key, {
        name: pName, 
        solution: pSolution,
        product: pProduct,
        t1: 0, t2: 0, t3: 0, t4: 0 
      });
    }

    const entry = aggregation.get(key);
    if (tier === 'Tier 1') entry.t1++;
    if (tier === 'Tier 2') entry.t2++;
    if (tier === 'Tier 3') entry.t3++;
    if (tier === 'Tier 4') entry.t4++;
  }

  // 4. Transform to Flat Table (Exploded by Country)
  // Schema: Country, Partner Name, Solution, Product, Tier 1, Tier 2, Tier 3, Tier 4, Total
  const finalRows = [];

  for (const [key, metrics] of aggregation) {
    // Lookup active countries from Landscape
    // We try to match by Domain first (if we have it), else Name
    // Since we don't have Domain distinct in aggregation yet (unless we parse key), let's use Name from metrics.
    // Ideally we pass Domain through aggregation.
    // Let's assume metrics.name maps to partnerMap values.

    // We need to find the partner in partnerMap
    let activeCountries = ["Unknown"];

    // Inefficient lookup, we should map Name -> Info if Domain is missing
    // But let's try to parse domain from key?
    // Key was: `${dom}|${sol}|${prod}`
    const keyParts = key.split('|');
    const domain = keyParts[0];

    if (partnerMap.has(domain)) {
      const info = partnerMap.get(domain);
      if (info.activeCountries.size > 0) {
        activeCountries = Array.from(info.activeCountries);
      }
    }

    for (const country of activeCountries) {
      finalRows.push([
        country,
        metrics.name,
        metrics.solution,
        metrics.product,
        metrics.t1,
        metrics.t2,
        metrics.t3,
        metrics.t4,
        metrics.t1 + metrics.t2 + metrics.t3 + metrics.t4
      ]);
    }
  }

  // 5. Write to VIEW_Selector
  const viewSheetName = "VIEW_Selector";
  let viewSheet = ss.getSheetByName(viewSheetName);
  if (!viewSheet) {
    viewSheet = ss.insertSheet(viewSheetName);
    viewSheet.setTabColor("ff9900");
  }

  // Clear generic content but careful with Slicers?
  // ensureCleanSheet(viewSheet) might be safer, removing old Slicers manually if needed.
  viewSheet.clear();
  const existingSlicers = viewSheet.getSlicers();
  for (const s of existingSlicers) s.remove();

  const headers = [
    "Residing Country", "Partner Name", "Solution", "Product",
    "Tier 1 (Experts)", "Tier 2", "Tier 3", "Tier 4", "Total Profiles"
  ];

  if (finalRows.length > 0) {
    viewSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#efefef")
      .setFontWeight("bold");

    const dataRange = viewSheet.getRange(2, 1, finalRows.length, headers.length);
    dataRange.setValues(finalRows);

    viewSheet.setFrozenRows(1);
    viewSheet.autoResizeColumns(1, headers.length);

    // 6. Add Slicers (Native UI)
    const wholeRange = viewSheet.getRange(1, 1, finalRows.length + 1, headers.length);

    // Slicer 1: Country (Col 1)
    const slicerCountry = viewSheet.insertSlicer(wholeRange, 2, 1);
    slicerCountry.setPosition(2, 1, 0, 0);
    slicerCountry.setColumnFilterCriteria(1, null); // Clear filters
    slicerCountry.setTitle("Filter by Country");

    // Slicer 2: Solution (Col 3)
    const slicerSol = viewSheet.insertSlicer(wholeRange, 2, 4); // Anchor at Row 2, Col 4 (D)
    slicerSol.setPosition(2, 4, 10, 0);
    slicerSol.setTitle("Filter by Solution");

    // Slicer 3: Product (Col 4)
    const slicerProd = viewSheet.insertSlicer(wholeRange, 2, 7); // Anchor at Row 2, Col 7 (G)
    slicerProd.setPosition(2, 7, 20, 0);
    slicerProd.setTitle("Filter by Product");
  }

  Logger.log(`[Selector] Built with ${finalRows.length} rows and Slicers.`);
}
