/**
 * ****************************************
 * Google Apps Script - Selector View (Pivoted Analytics)
 * File: Selector_View.js
 * Description: Generates a consolidated "Selector" view in a Pivot-style grid.
 *              Rows: Partners (with Metadata)
 *              Columns: Solutions -> Products -> Tiers (1-4)
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
  const partnerMap = new Map();
  const countryHeaders = landscapeData[0].slice(5);

  for (let i = 1; i < landscapeData.length; i++) {
    const row = landscapeData[i];
    const domain = String(row[1]).toLowerCase().trim();
    const name = row[0];
    const isManaged = row[2];
    const emailTo = row[3];
    const emailCC = row[4];

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
        emailTo: emailTo,
        emailCC: emailCC,
        activeCountries: new Set(activeCountries)
      });
    }
  }

  // 3. Process Profiles & Build Hierarchy
  // We need a complete list of Solutions and Products to build the columns.
  // And we need aggregated data per Partner.

  const hierarchy = new Map(); // Solution -> Set<Product>
  const partnerData = new Map(); // Domain -> Map<ProductKey, {t1, t2, t3, t4}>

  for (let i = 1; i < profilesData.length; i++) {
    const row = profilesData[i];
    const domain = String(row[1]).toLowerCase().trim();
    const product = row[5];
    const solution = row[7];
    const score = Number(row[6]) || 0;

    if (!product || !solution) continue;

    // Update Hierarchy
    if (!hierarchy.has(solution)) {
      hierarchy.set(solution, new Set());
    }
    hierarchy.get(solution).add(product);

    // Update Partner Data
    if (!partnerData.has(domain)) {
      partnerData.set(domain, new Map());
    }
    const pMap = partnerData.get(domain);

    // Key by Solution|Product to be safe, or just Product if unique? 
    // Safest is Solution|Product
    const prodKey = `${solution}|${product}`;

    if (!pMap.has(prodKey)) {
      pMap.set(prodKey, { t1: 0, t2: 0, t3: 0, t4: 0 });
    }
    const entry = pMap.get(prodKey);

    // Calc Tier
    if (score >= 50) entry.t1++;
    else if (score >= 35) entry.t2++;
    else if (score >= 20) entry.t3++;
    else entry.t4++;
  }

  // Sort Hierarchy
  const sortedSolutions = Array.from(hierarchy.keys()).sort();
  const solutionCols = []; // { solution, products: [] }

  for (const sol of sortedSolutions) {
    const prods = Array.from(hierarchy.get(sol)).sort();
    solutionCols.push({
      name: sol,
      products: prods
    });
  }

  // 4. Build Output Table

  // -- Headers --
  // Row 1: Solution Headers (Merged)
  // Row 2: Product Headers (Merged)
  // Row 3: Tier Headers (Repeated)

  // Fixed Columns: Country, Partner Name, Domain, Managed, Email To, Email CC
  const fixedHeadersLength = 6;
  const fixedHeaders = ["Country", "Partner Name", "Domain", "Managed", "Email To", "Email CC"];

  const row1 = [...fixedHeaders]; // Solutions
  const row2 = [...fixedHeaders.map(() => "")]; // Products
  const row3 = [...fixedHeaders.map(() => "")]; // Tiers

  // Fill Headers
  let colIndex = fixedHeadersLength;
  const merges = []; // { row, col, numRows, numCols }
  const bgColors = []; // Store colors for merging steps if needed, or apply later.

  // Helper to get random distinct pastel color for solutions
  // Fixed set of colors for consistency?
  const solColors = ["#E6E6FA", "#F0F8FF", "#F5F5DC", "#FFE4E1", "#E0FFFF", "#FAFAD2"];
  let cIdx = 0;

  for (const solObj of solutionCols) {
    const startCol = colIndex;
    const products = solObj.products;
    const solWidth = products.length * 4; // 4 Tiers per product

    // Solution Header
    row1[startCol] = solObj.name;
    for (let k = 1; k < solWidth; k++) row1.push(""); // Fill void for merge
    merges.push({ row: 6, col: startCol + 1, numRows: 1, numCols: solWidth, color: solColors[cIdx % solColors.length] });
    cIdx++;

    for (const prod of products) {
      const pStartCol = colIndex;
      // Product Header
      row2[pStartCol] = prod;
      for (let k = 1; k < 4; k++) row2.push(""); // Fill void
      merges.push({ row: 7, col: pStartCol + 1, numRows: 1, numCols: 4, color: "#FFFFFF" }); // White or lighter shade?

      // Tier Header
      row3[colIndex] = "Tier 1";
      row3[colIndex + 1] = "Tier 2";
      row3[colIndex + 2] = "Tier 3";
      row3[colIndex + 3] = "Tier 4";

      colIndex += 4;
    }
  }

  // -- Body Rows --
  // Iterate Partner Map, but we need to pivot by Country?
  // Previous view was Exploded by Country. Pivot tables usually list Unique IDs.
  // If a partner is in multiple countries, do we list them multiple times or once?
  // User asked for "like image 2". Image 2 doesn't show rows clearly.
  // Assuming "Selector" typically aims to select a partner for a country.
  // I will explod by country again to be safe (Row = Country + Partner).

  const bodyRows = [];

  // We need to iterate over all partners we found in Landscape (or Profiles? Landscape is master).
  // Profiles might have partners NOT in Landscape? Unlikely if integrity is kept.
  // We'll iterate partnerMap.

  for (const [domain, meta] of partnerMap) {
    const pData = partnerData.get(domain);
    // If no profile data, they have all zeros.

    // Default country "Unknown" if set empty
    const countries = meta.activeCountries.size > 0 ? Array.from(meta.activeCountries) : ["Unknown"];

    for (const country of countries) {
      const row = [
        country,
        meta.name,
        domain,
        meta.isManaged,
        meta.emailTo,
        meta.emailCC
      ];

      // Add metrics
      for (const solObj of solutionCols) {
        for (const prod of solObj.products) {
          const key = `${solObj.name}|${prod}`;
          const entry = pData ? pData.get(key) : null; // Check if pData exists before getting entry
          if (entry) {
            row.push(entry.t1 || 0); // show 0? or hyphen?
            row.push(entry.t2 || 0);
            row.push(entry.t3 || 0);
            row.push(entry.t4 || 0);
          } else {
            row.push(0, 0, 0, 0); // Zeros
          }
        }
      }
      bodyRows.push(row);
    }
  }

  // 5. Write to Sheet
  const viewSheetName = "VIEW_Selector";
  let viewSheet = ss.getSheetByName(viewSheetName);
  if (!viewSheet) {
    viewSheet = ss.insertSheet(viewSheetName);
    viewSheet.setTabColor("ff9900");
  }

  viewSheet.clear();
  const existingSlicers = viewSheet.getSlicers();
  for (const s of existingSlicers) s.remove();

  const startRow = 6;

  if (bodyRows.length > 0) {
    // Write Headers
    // Row 6 (Sol), Row 7 (Prod), Row 8 (Tier)
    const headerRange = viewSheet.getRange(startRow, 1, 3, row1.length);
    headerRange.setValues([row1, row2, row3]);

    // Apply Merges & Colors
    for (const m of merges) {
      viewSheet.getRange(m.row, m.col, m.numRows, m.numCols).merge().setBackground(m.color).setHorizontalAlignment("center").setFontWeight("bold");
    }

    // Style Tier Row
    viewSheet.getRange(startRow + 2, fixedHeadersLength + 1, 1, row1.length - fixedHeadersLength)
      .setBackground("#f3f3f3")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true);

    // Style Fixed Headers
    viewSheet.getRange(startRow, 1, 3, fixedHeadersLength)
      .setBackground("#e0e0e0")
      .setFontWeight("bold")
      .setVerticalAlignment("middle");

    // Write Data
    const dataRange = viewSheet.getRange(startRow + 3, 1, bodyRows.length, row1.length);
    dataRange.setValues(bodyRows);

    // Formatting Data
    // Zeros as hyphens? optional.

    // Freeze
    viewSheet.setFrozenRows(startRow + 2);
    viewSheet.setFrozenColumns(2); // Freeze Country and Partner Name

    // Auto resize
    viewSheet.autoResizeColumns(1, fixedHeadersLength);
    // Fixed width for metric columns to save space?
    viewSheet.setColumnWidths(fixedHeadersLength + 1, row1.length - fixedHeadersLength, 50);

    SpreadsheetApp.flush();

    // 6. Add Slicers for Fixed Columns
    const wholeRange = viewSheet.getRange(startRow + 2, 1, bodyRows.length + 1, row1.length);
    const defaultCriteria = SpreadsheetApp.newFilterCriteria().build();

    // Country
    const s1 = viewSheet.insertSlicer(wholeRange, 1, 1);
    s1.setTitle("Filter Country");
    s1.setPosition(2, 1, 0, 0);

    // Partner
    const s2 = viewSheet.insertSlicer(wholeRange, 2, 1);
    s2.setTitle("Filter Partner");
    s2.setPosition(2, 3, 0, 0);

    // Managed
    const s3 = viewSheet.insertSlicer(wholeRange, 4, 1);
    s3.setTitle("Filter Managed");
    s3.setPosition(2, 5, 0, 0);

  }

  // Workaround redraw
  const tempSheet = ss.getSheetByName(SHEETS.CACHE_PARTNER_LANDSCAPE);
  if (tempSheet) {
    tempSheet.activate();
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    viewSheet.activate();
  }

  Logger.log(`[Selector] Built Pivot View with ${bodyRows.length} rows.`);
}
