/**
 * ============================================================================
 * 📘 PROJECT DOCUMENTATION: PARTNER PERFORMANCE ENGINE (ETL & DASHBOARD)
 * ============================================================================
 * 
 * OVERVIEW:
 * This project is a Hybrid ETL (Extract, Transform, Load) system.
 * 1. EXTRACT: Pulls partner data & scores from BigQuery (`concord-prod`).
 * 2. TRANSFORM: Aggregates data into local "Database Tabs" within this Spreadsheet.
 * 3. LOAD (Decks): Generates/Updates individual Google Sheets for ~150+ partners.
 * 4. LOAD (Dashboard): Powers an interactive "Slicer" dashboard for internal use.
 * 
 * ARCHITECTURE:
 * [BigQuery] <==> [This Spreadsheet (DB Layer)] ==> [Individual Partner Files]
 *                                      ^
 *                                      |
 *                                 [Dashboard Slicer]
 * 
 * ============================================================================
 * 📂 1. KEY FILES & RESPONSIBILITIES
 * ============================================================================
 * 
 * A. CONFIGURATION
 *    - Config.gs: Central hub for Project ID, Spreadsheet IDs, Folder IDs, and Sheet Names.
 *      > CHANGE THIS if you move folders or change the BigQuery project.
 * 
 * B. DATA INGESTION (The SQL Layer)
 *    - LATAM_Partner_DB.gs: 
 *      Fetches the master list of partners, flags (Managed, MCO, etc.), and Profile Counts.
 *      *Concept:* Uses a "Virtual Table" (STRUCT) to inject our specific domain list into SQL.
 *    - Partner_Scoring.gs:
 *      Calculates Tier 1-4 scores per product. Pivots data from Vertical (SQL) to Horizontal (Sheet).
 *    - Profile_DeepDive.gs:
 *      Fetches raw (~50k rows) profile-level data (Who certified in what?) for the "Deep Dive" tabs.
 * 
 * C. FILE GENERATION (The Batch Layer)
 *    - Partner_Individual_Decks.gs:
 *      The core engine. Loops through partners, creates/opens their specific Spreadsheet,
 *      and writes two tabs: "Tier Dashboard" (Matrix) and "Profile Deep Dive" (Pivot).
 *      *Key Feature:* Auto-saves generated links to the Cache System.
 *    - Lock_System.gs:
 *      Protects the generated sheets so only the Admin can edit them, preventing user error.
 * 
 * D. DASHBOARD & INTERFACE
 *    - Partner_Region_Solution_Selector.gs:
 *      Controls the "Partner & Solution Slicer" tab. Handles dropdown logic and table rendering.
 *      *Key Feature:* Uses VLOOKUP formulas for instant hyperlinks (High Performance).
 *    - Link_System.gs:
 *      Maintains a hidden sheet (`System_Link_Cache`) mapping Partner Names -> Drive URLs.
 *    - Menu.gs:
 *      Creates the "🚀 Partner Engine" menu in the UI.
 * 
 * ============================================================================
 * ⚙️ 2. CRITICAL WORKFLOWS (HOW TO RUN)
 * ============================================================================
 * 
 * PHASE 1: DATA REFRESH (Run Weekly/Daily)
 *    1. Run `runBigQueryQuery` (Updates DB Flags).
 *    2. Run `runPartnerScorePivot` (Updates Tiers).
 *    3. Run `runDeepDiveQuerySource` (Updates Profile Data).
 *    *Result:* The tabs in this spreadsheet are now up to date with BigQuery.
 * 
 * PHASE 2: DECK GENERATION (Run after Phase 1)
 *    1. Run `runManagedBatch` (Updates the VIP partners).
 *    2. Run `runUnManagedBatch` (Updates the rest).
 *    *Note:* If script times out (6 mins), simply run it again. It skips finished files.
 * 
 * PHASE 3: MAINTENANCE
 *    1. Run `runLinkUpdateManual`: If the dashboard links look broken or missing.
 *    2. Run `setupDashboard`: If the dropdowns stop working or show old data.
 * 
 * ============================================================================
 * ⚠️ 3. KNOWN CONSTRAINTS & LOGIC NOTES
 * ============================================================================
 * 
 * 1. THE "VIRTUAL TABLE" STRATEGY:
 *    We do not have Write access to BigQuery. To filter for *our* specific partners,
 *    we read the domains from `Consolidate by Partner`, format them into a SQL `STRUCT`,
 *    and inject them into the query as a temporary table (`WITH Spreadsheet_Data AS...`).
 * 
 * 2. EXECUTION TIME LIMITS:
 *    - Consumer Gmail: 6 minutes / execution.
 *    - Workspace (Corp): 30 minutes / execution.
 *    *Mitigation:* The Batch scripts process partners one by one. If it stops, restart it;
 *    it checks if a file exists before creating (Update = Fast, Create = Slow).
 * 
 * 3. HYPERLINKS:
 *    We do NOT use `range.setRichTextValue` for links because it is slow and breaks on Slicer updates.
 *    We use `=HYPERLINK(VLOOKUP(...))` referencing the hidden cache sheet. This is instant.
 * 
 * 4. COLUMN MAPPING (If DB structure changes, update these!):
 *    In `Partner_Individual_Decks.gs`:
 *    - COL_INDEX_MANAGED = 5 (Column F)
 *    - COL_INDEX_BRAZIL  = 7 (Column H)
 *    - COL_INDEX_MCO     = 8 (Column I)
 * 
 * ============================================================================
 * 🛠️ TROUBLESHOOTING
 * ============================================================================
 * 
 * Q: The Slicer is stuck on "UPDATING..."
 * A: An error occurred silently. Run `setupDashboard` from the Menu or Script Editor to reset it.
 * 
 * Q: "No data found" in BigQuery?
 * A: Check `Config.gs`. Ensure `SOURCE_SS_ID` is correct and the tab `Consolidate by Partner`
 *    has domains in Column AH (Index 33/34).
 * 
 * Q: Links in the dashboard are plain text.
 * A: The cache might be empty. Run "🔗 Refresh Links (Manual)" from the menu.
 * 
 * ============================================================================
 */

function showDocumentation() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Documentation", 
           "Please open the '_Documentation.gs' file in the Apps Script Editor to read the technical documentation.", 
           ui.ButtonSet.OK);
}