/*
# Operations Guide - DRP 2026

## Manual Operations (Files and Functions)

To perform operations manually in the Google Apps Script editor, select the specified file and run the corresponding function:

### 1. Update the Data
Both of these functions are independent and read from the main DB sheet (`LATAM_Partner_DB_2026`) to query BigQuery. You can run them in any order depending on what you need to update:

- **Option A: Update Scoring Pivot**
  - File: `Partner_Scoring_2026.js`
  - Function: `runPartnerScorePivot2026()`
  - Description: Runs the BigQuery query to generate the scoring pivot sheet.

- **Option B: Update Profile Deep Dive**
  - File: `DeepDive_2026.js`
  - Function: `runDeepDive2026()`
  - Description: Runs the BigQuery query to generate the detailed profile deep dive sheet.

### 2. Recreate / Refresh Dashboard
- File: `Partner_Region_Solution_Selector_2026.js`
- Function: `setupDashboard2026()` (to initialize/recreate) or `refreshDashboardData2026(dashSheet)` (to update data).
- Description: Updates or initializes the dashboard view based on the latest data.

### 3. Update Partner Files (Decks)
- File: `Partner_Individual_Decks_2026.js`
- Function: `runFullBatchDecks2026()`
- Description: Generates or updates individual decks for all partners in batch.

### 4. Send Emails
- File: `Partner_Summary_Email_2026.js`
- Function: `runBatchEmailSender2026()`
- Description: Sends summary emails to partners in batch.

---

## Workflow Explanation

### Clasp Setup and Sync
We set up clasp (Google Apps Script CLI) to allow working on files locally and syncing them with the remote Google Apps Script project.
- Usage: Due to permission issues with global installation on this machine, we use npx @google/clasp to run commands without needing a global install.
- Login: You completed the login manually in your terminal using npx @google/clasp login.
- Sync: We use npx @google/clasp pull to fetch files and npx @google/clasp push to send local changes to the remote project.

### Directory Cleanup
To keep the project focused on the 2026 version:
- We moved all files that do not end with 2026.js to an archive_old folder locally to keep them as backup.
- We kept appsscript.json and .clasp.json in the root as they are required for project function.
- Important: clasp does not automatically delete remote files on push when they are deleted locally. To remove the old files from the Apps Script web editor, you must delete them manually there.
*/
