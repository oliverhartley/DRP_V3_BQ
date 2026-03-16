/**
 * ****************************************
 * Google Apps Script - Partner Summary Email 2026
 * File: Partner_Summary_Email_2026.js
 * Description: Generates an executive summary using Gemini and sends it via email.
 * Iterates over LATAM_Partner_DB_2026 and aggregates by Partner Name.
 * ****************************************
 */

const BATCH_EMAIL_TIME_LIMIT_MS_2026 = 1200000; // 20 minutes

function runBatchEmailSender2026() {
  const startTime = new Date().getTime();
  const currentBatchId = getBatchId2026(); 
  Logger.log(`>>> STARTING 2026 BATCH EMAIL PROCESS [Batch ID: ${currentBatchId}] <<<`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAME_2026);
  
  if (!dbSheet) {
    Logger.log(`ERROR: Sheet ${SHEET_NAME_2026} not found.`);
    return;
  }

  const dataRange = dbSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());

  // Dynamically find columns
  const colPartnerName = 0; // Partner Name is A (0)
  
  let colToEmail = -1;
  let colCcEmail = -1;
  let colStatus = -1;
  let colSpreadsheetId = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h.includes("to") && h.includes("email")) colToEmail = c;
    if (h.includes("cc") && h.includes("email")) colCcEmail = c;
    if (h.includes("email") && h.includes("sent")) colStatus = c;
    if (h.includes("spreadsheet") && h.includes("id")) colSpreadsheetId = c;
  }

  if (colToEmail === -1 || colCcEmail === -1 || colSpreadsheetId === -1 || colStatus === -1) {
      Logger.log("ERROR: Could not find one or more required columns (To Email, CC Email, Spreadsheet ID, Status). Please ensure they exist.");
      return;
  }

  // Group by Partner Name
  const partnerMap = new Map();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partnerName = String(row[colPartnerName]).trim();
    if (!partnerName) continue;

    if (!partnerMap.has(partnerName)) {
      partnerMap.set(partnerName, {
        rows: [],
        toEmails: new Set(),
        ccEmails: new Set(),
        status: "",
        spreadsheetId: ""
      });
    }

    const pData = partnerMap.get(partnerName);
    pData.rows.push(i); // Store 0-indexed row number for updating status
    
    // Default to the first found status and ID for the group
    if (!pData.status && row[colStatus]) pData.status = String(row[colStatus]).trim();
    if (!pData.spreadsheetId && row[colSpreadsheetId]) pData.spreadsheetId = String(row[colSpreadsheetId]).trim();

    // Collect all unique emails for this partner
    const to = String(row[colToEmail] || "");
    const cc = String(row[colCcEmail] || "");
    
    if (to) to.split(',').forEach(e => pData.toEmails.add(e.trim()));
    if (cc) cc.split(',').forEach(e => pData.ccEmails.add(e.trim()));
  }

  let processedCount = 0;
  let skippedCount = 0;

  for (const [partnerName, pData] of partnerMap.entries()) {
    if (isTimeLimitApproaching2026(startTime)) {
      Logger.log("WARNING: Time limit approaching. Stopping to allow safe resume on next trigger.");
      break;
    }

    if (pData.ccEmails.size === 0 && pData.toEmails.size === 0) {
      Logger.log(`Skipping ${partnerName} - No emails found.`);
      continue;
    }

    if (pData.status === currentBatchId) {
      Logger.log(`Skipping ${partnerName} - Already processed for this batch.`);
      skippedCount++;
      continue;
    }

    if (!pData.spreadsheetId) {
      Logger.log(`Skipping ${partnerName} - No Spreadsheet ID found. Please generate their deck first.`);
      continue;
    }

    Logger.log(`Processing Partner: ${partnerName}...`);
    
    const toEmailStr = Array.from(pData.toEmails).join(",");
    const ccEmailStr = Array.from(pData.ccEmails).join(",");

    try {
      generateAndSendPartnerSummary2026(partnerName, pData.spreadsheetId, toEmailStr, ccEmailStr);

      // Update Status for all rows belonging to this partner
      pData.rows.forEach(r => {
        dbSheet.getRange(r + 1, colStatus + 1).setValue(currentBatchId);
      });
      SpreadsheetApp.flush();
      processedCount++;

      Utilities.sleep(5000); // Respect Gemini quotas
    } catch (e) {
      Logger.log(`  ERROR processing ${partnerName}: ${e.toString()}`);
    }
  }

  Logger.log(`>>> 2026 BATCH RUN COMPLETE. Sent: ${processedCount}, Skipped: ${skippedCount} <<<`);
}

function getBatchId2026() {
  const now = new Date();
  const shiftedDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const year = shiftedDate.getFullYear();
  const onejan = new Date(year, 0, 1);
  const week = Math.ceil((((shiftedDate.getTime() - onejan.getTime()) / 86400000) + onejan.getDay() + 1) / 7);
  return `SENT_2026_${year}_${week}`;
}

function isTimeLimitApproaching2026(startTime) {
  return (new Date().getTime() - startTime) > BATCH_EMAIL_TIME_LIMIT_MS_2026;
}

function generateAndSendPartnerSummary2026(partnerName, ssId, toEmails, ccEmails) {
  Logger.log(`  Generating 2026 summary for ${partnerName}...`);

  const sheetData = getPartnerSheetData2026(ssId);
  if (!sheetData) {
    Logger.log("  ERROR: Failed to retrieve sheet data.");
    return;
  }
  
  const fullPrompt = `
    You are an expert Data Analyst and Executive Assistant.
    Please analyze the following 2026 data for partner: "${partnerName}".
    
    Data from "Tier Dashboard":
    ${sheetData.tierDashboard}
    
    Data from "Profile Deep Dive":
    ${sheetData.profileDeepDive}
    
    Task:
    Create a comprehensive Email Report containing TWO SECTIONS:
    
    SECTION 1: VISUAL EXECUTIVE DASHBOARD (The "Infographic")
    - Start with this EXACT greeting: "Hola ${partnerName},<br><br>Aquí su informe semanal del DRP Status para su análisis. Cualquier duda puedes contactar al equipo de Partner (copiado en este correo)."
    - This must be a graphical representation using ONLY HTML/CSS (Files, Tables, Divs).
    - Do NOT use images or external charts. Use HTML/CSS to create "Bar Charts" and "Scorecards".
    - Layout:
        - **Header**: Partner Name & "Readiness Snapshot".
        - **KPI Row**: 3 Cards showing (Total Profiles, Top Solution, Readiness Score/Tier 1 Count).
        - **Strengths Chart**: A Visual List simulating a Bar Chart (e.g., <div style="width: 80%; background: #4285f4; height: 10px;"></div>) for Tier 1 counts by Solution.
        - **Upskilling Gaps**: A Table showing "Beginner Count" vs "Target".
        - **Top Talent**: A clean table of the top 3-5 individuals.
    - Style: Use Google Brand colors (Blue #4285f4, Red #ea4335, Yellow #fbbc04, Green #34a853). Use Grey #f1f3f4 for backgrounds.
    
    SECTION 2: DETAILED EXECUTIVE SUMMARY
    - Written narrative explaining the data.
    - Tiers Definitions:
      - Tier 1: Delivery Ready (Expert).
      - Tier 2: Intermediate.
      - Tier 3: Beginner-Intermediate.
      - Tier 4: Beginner.
    - Sections: "Key Strengths", "Critical Gaps", "Recommendations".
    
    Output Format:
    Return ONE block of clean, professional HTML.
    - Use Inline CSS for everything (Gmail compatible).
    - Make it look premium (padding, border-radius, shadows).
  `;

  const finalHtml = callGeminiWithFallback2026(fullPrompt);
  if (!finalHtml) {
    Logger.log("  ERROR: Failed to generate summary from Gemini.");
    return;
  }

  const subject = `[GCP DRP Readiness 2026] Partner Executive Summary: ${partnerName}`;
  const fileUrl = `https://docs.google.com/spreadsheets/d/${ssId}/edit`;
  
  let cleanHtml = finalHtml.replace(/```html/g, "").replace(/```/g, "").trim();

  const emailBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: 0 auto;">
      ${cleanHtml}
      <br><br>
      <hr>
      <p style="text-align: center;">
        <a href="${fileUrl}" style="background-color: #4285f4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">
          Open 2026 Partner Dashboard
        </a>
      </p>
      <p style="font-size: 12px; color: #666; text-align: center;">
        Link to file: <a href="${fileUrl}">${fileUrl}</a>
      </p>
      
      <!-- Footer -->
      <br>
      <div style="text-align: center; color: #999; font-size: 11px; margin-top: 20px;">
        <p>&copy; 2026 Google Cloud Partner Team. Confidential.</p>
        <p style="font-style: italic;">
          This summary was generated by Gemini. Any imprecision, please let the team know.
        </p>
      </div>
    </div>
  `;

  sendEmail2026(subject, emailBody, toEmails, ccEmails);
}

function getPartnerSheetData2026(ssId) {
  try {
    const ss = SpreadsheetApp.openById(ssId);

    const tierSheet = ss.getSheetByName("Tier Dashboard");
    const deepDiveSheet = ss.getSheetByName("Profile Deep Dive");

    if (!tierSheet || !deepDiveSheet) {
      Logger.log("ERROR: Missing required sheets in deck.");
      return null;
    }

    const tierData = tierSheet.getDataRange().getValues().map(row => row.join(", ")).join("\\n");
    const deepDiveData = deepDiveSheet.getRange(1, 1, Math.min(deepDiveSheet.getLastRow(), 200), deepDiveSheet.getLastColumn()).getValues().map(row => row.join(", ")).join("\\n");

    return {
      tierDashboard: tierData,
      profileDeepDive: deepDiveData
    };
  } catch (e) {
    Logger.log(`Error reading sheets for deck ${ssId}: ${e.toString()}`);
    return null;
  }
}

function callGeminiWithFallback2026(prompt) {
  const userModels = [
    { name: 'gemini-3-flash-preview', version: 'v1beta' }
  ];

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties.");
    return null;
  }

  for (const model of userModels) {
    try {
      const url = `https://generativelanguage.googleapis.com/${model.version}/models/${model.name}:generateContent?key=${apiKey}`;

      const payload = {
        contents: [{
          parts: [{ text: prompt }]
        }]
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        const json = JSON.parse(responseText);
        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts) {
          return json.candidates[0].content.parts[0].text;
        }
      } else {
        Logger.log(`FAILED: Model ${model.name} returned code ${responseCode}. Response: ${responseText}`);
      }
    } catch (e) {
      Logger.log(`EXCEPTION: Model ${model.name} failed with error: ${e.toString()}`);
    }
  }
  return null;
}

function sendEmail2026(subject, htmlBody, to, cc) {
  try {
    const emailOptions = {
      to: to,
      subject: subject,
      htmlBody: htmlBody
    };

    if (cc && String(cc).trim() !== "") {
      emailOptions.cc = cc;
    }

    if (!to || String(to).trim() === "") {
      Logger.log("  WARNING: 'TO' email is empty. Attempting to send using CC only if possible, or aborting.");
      if (emailOptions.cc) {
        // If TO is empty, put CCs in the TO field but LEAVE them in CC as well for visibility
        emailOptions.to = emailOptions.cc;
      } else {
        Logger.log("  ERROR: No recipients defined. Skipping email.");
        return;
      }
    }

    MailApp.sendEmail(emailOptions);
    Logger.log(`  Email sent to: ${emailOptions.to} (CC: ${cc || 'None'})`);
  } catch (e) {
    Logger.log(`  Error sending email: ${e.toString()}`);
  }
}
