/**
 * ****************************************
 * Google Apps Script - Partner Summary Email
 * File: Partner_Summary_Email.gs
 * Description: Generates an executive summary using Gemini and sends it via email.
 * Includes Batch Processing capabilities.
 * ****************************************
 */

// NOTE: Uses Global Constants from Config.gs
// SOURCE_SS_ID, PARTNER_FOLDER_ID

// --- BATCH CONFIGURATION ---
const COL_PARTNER_NAME = 33; // Column AH
const COL_TO_EMAIL = 35;     // Column AJ
const COL_CC_EMAIL = 36;     // Column AK
const COL_STATUS = 37;       // Column AL (New: For Status Tracking)
const MAX_EXECUTION_TIME_MS = 1200000; // 20 minutes (Workspace Account Limit support)

function runBatchEmailSender() {
  const startTime = new Date().getTime();
  const currentBatchId = getBatchId(); // Format: SENT_YYYY_WW
  Logger.log(`>>> STARTING BATCH EMAIL PROCESS [Batch ID: ${currentBatchId}] <<<`);

  const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sheet = ss.getSheetByName("Copy of Consolidate by Partner");
  if (!sheet) {
    Logger.log("ERROR: 'Copy of Consolidate by Partner' sheet not found in Source SS.");
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  let processedCount = 0;
  let skippedCount = 0;

  for (let i = 1; i < data.length; i++) {
    // Check time limit before starting next row
    if (isTimeLimitApproaching(startTime)) {
      Logger.log("WARNING: Time limit approaching. Stopping to allow safe resume on next trigger.");
      break;
    }

    const row = data[i];
    const partnerName = row[COL_PARTNER_NAME];
    const toEmails = row[COL_TO_EMAIL];
    const ccEmails = row[COL_CC_EMAIL];
    const currentStatus = row[COL_STATUS];

    // Trigger condition:
    // 1. Column AK (CC) is NOT empty (Original requirement)
    // 2. Status != currentBatchId (New: Avoid duplicates for this week)
    if (ccEmails && String(ccEmails).trim() !== "") {

      if (currentStatus === currentBatchId) {
        Logger.log(`[Row ${i + 1}] Skipping ${partnerName} (Already processed for this batch).`);
        skippedCount++;
        continue;
      }

      Logger.log(`[Row ${i + 1}] Processing Partner: ${partnerName}`);

      const fileId = findPartnerFileId(partnerName);
      if (fileId) {
        Logger.log(`  File Found: ${fileId}`);
        try {
          generateAndSendPartnerSummary(partnerName, fileId, toEmails, ccEmails);

          // Update Status to Current Batch ID
          sheet.getRange(i + 1, COL_STATUS + 1).setValue(currentBatchId);
          processedCount++;

          // Respect Gemini quotas
          Utilities.sleep(5000); 
        } catch (e) {
          Logger.log(`  ERROR processing ${partnerName}: ${e.toString()}`);
        }
      } else {
        Logger.log(`  WARNING: Partner file not found for ${partnerName}`);
      }
    }
  }
  Logger.log(`>>> BATCH RUN COMPLETE. Sent: ${processedCount}, Skipped: ${skippedCount} <<<`);
}

function getBatchId() {
  const now = new Date();
  const year = now.getFullYear();
  // Simple Week Number Calculation
  const onejan = new Date(year, 0, 1);
  const week = Math.ceil((((now.getTime() - onejan.getTime()) / 86400000) + onejan.getDay() + 1) / 7);
  return `SENT_${year}_${week}`;
}

function isTimeLimitApproaching(startTime) {
  return (new Date().getTime() - startTime) > MAX_EXECUTION_TIME_MS;
}

function findPartnerFileId(partnerName) {
  try {
    const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID);
    // Naming convention: "{Partner Name} - Partner Dashboard"
    const fileName = `${partnerName} - Partner Dashboard`;
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      return files.next().getId();
    }
  } catch (e) {
    Logger.log(`Error searching for file: ${e.toString()}`);
  }
  return null;
}

function generateAndSendPartnerSummary(partnerName, ssId, toEmails, ccEmails) {
  Logger.log(`  Generating summary for ${partnerName}...`);

  // 1. Get Data from Sheets
  const sheetData = getPartnerSheetData(ssId);
  if (!sheetData) {
    Logger.log("  ERROR: Failed to retrieve sheet data.");
    return;
  }
  
  // 2. Prepare Prompt for Gemini
  const prompt = `
    You are an expert Data Analyst and Executive Assistant.
    Please analyze the following data from a Partner Dashboard and a Profile Deep Dive for partner: "${partnerName}".
    
    Data from "Tier Dashboard":
    ${sheetData.tierDashboard}
    
    Data from "Profile Deep Dive":
    ${sheetData.profileDeepDive}
    
    Task:
    Write a concise Executive Summary of this partner's readiness and profile status.
    
    IMPORTANT: Use the following definitions for Tiers in your analysis:
    - Tier 1: Delivery Ready (Practitioner has considerable technical capabilities on Google Cloud).
    - Tier 2: Intermediate (Practitioner can become delivery ready through certifications, challenge labs, and training).
    - Tier 3: Beginner to Intermediate (Practitioner can become delivery ready through certifications, challenge labs, and training).
    - Tier 4: Beginner (Practitioner is just starting out on Google Cloud).

    Output Format:
    Provide the response in clean, professional HTML format suitable for an email body.
    - Use <h2> for section headers.
    - Use <ul> and <li> for lists.
    - Use <b> for emphasis.
    - Do NOT use markdown (like ** or ##).
    - Do NOT include the subject line in the HTML body (I will add it separately).
    - Style the HTML to be readable and professional (e.g., standard fonts).
  `;

  // 3. Call Gemini
  const summaryHtml = callGeminiWithFallback(prompt);
  if (!summaryHtml) {
    Logger.log("  ERROR: Failed to generate summary from Gemini.");
    return;
  }

  // 4. Send Email
  const subject = `[GCP DRP Readiness] Partner Executive Summary: ${partnerName}`;
  const fileUrl = `https://docs.google.com/spreadsheets/d/${ssId}/edit`;
  
  // Clean up any potential markdown code blocks
  let cleanHtml = summaryHtml.replace(/```html/g, "").replace(/```/g, "").trim();

  const emailBody = `
    <div style="font-family: Arial, sans-serif; color: #333;">
      ${cleanHtml}
      <br><br>
      <hr>
      <p>
        <a href="${fileUrl}" style="background-color: #4285f4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;">
          Open Partner Dashboard
        </a>
      </p>
      <p style="font-size: 12px; color: #666;">
        Link to file: <a href="${fileUrl}">${fileUrl}</a>
      </p>
    </div>
  `;

  sendEmail(subject, emailBody, toEmails, ccEmails);
}

function getPartnerSheetData(ssId) {
  try {
    const ss = SpreadsheetApp.openById(ssId);

    const tierSheet = ss.getSheetByName("Tier Dashboard");
    const deepDiveSheet = ss.getSheetByName("Profile Deep Dive");

    if (!tierSheet || !deepDiveSheet) {
      Logger.log("ERROR: Missing required sheets.");
      return null;
    }

    // Get all data as text (simplified for token limit, can be optimized)
    const tierData = tierSheet.getDataRange().getValues().map(row => row.join(", ")).join("\n");

    // For Deep Dive, limit to reasonable amount
    const deepDiveData = deepDiveSheet.getRange(1, 1, Math.min(deepDiveSheet.getLastRow(), 200), deepDiveSheet.getLastColumn()).getValues().map(row => row.join(", ")).join("\n");

    return {
      tierDashboard: tierData,
      profileDeepDive: deepDiveData
    };
  } catch (e) {
    Logger.log(`Error reading sheets: ${e.toString()}`);
    return null;
  }
}

function callGeminiWithFallback(prompt) {
  const userModels = [
    { name: 'gemini-3-pro-preview', version: 'v1beta' }
  ];

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties.");
    return null;
  }

  for (const model of userModels) {
    // Logger.log(`Attempting to call model: ${model.name}...`); // Reduced logging for batch
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

function sendEmail(subject, htmlBody, to, cc) {
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
      // MailApp might fail if 'to' is empty. Let's try to just log error if TO is missing.
      if (emailOptions.cc) {
        Logger.log("  Using CC address as TO since TO is empty (Not recommended but trying).");
        emailOptions.to = emailOptions.cc;
        delete emailOptions.cc;
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
