/**
 * ****************************************
 * Google Apps Script - Partner Summary Email
 * File: Partner_Summary_Email.gs
 * Description: Generates an executive summary using Gemini and sends it via email.
 * ****************************************
 */

const TARGET_PARTNER_SS_ID = "12I1uGum5FnT2nxdFlBanTRC4vAdZuVx6u16sWIc9FSA";
const EMAIL_RECIPIENT = "oliverhartley@google.com";

function sendPartnerSummaryEmail() {
  Logger.log(">>> STARTING PARTNER SUMMARY EMAIL PROCESS <<<");

  // 1. Get Data from Sheets
  const sheetData = getPartnerSheetData(TARGET_PARTNER_SS_ID);
  if (!sheetData) {
    Logger.log("ERROR: Failed to retrieve sheet data.");
    return;
  }
  
  // 2. Prepare Prompt for Gemini
  const prompt = `
    You are an expert Data Analyst and Executive Assistant.
    Please analyze the following data from a Partner Dashboard and a Profile Deep Dive.
    
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
    Logger.log("ERROR: Failed to generate summary from Gemini.");
    return;
  }

  // 4. Send Email
  const subject = "[GCP DRP Readiness] Partner Executive Summary";
  const fileUrl = `https://docs.google.com/spreadsheets/d/${TARGET_PARTNER_SS_ID}/edit`;
  
  // Clean up any potential markdown code blocks if Gemini wraps HTML in ```html ... ```
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

  sendEmail(subject, emailBody);

  Logger.log(">>> PROCESS COMPLETE <<<");
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

    // For Deep Dive, we might want to limit rows if it's huge, but let's try grabbing it all first or the pivot table part
    // The user mentioned "Profile Deep Dive" has a hidden raw data section at row 1000, but maybe we just want the visible part?
    // Let's grab the visible part (top 100 rows maybe?) or the whole thing if small.
    // Given the previous script, the visible part is the "Profile Details" table starting at row 6.
    // Let's grab the first 200 rows to be safe.
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
  const models = [
    { name: 'gemini-1.5-pro', version: 'v1beta' }, // Trying 1.5 Pro first as it's usually best for analysis
    { name: 'gemini-1.5-flash', version: 'v1beta' },
    { name: 'gemini-pro', version: 'v1' } // Fallback to older if needed, though 1.5 is standard now
  ];

  // User requested specific list:
  // { name: 'gemini-3-pro-preview', version: 'v1beta' },
  // { name: 'gemini-1.5-pro', version: 'v1' },
  // { name: 'gemini-1.5-flash', version: 'v1' }

  const userModels = [
    { name: 'gemini-3-pro-preview', version: 'v1beta' }
  ];

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties.");
    return null;
  }

  for (const model of userModels) {
    Logger.log(`Attempting to call model: ${model.name} (${model.version})...`);
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
          Logger.log(`SUCCESS: Model ${model.name} generated content.`);
          return json.candidates[0].content.parts[0].text;
        }
      } else {
        Logger.log(`FAILED: Model ${model.name} returned code ${responseCode}. Response: ${responseText}`);
      }
    } catch (e) {
      Logger.log(`EXCEPTION: Model ${model.name} failed with error: ${e.toString()}`);
    }
  }

  Logger.log("ALL MODELS FAILED. Please check API Key or Quota.");
  return null;
}

function sendEmail(subject, htmlBody) {
  try {
    MailApp.sendEmail({
      to: EMAIL_RECIPIENT,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log(`Email sent to ${EMAIL_RECIPIENT}`);
  } catch (e) {
    Logger.log(`Error sending email: ${e.toString()}`);
  }
}
