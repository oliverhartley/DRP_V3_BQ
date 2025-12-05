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
    Highlight key strengths, gaps, and any interesting findings from the deep dive.
    The summary should be professional and suitable for an email body.
    Format it with clear sections or bullet points.
  `;

  // 3. Call Gemini
  const summary = callGeminiWithFallback(prompt);
  if (!summary) {
    Logger.log("ERROR: Failed to generate summary from Gemini.");
    return;
  }

  // 4. Send Email
  const subject = "[GCP DRP Readiness] Partner Executive Summary";
  const fileUrl = `https://docs.google.com/spreadsheets/d/${TARGET_PARTNER_SS_ID}/edit`;
  const emailBody = `${summary}\n\nLink to Partner File: ${fileUrl}`;
  
  sendEmail(subject, emailBody, fileUrl);
  
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
    { name: 'gemini-2.0-flash-exp', version: 'v1beta' }, // "gemini-3-pro-preview" might be a typo for 2.0 or 1.5 pro preview, but let's stick to what they asked if valid, or standard ones. 
    // Actually, "gemini-3-pro-preview" doesn't exist publicly yet. I'll use what they asked but add a note or fallback to known working ones if it fails.
    // Wait, the user specifically asked for:
    // { name: 'gemini-3-pro-preview', version: 'v1beta' },
    // { name: 'gemini-1.5-pro', version: 'v1' },
    // { name: 'gemini-1.5-flash', version: 'v1' }
    
    // I will use EXACTLY what they asked, but I suspect 'gemini-3-pro-preview' might fail if it's not real yet.
    // I will add 'gemini-1.5-pro-latest' or similar if needed, but let's trust the user has access or is testing.
    // Actually, I'll stick to their exact list.
    { name: 'gemini-3-pro-preview', version: 'v1beta' },
    { name: 'gemini-1.5-pro', version: 'v1' },
    { name: 'gemini-1.5-flash', version: 'v1' }
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

function sendEmail(subject, body, link) {
  try {
    MailApp.sendEmail({
      to: EMAIL_RECIPIENT,
      subject: subject,
      body: body // Plain text body
      // htmlBody: body.replace(/\n/g, '<br>') + `<br><br><a href="${link}">Link to Partner File</a>` // Optional HTML version
    });
    Logger.log(`Email sent to ${EMAIL_RECIPIENT}`);
  } catch (e) {
    Logger.log(`Error sending email: ${e.toString()}`);
  }
}
