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

// ... (getPartnerSheetData remains the same) ...

// ... (callGeminiWithFallback remains the same) ...

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
