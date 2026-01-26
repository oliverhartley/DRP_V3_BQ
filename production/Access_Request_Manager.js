/**
 * @fileoverview Script to check for and approve pending access requests for ALL files in a specific Google Drive folder.
 * 
 * REQUIRED SCOPES:
 * - https://www.googleapis.com/auth/drive
 * or
 * - https://www.googleapis.com/auth/drive.file (might be insufficient for listing all files if not created by script)
 */

/**
 * Main entry point: specific folder ID to check.
 * Replace 'YOUR_FOLDER_ID' with the actual folder ID if you want to hardcode it,
 * or pass it as an argument if calling from another function.
 */
function runAccessApprovalForFolder() {
  const folderId = "1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC"; // <--- CHANGE THIS or pass as argument
  approveAccessRequestsInFolder(folderId);
}

/**
 * Checks for pending access requests for all files within a specific folder and approves them.
 * @param {string} folderId - The ID of the folder to scan.
 */
function approveAccessRequestsInFolder(folderId) {
  try {
    console.log(`Scanning folder: ${folderId} for files...`);
    const files = listFilesInFolder(folderId);
    console.log(`Found ${files.length} file(s) in folder. Checking for access requests...`);

    files.forEach((file, index) => {
      resolveAccessProposals(file.id, file.name, index + 1, files.length);
    });

    console.log("Folder scan complete.");
  } catch (e) {
    console.error("FATAL ERROR in approveAccessRequestsInFolder: " + e.message);
  }
}

/**
 * Lists all file IDs and names in a folder using the Drive API.
 * Uses pagination to ensure all files are retrieved.
 * @param {string} folderId 
 * @returns {Array<{id: string, name: string}>}
 */
function listFilesInFolder(folderId) {
  const token = ScriptApp.getOAuthToken();
  const foundFiles = [];
  let pageToken = null;

  try {
    do {
      // Query to list files where the parent is folderId and trash is false
      // Fields: nextPageToken, files(id, name)
      const query = `'${folderId}' in parents and trashed = false`;
      const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=nextPageToken,files(id,name)&pageSize=1000`;

      const params = {
        method: 'get',
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true
      };

      const finalUrl = pageToken ? `${url}&pageToken=${pageToken}` : url;
      const response = UrlFetchApp.fetch(finalUrl, params);

      if (response.getResponseCode() !== 200) {
        throw new Error(`Failed to list files. Code: ${response.getResponseCode()}, Body: ${response.getContentText()}`);
      }

      const json = JSON.parse(response.getContentText());
      if (json.files && json.files.length > 0) {
        foundFiles.push(...json.files);
      }

      pageToken = json.nextPageToken;
    } while (pageToken);
  } catch (e) {
    console.error("Error listing files in folder: " + e.message);
    throw e;
  }

  return foundFiles;
}

/**
 * Fetches pending proposals and iterates through them for a single file.
 * @param {string} fileId 
 * @param {string} fileName - For logging purposes
 * @param {number} current - Current file number
 * @param {number} total - Total files
 */
function resolveAccessProposals(fileId, fileName, current, total) {
  const token = ScriptApp.getOAuthToken();
  const prefix = `[${current}/${total}]`;
  // Using the original logic but added error handling for 403/404 if needed
  const listUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/accessproposals`;

  const params = {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(listUrl, params);

  if (response.getResponseCode() !== 200) {
    // Sometimes accessing proposals might fail if we don't have permission to manage that specific file
    console.warn(`${prefix} Could not fetch proposals for file '${fileName}' (${fileId}). Status: ${response.getResponseCode()}`);
    return;
  }

  const json = JSON.parse(response.getContentText());

  if (json.error) {
    console.error(`${prefix} API Error listing proposals for file '${fileName}': ` + json.error.message);
    return;
  }

  const proposals = json.accessProposals || [];

  if (proposals.length > 0) {
    console.log(`${prefix} [Pending] File '${fileName}' has ${proposals.length} request(s). Processing...`);
    proposals.forEach(prop => {
      // prop.proposalId is the key as per user's original correction
      approveProposal(fileId, prop.proposalId, prop.recipientEmailAddress, fileName);
    });
  } else {
    console.log(`${prefix} [Checked] File '${fileName}': No pending requests.`);
  }
}

/**
 * Approves a single proposal.
 * @param {string} fileId 
 * @param {string} proposalId 
 * @param {string} email 
 * @param {string} fileName - For logging
 */
function approveProposal(fileId, proposalId, email, fileName) {
  const token = ScriptApp.getOAuthToken();
  const resolveUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/accessproposals/${proposalId}:resolve`;

  // VALID ACTIONS: 'ACCEPT' or 'DENY'
  // VALID ROLES: 'writer', 'commenter', 'reader'
  const payload = {
    action: "ACCEPT",
    role: "writer"
  };

  const params = {
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(resolveUrl, params);

  if (response.getResponseCode() === 200) {
    console.log(`✅ Approved access for: ${email} on file '${fileName}'`);
    sendAccessGrantedEmail(email, fileName, fileId);
  } else {
    console.error(`❌ Failed to approve ${email} on file '${fileName}'. Status: ${response.getResponseCode()}`);
    console.error(`   Response: ${response.getContentText()}`);
  }
}

/**
 * Sends a beautifully formatted HTML email to the user notifying them of access.
 * @param {string} recipientEmail 
 * @param {string} fileName 
 * @param {string} fileId 
 */
function sendAccessGrantedEmail(recipientEmail, fileName, fileId) {
  const fileUrl = `https://docs.google.com/open?id=${fileId}`;
  const subject = `Acceso Concedido: ${fileName}`;

  // HTML Template with inline CSS for better compatibility
  const htmlBody = `
    <div style="font-family: 'Google Sans', Roboto, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff;">
      <div style="text-align: center; margin-bottom: 24px;">
        <img src="https://www.gstatic.com/images/branding/product/2x/drive_2020q4_48dp.png" alt="Google Drive" style="width: 48px; height: 48px;">
      </div>
      <h2 style="color: #202124; text-align: center; margin-bottom: 16px; font-weight: 500;">Acceso Concedido</h2>
      <p style="color: #3c4043; font-size: 16px; line-height: 24px; text-align: center; margin-bottom: 24px;">
        Se te ha concedido acceso al archivo: <br>
        <strong>${fileName}</strong>
      </p>
      <div style="text-align: center; margin-bottom: 32px;">
        <a href="${fileUrl}" style="background-color: #1a73e8; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: 500; font-size: 14px; display: inline-block;">Abrir Documento</a>
      </div>
      <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 24px 0;">
      <p style="color: #5f6368; font-size: 12px; text-align: center;">
        Este es un mensaje automático. Por favor, no respondas.
      </p>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    console.log(`📧 Notification email sent to ${recipientEmail}`);
  } catch (e) {
    console.error(`Failed to send email to ${recipientEmail}: ` + e.message);
  }
}
