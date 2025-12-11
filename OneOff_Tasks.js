function processFolderImages() {
  const folderId = "1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC";

  // Array of images to insert: { id, row, col }
  // Col 9 = I
  const images = [
    { id: "1RrY--a7cZ9gYZKFZJa0v4ZIAT75aM0VH", row: 5, col: 9 },
    { id: "1Gf9sghdhjs-tnszdSP00IXlWR52UBaQs", row: 15, col: 9 }
  ];

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const token = ScriptApp.getOAuthToken();

    let count = 0;
    while (files.hasNext()) {
      const file = files.next();
      Logger.log(`[${++count}] Processing file: ${file.getName()} (${file.getId()})`);
      processSingleFile(file.getId(), images, token);
    }

    Logger.log(">>> BATCH COMPLETE <<<");

  } catch (e) {
    Logger.log("Critical Error in Batch: " + e.toString());
  }
}

function processSingleFile(fileId, images, token) {
  const sheetName = "Tier Dashboard";
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`  -> [SKIP] Sheet '${sheetName}' not found in ${ss.getName()}`);
      return;
    }

    // Check existing images
    const existingImages = sheet.getImages();
    const occupiedCells = new Set();

    existingImages.forEach(img => {
      try {
        const anchor = img.getAnchorCell();
        const key = `${anchor.getRow()}_${anchor.getColumn()}`;
        occupiedCells.add(key);
      } catch (e) {
        // Ignore images with issues getting anchor
      }
    });

    images.forEach(img => {
      try {
        const cellKey = `${img.row}_${img.col}`;
        if (occupiedCells.has(cellKey)) {
          Logger.log(`  -> [SKIP] Image already exists at ${sheetName}!R${img.row}C${img.col}`);
          return;
        }

        Logger.log(`  -> Inserting image ${img.id} at R${img.row}C${img.col}...`);

        // Workaround for "Blob too large" or "Pixel limit exceeded"
        const resizeUrl = `https://drive.google.com/thumbnail?id=${img.id}&sz=w1000`;

        const response = UrlFetchApp.fetch(resizeUrl, {
          headers: { 'Authorization': 'Bearer ' + token },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
          Logger.log(`    -> Failed to fetch resized image ${img.id}. Code: ${response.getResponseCode()}. Attempting direct DriveApp fetch...`);
          try {
            const file = DriveApp.getFileById(img.id);
            const blob = file.getBlob();
            sheet.insertImage(blob, img.col, img.row);
            Logger.log(`    -> Success! Inserted (Direct)`);
            return;
          } catch (fallbackEx) {
            Logger.log(`    -> Fallback failed for ${img.id}: ${fallbackEx.toString()}`);
            return;
          }
        }

        const imageBlob = response.getBlob();
        sheet.insertImage(imageBlob, img.col, img.row);
        Logger.log(`    -> Success! Inserted (Resized)`);

      } catch (innerEx) {
        Logger.log(`    -> Error inserting image ${img.id}: ${innerEx.toString()}`);
      }
    });

  } catch (e) {
    Logger.log(`  -> Error processing file ${fileId}: ${e.toString()}`);
  }
}

function shareFolderFiles() {
  const folderId = "1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC";
  const editors = [
    "jcarrique@google.com",
    "thiagodaponte@google.com",
    "ignaciorauda@google.com",
    "erikasavio@google.com"
  ];

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let count = 0;

    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      count++;
      Logger.log(`[${count}] Processing file: ${file.getName()}`);

      editors.forEach(email => {
        try {
          // Use Advanced Drive Service to suppress notifications
          // Requires "Drive" service enabled in appsscript.json
          Drive.Permissions.insert(
            {
              'role': 'writer',
              'type': 'user',
              'value': email
            },
            fileId,
            {
              'sendNotificationEmails': false
            }
          );
          Logger.log(`  -> Shared with ${email} (No Notification)`);
        } catch (e) {
          Logger.log(`  -> Failed to share with ${email}: ${e.toString()}`);
        }
      });
    }
    Logger.log(">>> SHARING COMPLETE <<<");
  } catch (e) {
    Logger.log("Critical Error: " + e.toString());
  }
}

/**
 * One-click setup for the Autonomous Weekly Partner Summary Batch.
 * Creates a Time-Driven trigger to run every 1 hour.
 * Prevents duplicates by checking existing triggers.
 */
function setupBatchEmailTrigger() {
  const functionName = 'runBatchEmailSender';

  // 1. Check for existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      Logger.log(`Trigger for '${functionName}' already exists. skipping.`);
      return;
    }
  }

  // 2. Create new Trigger (Every 1 Hour)
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log(`SUCCESS: Created hourly trigger for '${functionName}'.`);
}
