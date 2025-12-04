function debugAccentureData() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const dbSheet = ss.getSheetByName("LATAM_Partner_DB");
  const scoreSheet = ss.getSheetByName("LATAM_Partner_Score_DRP");
  
  const dbData = dbSheet.getDataRange().getValues();
  const scoreData = scoreSheet.getDataRange().getValues();
  
  let dbCount = "Not Found";
  for (let i = 1; i < dbData.length; i++) {
    if (String(dbData[i][1]).includes("Accenture")) { // Name is col 1
      // Find Total Profiles column index
      const headers = dbData[0];
      const idx = headers.indexOf("Total_Profiles");
      dbCount = dbData[i][idx];
      break;
    }
  }
  
  let scoreCount = "Not Found";
  for (let i = 3; i < scoreData.length; i++) {
    if (String(scoreData[i][1]).includes("Accenture")) { // Name is col 1
      scoreCount = scoreData[i][2]; // Total Profiles is col 2
      break;
    }
  }
  
  Logger.log("Accenture in DB (Dashboard Source): " + dbCount);
  Logger.log("Accenture in Score Sheet (Delta Source): " + scoreCount);
}
