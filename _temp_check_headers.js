function checkDBHeaders() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_DB);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  console.log("Headers:", headers);
  console.log("Last Column Index:", sheet.getLastColumn());
}
