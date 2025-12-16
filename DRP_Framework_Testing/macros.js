function helppainting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E1:H1').activate();
  spreadsheet.getActiveRangeList().setBackground('#ffff00')
  .setFontColor('#ff0000')
  .setHorizontalAlignment('left');
};