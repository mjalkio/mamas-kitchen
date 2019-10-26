function columnLetterToIndex(letter) {
  // Columns start from 1 in Google Sheets
  // 'A'.charCodeAt(0) === 65
  return letter.charCodeAt(0) - 64
}

function getCellValue(sheet, rowNum, columnNum) {
  return sheet.getRange(rowNum, columnNum).getValue();
}

function deleteBlankColumns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // These are the blank columns in the sheet
  // We iterate backwards, so that the indices don't change with each deletion
  var columnsToDelete = ['W', 'V', 'T', 'R', 'P', 'N', 'L', 'J', 'H', 'G'];
  for (var i = 0; i < columnsToDelete.length; i++) {
    indexToDelete = columnLetterToIndex(columnsToDelete[i]);
    sheet.deleteColumn(indexToDelete);
  }
}

function deleteBalanceSheetRows() {
  // Need to delete all rows above "Total Unrealized Capital Gain/Loss"
  var sheet = SpreadsheetApp.getActiveSheet();
  while (getCellValue(sheet, 2, 2) != 'Total Unrealized Capital Gain/Loss') {
    sheet.deleteRow(2);
  }
  sheet.deleteRow(2);
}