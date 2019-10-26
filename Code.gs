function columnLetterToIndex(letter) {
  // Columns start from 1 in Google Sheets
  // 'A'.charCodeAt(0) === 65
  return letter.charCodeAt(0) - 64
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
