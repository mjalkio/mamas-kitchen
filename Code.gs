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
  var range = sheet.getDataRange();
  var values = range.getValues();
  var totalUnrealizedCapitalGainLossColumnIndex = 1;
  for (var i = 0; i < values.length; i++) {
    if (values[i][totalUnrealizedCapitalGainLossColumnIndex] == 'Total Unrealized Capital Gain/Loss') {
      break;
    }
  }

  var firstNonHeaderRowNum = 2;
  sheet.deleteRows(firstNonHeaderRowNum, i);
}

function fillInAccountValues() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var accountColumnIndex = 1;
  var lastAccountValueIndex = 1;
  for (var i = 2; i < values.length; i++) {
    if (values[i][accountColumnIndex] != '') {
      var emptyAccountRange = sheet.getRange(lastAccountValueIndex + 1, accountColumnIndex + 1, i - lastAccountValueIndex, 1);
      Logger.log(values[lastAccountValueIndex][accountColumnIndex]);
      emptyAccountRange.setValue(values[lastAccountValueIndex][accountColumnIndex]);
      lastAccountValueIndex = i;
    }
  }
}
