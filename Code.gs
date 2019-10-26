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
  var ACCOUNT_TYPE_MAP = {
    'Capital Campaign Income': 'Income',
    'Contract Income': 'Income',
    'Events Income': 'Income',
    'Fee for Service': 'Income',
    'Grants Income': 'Income',
    'Individual Income': 'Income',
    'Current Payables': 'Expenses',
    'Bank/Credit Card fees': 'Expenses',
    'Client Expense': 'Expenses',
    'Consulting & Professional Fees': 'Expenses',
    'Consumables': 'Expenses',
    'Containers/Bags': 'Expenses',
    'Data Costs': 'Expenses',
    'Employee Expenses': 'Expenses',
    'Facilities': 'Expenses',
    'Food Costs': 'Expenses',
    'Food Waste': 'Expenses',
    'Insurance': 'Expenses',
    'Interest Exp - Mortgage & LMA': 'Expenses',
    'Investment Admin Fees': 'Expenses',
    'Kitchen Equip': 'Expenses',
    'Marketing & Public Relations': 'Expenses',
    'Miscellaneous': 'Expenses',
    'Office Supplies & Equip Lease': 'Expenses',
    'Postage': 'Expenses',
    'Staff Development': 'Expenses',
    'Van Expenses': 'Expenses',
    'Volunteer Expenses': 'Expenses',
    'InKind Income': 'Other Income',
    'Investment Income': 'Other Income',
    'Other Income': 'Other Income',
    'InKind Expenses': 'Other Expenses',
    'Other Expense': 'Other Expenses'
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var accountColumnIndex = 1;
  var lastAccountValueIndex = 1;
  for (var i = 2; i < values.length; i++) {
    if (values[i][accountColumnIndex] != '') {
      var lastAccountValue = values[lastAccountValueIndex][accountColumnIndex];
      var emptyAccountRange = sheet.getRange(lastAccountValueIndex + 1, accountColumnIndex + 1, i - lastAccountValueIndex, 1);
      emptyAccountRange.setValue(lastAccountValue);

      var accountTypeRange = sheet.getRange(lastAccountValueIndex + 1, accountColumnIndex, i - lastAccountValueIndex, 1);
      accountTypeRange.setValue(ACCOUNT_TYPE_MAP[lastAccountValue]);
      lastAccountValueIndex = i;
    }
  }
}

function fillInAccountType1Values() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var accountType1ColumnIndex = 2;
  var lastAccountType1ValueIndex = -1;
  var lastAccountType1Value = '';

  for (var i = 1; i < values.length; i++) {
    if (values[i][accountType1ColumnIndex] == 'Total ' + lastAccountType1Value) {
      var emptyAccountType1Range = sheet.getRange(lastAccountType1ValueIndex + 1, accountType1ColumnIndex + 1, i - lastAccountType1ValueIndex, 1);
      emptyAccountType1Range.setValue(lastAccountType1Value);
    } else if (values[i][accountType1ColumnIndex] != '') {
      var lastAccountType1ValueIndex = i;
      var lastAccountType1Value = values[i][accountType1ColumnIndex];
    }
  }
}
