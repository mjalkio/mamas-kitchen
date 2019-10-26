function main() {
  // Cleans up a Google Sheet for Mama's Kitchen!
  deleteBlankColumns();
  fillInHeaders();
  deleteBalanceSheetRows();
  fillInAccountValues();
  fillInAccountType1Values();
  fillInAccountType2Values();
  fillInAccountType3Values();
  fillInAccountType4Values();
  deleteRowsWithTotalsOrBlanks();
}

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

function fillInHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("A1").setValue('Account Type');
  sheet.getRange("B1").setValue('Account');
  sheet.getRange("C1").setValue('Account Type 1');
  sheet.getRange("D1").setValue('Account Type 2');
  sheet.getRange("E1").setValue('Account Type 3');
  sheet.getRange("F1").setValue('Account Type 4');
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

function fillInAccountType2Values() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var accountType2ColumnIndex = 3;
  var lastAccountType2ValueIndex = -1;
  var lastAccountType2Value = '';

  for (var i = 1; i < values.length; i++) {
    if (values[i][accountType2ColumnIndex] == 'Total ' + lastAccountType2Value) {
      var emptyAccountType2Range = sheet.getRange(lastAccountType2ValueIndex + 1, accountType2ColumnIndex + 1, i - lastAccountType2ValueIndex, 1);
      emptyAccountType2Range.setValue(lastAccountType2Value);
    } else if (values[i][accountType2ColumnIndex] != '') {
      var lastAccountType2ValueIndex = i;
      var lastAccountType2Value = values[i][accountType2ColumnIndex];
    }
  }
}

function fillInAccountType3Values() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var accountType3ColumnIndex = 4;
  var lastAccountType3ValueIndex = -1;
  var lastAccountType3Value = '';

  for (var i = 1; i < values.length; i++) {
    if (values[i][accountType3ColumnIndex] == 'Total ' + lastAccountType3Value) {
      var emptyAccountType3Range = sheet.getRange(lastAccountType3ValueIndex + 1, accountType3ColumnIndex + 1, i - lastAccountType3ValueIndex, 1);
      emptyAccountType3Range.setValue(lastAccountType3Value);
    } else if (values[i][accountType3ColumnIndex] != '') {
      var lastAccountType3ValueIndex = i;
      var lastAccountType3Value = values[i][accountType3ColumnIndex];
    }
  }
}

function fillInAccountType4Values() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("F2").setValue('=B2&if(ISBLANK(C2),""," | ")&C2&if(ISBLANK(D2),""," | ")&D2&if(ISBLANK(E2),""," | ")&E2');
  sheet.getRange("F3").setValue('=B3&if(ISBLANK(C3),""," | ")&C3&if(ISBLANK(D3),""," | ")&D3&if(ISBLANK(E3),""," | ")&E3');
  sheet.getRange('F2:F3').copyTo(sheet.getRange('F4:F'));
}

function deleteRowsWithTotalsOrBlanks() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();

  var goodRowValues = [];
  for (var i = 0; i < values.length; i++) {
    var isTotalRow = false;
    // Columns 1 through 5 can signify that a row is a total
    for (var j = 1; j < 6; j++) {
      if (values[i][j].indexOf('Total') == 0) {
        isTotalRow = true;
        break;
      }
    }

    var amountColumnIndex = 12;
    var isBlankRow = values[i][amountColumnIndex] == '';

    if (isTotalRow || isBlankRow) {
      continue;
    } else {
      goodRowValues.push(values[i]);
    }
  }

  range.clearContent();
  var newRange = sheet.getRange(1, 1, goodRowValues.length, goodRowValues[0].length);
  newRange.setValues(goodRowValues);
}
