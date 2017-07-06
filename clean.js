
function OnCleanSheets() {

  resetSheet("Dashboard");
  resetSheet("0050");
}

function resetSheet(SheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SheetName);
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();

  var rangeValues = sheet.getRange(2, 7, numRows, 13);
  rangeValues.clearContent();

}

function deleteSheet(symbol) {

  if (symbol != "") {
    // create new sheet from template
    var symbolsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(symbol);
    if (symbolsheet != null) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(symbolsheet);
    }
  }
}
