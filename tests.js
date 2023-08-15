function testYahoo(){
  t = new YahooFinance();
  var data = t.QueryHistorical('0056.TW', 100, 1, 1000);
  console.log(data);
}

function testYahooQuote(){
  t = new YahooFinance();
  var data = t.Quote('0050.TW');
  console.log(data);
}

function testProvision() {
  provisionSheetsQuote("0050.TW");
}

function testProvisionSheet() {
  var SheetName = "Rich";
  resetSheet( SheetName );
  provisionSheets( SheetName );
};

function testQuery() {

  var rows = QueryStockValuesEx("Day", "V", 380, 1, 250);
  var rows = QueryStockValuesEx("Week", "V", 3650, 7, 500);

  Logger.log(rows);

}

function testStockPrice()
{
  var price = FinanceApp.getStockInfo("0050")
}

function testLastRow(symbol) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template = spreadsheet.getSheetByName("Day");

  var lastRow = template.getLastRow();

  Logger.log(lastRow);

}

function testRangeHeight()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template = spreadsheet.getSheetByName("Day");
  var range = template.getRange("A:A");

  Logger.log(range.getLastRow());
  var aaa = 100;
  Logger.log("A"+aaa);

}
