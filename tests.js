function testProvision() {
  
  provision("0050.TW");
  
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

function testCSV()
{
  //var rows = QueryStockValuesEx("Day", "AAPL", 380, 1, 250);
  //var rows = QueryStockValuesEx("Week", "AAPL", 3650, 5, 500);
  var rows = YahooQueryStockHistorical("YHOO", 3650, 1, 500);
  
  Logger.log(rows);
  
  //"http://www.google.com/finance/historical?cid=660463&startdate=Nov+3%2C+2013&enddate=Nov+2%2C+2014&num=30&ei=uPdVVNC4HMyhkgW82IGgCA&output=csv"
}