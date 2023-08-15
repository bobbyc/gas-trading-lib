var info_range = "A3:D250";
var copy_range = "A3:AQ250";

// description: duplicate symbol sheet from template
//
// symbol: stock symbol
//
function duplicateSheet(fromSheet, toSheet) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template = spreadsheet.getSheetByName(fromSheet);

  var symbolsheet = null;
  if (toSheet != "") {

    // create new sheet from template
    symbolsheet = spreadsheet.getSheetByName(toSheet);
    if (symbolsheet == null) {
      symbolsheet = template.copyTo(spreadsheet);
      symbolsheet.setName(toSheet);
    }

    // update template to symbolsheet
    var templateData = template.getRange(copy_range);
    var dataRange = symbolsheet.getRange(copy_range);
    dataRange.setValues(templateData.getValues());
  }

  return symbolsheet;
}

//
// example: valuerange = [[ 1, 2, 3, 4, 5 ]];
//
function checkBuyingState(price, valuerange) {

  var RATING = ["OS", "SS", "NS", "NB", "SB", "OB"];

  var vrange = valuerange[0];
  if (!vrange[0] || !vrange[4])
    return "";

  if (price <= vrange[0])
    return RATING[0];
  else if (price > vrange[0] && price <= vrange[1])
    return RATING[1];
  else if (price > vrange[1] && price <= vrange[2])
    return RATING[2];
  else if (price > vrange[2] && price <= vrange[3])
    return RATING[3];
  else if (price > vrange[3] && price <= vrange[4])
    return RATING[4];
  else if (price > vrange[4])
    return RATING[5];

  return "";

}

//function QueryStockValues(Sheet, Symbol, HistoryDays, Interval, MaxRecords) {
//  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//  var Sheet = spreadsheet.getSheetByName(Sheet);
//
//  // Range Variables
//  var SymbolCell = "A1";
//  var HistoricalCell = "A2";
//  var RngQuery = "B1:B";
//  var TitleStart = 3;
//  var TitleEnd = Sheet.getLastColumn()-TitleStart+1;
//
//  //Reset symbol to query values
//  Sheet.getRange(1, TitleStart, 1, TitleEnd).clearContent();
//  Sheet.getRange(1, 1, MaxRecords, 2).clearContent();
//  Sheet.getRange(SymbolCell).setValue(Symbol);
//  Sheet.getRange(HistoricalCell).setFormula("googlefinance(\""+ Symbol +"\", \"close\", TODAY()-"+ HistoryDays +", TODAY(), "+ Interval +")");
//
//  // workaround to reflect data
//  var count=100;
//  var bDateFound = -1
//  while ( bDateFound == -1 && count >= 0) {
//
//    //
//    --count;
//
//    // Find stock value ready
//    var randomWait = Math.floor(Math.random()*1+0); randomWait;
//    var values = Sheet.getRange("A1:B2").getValues();
//    bDateFound = values.toString().search("Date");
//    console.log(values);
//
//  }
//
//  // Calculate Max rows
//  var rows = 1;
//  if (bDateFound != -1) {
//    var QueryValues = Sheet.getRange(RngQuery).getValues();
//    var RangeRows = QueryValues.length;
//    for (; rows < RangeRows; ++rows) {
//      if (QueryValues[rows][0] == "")
//        break;
//    }
//    console.log(QueryValues);
//    console.log(rows);
//
//    //Update title
//    Sheet.getRange(1, TitleStart, 1, TitleEnd).setValues( Sheet.getRange(rows, TitleStart, 1, TitleEnd).getValues() );
//  }
//
//  return rows;
//}

function QueryStockValuesEx(SheetName, symbol, HistoryDays, Interval, MaxRecords) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = spreadsheet.getSheetByName(SheetName);

  var yahoo = new YahooFinance();

  // Range Variables
  var SymbolCell = "A1";
  var HistoricalCell = "A2";
  var RngQuery = "B1:B";
  var TitleStart = 3;
  var TitleEnd = Sheet.getLastColumn() - TitleStart + 1;

  //Reset symbol to query values
  //Sheet.getRange(1, TitleStart, 1, TitleEnd).clearContent();
  Sheet.getRange(1, 1, MaxRecords, 2).clearContent();

  // Fetch Data
  var data = yahoo.QueryHistorical(symbol, HistoryDays, Interval);

  // Calculate history data
  var dataRow = data.length;
  var offsetDate = 0;
  var offsetClose = 4;
  var endRow = dataRow <= MaxRecords ? dataRow : MaxRecords;
  for (var i = 0; i < endRow - 1; ++i) {
    // Copy history data
    var index = i + 1;
    Sheet.getRange(endRow - i, 1, 1, 1).setValue(data[index][offsetDate]);
    Sheet.getRange(endRow - i, 2, 1, 1).setValue(data[index][offsetClose]);
  }

  //Update title
  Sheet.getRange(SymbolCell).setValue(symbol);
  Sheet.getRange(HistoricalCell).setValue(data[0][0]);
  Sheet.getRange("B2").setValue(data[0][4]);

  return dataRow;
}

function provisionEx(symbol) {

  // Query value to Day and Week Sheet
  QueryStockValuesEx("Day", symbol, 380, 1, 250);
  QueryStockValuesEx("Week", symbol, 3650, 7, 500);

}

function provisionSheets(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(sheetName);
  var DaySheet = spreadsheet.getSheetByName("Day");
  var WeekSheet = spreadsheet.getSheetByName("Week");

  var range = dashboard.getDataRange();
  var values = range.getValues();
  var numRows = range.getNumRows();

  for (var i = 1; i < numRows; i++) {

    var row = values[i];
    var symbol = row[0];
    var rsi = row[18];

    if ((symbol == "") || (rsi != ""))
      continue;

    // create new sheet from template
    provisionEx(symbol);

    // Update Dashboard table
    if (symbol != "") {

      // update sold/bought rating
      var rating = []
      var day20 = DaySheet.getRange("D1:H1").getValues();
      var day50 = DaySheet.getRange("K1:O1").getValues();
      var day120 = DaySheet.getRange("R1:V1").getValues();

      var y1 = WeekSheet.getRange("D1:H1").getValues();
      var y5 = WeekSheet.getRange("K1:O1").getValues();
      var y10 = WeekSheet.getRange("R1:V1").getValues();

      rating.push([ checkBuyingState(row[2], day20),
                      checkBuyingState(row[2], day50),
                      checkBuyingState(row[2], day120),
                      checkBuyingState(row[2], y1),
                      checkBuyingState(row[2], y5),
                      checkBuyingState(row[2], y10)
      ]);

      //
      // modify here if dashboard change
      //
      var col_index = 6;
      var range_day50 = DaySheet.getRange("K1:P1");
      var range_RSI = DaySheet.getRange("AA1:AA1");

      // update average
      range.offset(i, col_index, 1, range_day50.getNumColumns()).setValues(range_day50.getValues());
      col_index += range_day50.getNumColumns();

      // update rating
      range.offset(i, col_index, 1, rating[0].length).setValues(rating);
      col_index += rating[0].length;

      // update RSI
      range.offset(i, col_index, 1, range_RSI.getNumColumns()).setValues(range_RSI.getValues());
      col_index += range_RSI.getNumColumns();

      // update table
      //console.log(rating);
    }
  }
}

function provisionSheetsQuote(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(sheetName);

  var range = dashboard.getDataRange();
  var values = range.getValues();
  var numRows = range.getNumRows();

  var yahoo = new YahooFinance();
  for (var i = 1; i < numRows; i++) {
    var row = values[i];
    var symbol = row[0];
    if ((symbol != ""))
    {
      // create new sheet from template
      var quote = yahoo.Quote(symbol);
      // console.log(quote);

      // Update Dashboard table
      if (quote != undefined) {
        // update quote
        let summary = quote['quoteSummary']['result'][0]
        let price = summary['price'];
        let price_market = price['regularMarketPrice']['fmt'];

        let detail_52w_low = 0;
        let detail_52w_high = 0;
        if ("summaryDetail" in summary) {
          let detail = quote['quoteSummary']['result'][0]['summaryDetail'];
          detail_52w_low = detail['fiftyTwoWeekLow']['fmt']
          detail_52w_high = detail['fiftyTwoWeekHigh']['fmt']
        }

        range.offset(i, 1, 1, 3).setValues([[
          detail_52w_low, // quote['quoteSummary']['result'][0]['summaryDetail']['fiftyTwoWeekLow']['fmt'],
          price_market,
          detail_52w_high // quote['quoteSummary']['result'][0]['summaryDetail']['fiftyTwoWeekHigh']['fmt'],
        ]]);
      }
    }
  }
}

