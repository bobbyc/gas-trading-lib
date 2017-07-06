function YahooQueryStockHistorical(Code, HistoryDays, Interval, MaxRecords) {

  // Initialize
  var Today = new Date();
  var StartDay = new Date();
  StartDay.setDate(Today.getDate() - HistoryDays);
  var interval = (Interval == 7) ? "w" : "d";

  Logger.log(Today);
  Logger.log(StartDay);

  // Sample: https://query1.finance.yahoo.com/v7/finance/download/WOW.AX?period1=1496733149&period2=1499325149&interval=1d&events=history&crumb=MLX/DeGt.EH
  // Query historical data
  var options =
    {
      "period1": StartDay,
      "period2": Today,
      "interval": [Interval,'d'].join(),
      "events": 'history',
    };

  var url = ["https://query1.finance.yahoo.com/v7/finance/download/", Code, '?'].join();
  for (var o in options) {
    url = url + o + "=" + options[o] + "&";
  }
  Logger.log(url);

  var response = UrlFetchApp.fetch(url);
  data = Utilities.parseCsv(response.getContentText());
  //Logger.log(data);

  return data;
}

function GoogleQueryStockHistorical(Symbol, HistoryDays, Interval, MaxRecords) {

  // Initialize
  var Today = new Date();
  var StartDay = new Date();
  StartDay.setDate(Today.getDate() - HistoryDays);

  // fetch URL
  var options =
    {
      "q": Symbol,
      "startdate": Utilities.formatDate(StartDay, "GMT", "MMM d, YYYY"),
      "enddate": Utilities.formatDate(Today, "GMT", "MMM d, YYYY"),
      "output": "csv"
    };

  var url = "https://www.google.com/finance/historical?"
  for (var o in options) {
    url = url + o + "=" + options[o] + "&";
  }
  Logger.log(url);

  var response = UrlFetchApp.fetch(url);
  data = Utilities.parseCsv(response.getContentText());
  //Logger.log(data);

  return data;
}
