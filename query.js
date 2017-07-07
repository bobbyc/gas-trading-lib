
/**
 * Constructor - YahooFinance
 * @param err
 * @param result
 */
var YahooFinance = function () {
  this.name = 'Yahoo Finance';

  // Sheet render offsets
  this.cookie = undefined;
  this.crumb = undefined;
}

/**
 * Sheet Generate function
 * @param err
 * @param result
 */
YahooFinance.prototype.Generate = function () {

  // Report generation function
  for (var ver = 0; ver < FFversion.length; ver++) {

    // Initial val
    this.CountFixedBugsByTeam(startRowRuntime, startColumn + ver, contributorRowRuntime, FFversion[ver], [TaipeiDOM], undefined);
  }
}

/**
 * Sheet Generate function
 * @param err
 * @param result
 */
YahooFinance.prototype.GetCrumb = function () {

  // Query yahoo page to get site cookie
  var header = {
    'Connection': 'close',
    'Cookie': 'B=dm5t1ehclu8uh&b=3&s=o6; expires=Fri, 07-Jul-2018 06:01:53 GMT; path=/; domain=.yahoo.com',
    'User-Agent': "Mozilla/5.0 (Macintosh; U; Intel Mac OS X; de-de) AppleWebKit/523.10.3 (KHTML, like Gecko) Version/3.0.4 Safari/523.10",
    'Cache-Control': 'no-cache',
    'Accept-Encoding': 'gzip',
    'Accept-Charset': 'ISO-8859-1,UTF-8;q=0.7,*;q=0.7',
    'Accept-Language': 'de,en;q=0.7,en-us;q=0.3'
  };

  var opt = {
    'headers': header,
    'muteHttpExceptions': true
  };
  var regexp = new RegExp(/"CrumbStore":\{"crumb":"(.+?)"\}/);

  // var url = encodeURI("https://query1.finance.yahoo.com/v7/finance/download/GOOG?period1=1496814709&period2=1499406709&interval=1d&events=history&crumb=C6tnSPxu1Ei");
  // var response = UrlFetchApp.fetch(url, opt);
  // var data = response.getContentText();
  // var crumb = regexp.exec(data);
  // Logger.log("crumb: " + crumb[0]);
  // var url = "https://finance.yahoo.com/";
  // var response = UrlFetchApp.fetch(url);
  // if (response != undefined) {

  //     // Get cookie from response
  //     var cookie = response.getHeaders()["Set-Cookie"];
  //     this.cookie = cookie;

  //     // Get crumb from html content
  //     var regexp = new RegExp(/"CrumbStore":\{"crumb":"(.+?)"\}/);
  //     this.crumb = regexp.exec(response.getContentText());
  //     Logger.log("cookie: " + this.cookie);
  //     Logger.log("crumb: " + this.crumb);
  // }


  // Long story short
  this.cookie = header['Cookie'];
  this.crumb = 'C6tnSPxu1Ei';

}


/**
 * Sheet Generate function
 * @param err
 * @param result
 */
YahooFinance.prototype.QueryHistorical = function (Code, HistoryDays, Interval) {

  // Get query cookie and crumb
  if (this.cookie == undefined)
    this.GetCrumb();

  // Initialize
  var Today = new Date();
  var StartDay = new Date();
  StartDay.setDate(Today.getDate() - HistoryDays);
  var interval = (Interval == 7) ? "1wk" : [Interval,"d"].join("");
  // Logger.log(Today);
  // Logger.log(StartDay);

  // Sample: https://query1.finance.yahoo.com/v7/finance/download/WOW.AX?period1=1496733149&period2=1499325149&interval=1d&events=history&crumb=MLX/DeGt.EH
  // Query historical data
  var options = {
    "period1": (StartDay.getTime() / 1000).toFixed(0),
    "period2": (Today.getTime() / 1000).toFixed(0),
    "interval": interval,
    "events": 'history',
    "crumb": this.crumb
  };

  var url = "https://query1.finance.yahoo.com/v7/finance/download/$symbol?period1=$period1&period2=$period2&interval=$interval&events=history&crumb=$crumb";
  url = url.replace('$symbol', Code);
  url = url.replace('$period1', options['period1']);
  url = url.replace('$period2', options['period2']);
  url = url.replace('$interval', options['interval']);
  url = url.replace('$events', options['events']);
  url = url.replace('$crumb', options['crumb']);
  Logger.log(url);

  var header = {
    'Connection': 'close',
    'Cookie': 'B=dm5t1ehclu8uh&b=3&s=o6; expires=Fri, 07-Jul-2018 06:01:53 GMT; path=/; domain=.yahoo.com',
    'User-Agent': "Mozilla/5.0 (Macintosh; U; Intel Mac OS X; de-de) AppleWebKit/523.10.3 (KHTML, like Gecko) Version/3.0.4 Safari/523.10",
    'Cache-Control': 'no-cache',
    'Accept-Encoding': 'gzip',
    'Accept-Charset': 'ISO-8859-1,UTF-8;q=0.7,*;q=0.7',
    'Accept-Language': 'de,en;q=0.7,en-us;q=0.3'
  };

  var opt = {
    'headers': header,
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, opt);
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
  if (response) {
    data = Utilities.parseCsv(response.getContentText());
    //Logger.log(data);
  } else {
    data = null;
  }

  return data;
}
