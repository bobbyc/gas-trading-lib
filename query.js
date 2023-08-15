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
  this.header = {
    'Connection': 'close',
    'Cookie': 'A3=d=AQABBPB122QCEEPvwXAu-7ZbmGHGpFCYBZsFEgEBAQHH3GTlZL2oQDIB_eMAAA&S=AQAAAi5o0tA89LzdTIMt633Zz5s; Expires=Wed, 14 Aug 2024 18:56:16 GMT; Max-Age=31557600; Domain=.yahoo.com',
    'User-Agent': "Mozilla/5.0 (Macintosh; U; Intel Mac OS X; de-de) AppleWebKit/523.10.3 (KHTML, like Gecko) Version/3.0.4 Safari/523.20",
    'Cache-Control': 'no-cache',
    'Accept-Encoding': 'gzip',
    'Accept-Charset': 'ISO-8859-1,UTF-8;q=0.7,*;q=0.7',
    'Accept-Language': 'de,en;q=0.7,en-us;q=0.3'
  };
}

/**
 * Sheet Generate function
 * @param err
 * @param result
 */
YahooFinance.prototype.Generate = function () {
}

function fabonacci(n) {
  if (n <= 1) {
    return 1;
  } else {
    return fabonacci(n - 1) + fabonacci(n - 2);
  }
}

function RetryUrlFetch(url, opt) {
  // max loop 10 times, 1 second interval per loop
  // Fetch url if content of response is "Too many request"
  var i = 0;
  var max_loop = 10;
  while ( i < max_loop) {
    response = UrlFetchApp.fetch(url, opt);

    // check content of response is "Too many request"
    if (response.getContentText().indexOf("Too Many Requests") < 0) {
      // Logger.log(response.getHeaders());
      // Logger.log(response.getContentText());

      // var regexp = new RegExp(/"CrumbStore":\{"crumb":"(.+?)"\}/);
      // Get crumb from html content
      // this.crumb = regexp.exec(response.getContentText())[1];
      break;
    }

    Logger.log("Too many request, wait fabonacci second and try again");
    Utilities.sleep(fabonacci(i) * 1000);
    i++;
  }

  return response
}

/**
 * Sheet Generate function
 * @param err
 * @param result
 */
YahooFinance.prototype.GetCrumb = function () {

  // Query yahoo page to get site cookie
  var opt = {
    'headers': this.header,
    'muteHttpExceptions': true
  };

  var url = encodeURI("https://fc.yahoo.com");
  var response = UrlFetchApp.fetch(url, opt);
  if (response != undefined) {
    this.header['Cookie'] = response.getHeaders()["Set-Cookie"];
    Logger.log("cookie: " + this.header['Cookie']);

    url = encodeURI("https://query2.finance.yahoo.com/v1/test/getcrumb");
    response = RetryUrlFetch(url, opt)
    if (response != undefined) {
        // Logger.log(response.getHeaders());
        // Logger.log(response.getContentText());
        this.crumb = response.getContentText();

        // var regexp = new RegExp(/"CrumbStore":\{"crumb":"(.+?)"\}/);
        // Get crumb from html content
        // this.crumb = regexp.exec(response.getContentText())[1];
    }

    // Long story short
    // this.cookie = this.header['Cookie'];
    // this.crumb = 'C6tnSPxu1Ei';
  }
}

/**
 * Sheet Generate function
 * @param err
 * @param result
 */
YahooFinance.prototype.QueryHistorical = function (Code, HistoryDays, Interval) {

  // Get query cookie and crumb
  if (this.crumb == undefined)
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

  var url = "https://query2.finance.yahoo.com/v7/finance/download/$symbol?period1=$period1&period2=$period2&interval=$interval&events=history&crumb=$crumb";
  url = url.replace('$symbol', Code);
  url = url.replace('$period1', options['period1']);
  url = url.replace('$period2', options['period2']);
  url = url.replace('$interval', options['interval']);
  url = url.replace('$events', options['events']);
  url = url.replace('$crumb', options['crumb']);
  Logger.log(url);

  var opt = {
    'headers': this.header,
    'muteHttpExceptions': true
  };

  // var response = UrlFetchApp.fetch(url, opt);
  var response = RetryUrlFetch(url, opt)
  if (response != undefined) {
    // parse CSV
    data = Utilities.parseCsv(response.getContentText());

    // Parse Json
    // content = JSON.parse(response.getContentText());
    // data = content['chart']['result'][0]['indicators']['quote'][0]['close'];

    // Logger.log(data);
  }

  return data;
}

YahooFinance.prototype.Quote = function (Code) {

  // Get query cookie and crumb
  if (this.crumb == undefined)
    this.GetCrumb();

  // var url = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/$symbol?formatted=true&modules=$modules&corsDomain=$corsDomain&crumb=$crumb';
  var url = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/$symbol?formatted=true&modules=$modules&crumb=$crumb';
  url = url.replace('$symbol', encodeURI(Code));
  url = url.replace('$crumb', this.crumb);
  url = url.replace('$modules', 'price,summaryDetail');
  url = url.replace('$corsDomain', 'finance.yahoo.com');
  Logger.log(url);

  var opt = {
    'headers': this.header,
    'muteHttpExceptions': true
  };

  // var response = UrlFetchApp.fetch(url, opt);
  var response = RetryUrlFetch(url, opt)
  if (response != undefined) {
    data = JSON.parse(response.getContentText());
    // Logger.log(data); 
  }

  return data;
}

// function GoogleQueryStockHistorical(Symbol, HistoryDays, Interval, MaxRecords) {

//   // Initialize
//   var Today = new Date();
//   var StartDay = new Date();
//   StartDay.setDate(Today.getDate() - HistoryDays);

//   // fetch URL
//   var options =
//     {
//       "q": Symbol,
//       "startdate": Utilities.formatDate(StartDay, "GMT", "MMM d, YYYY"),
//       "enddate": Utilities.formatDate(Today, "GMT", "MMM d, YYYY"),
//       "output": "csv"
//     };

//   var url = "https://www.google.com/finance/historical?"
//   for (var o in options) {
//     url = url + o + "=" + options[o] + "&";
//   }
//   Logger.log(url);

//   var response = UrlFetchApp.fetch(url);
//   if (response) {
//     data = Utilities.parseCsv(response.getContentText());
//     //Logger.log(data);
//   } else {
//     data = null;
//   }

//   return data;
// }
