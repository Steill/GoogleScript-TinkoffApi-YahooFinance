const options = {
  'header':'User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 10_3_1 like Mac OS X) AppleWebKit/603.1.30 (KHTML, like Gecko) Version/10.0 Mobile/14E304 Safari/602.1',
  'muteHttpExceptions':true,
  'escaping': false };

/**
 * @customfunction
 */
function getDividends(tikersString, forceUpdate = false)
{
  //tikersString = 'msft ge intc amd atvi gm'; //for testing purposes only
  if (tikersString.length == 0) return "";
  if (tikersString == "#N/A") return "";
  var tikers = tikersString.toUpperCase().split(" ");

  var values = [];
  for (var i = 0; i < tikers.length; i++)
  {
    var tiker = tikers[i];
    if (tiker == "") continue;
    var cached = cache.get(tiker);
    if (!forceUpdate && cached != null) 
    {
      var rows = cached.split("#");
      var url = rows[0];
      Logger.log(url);
      for (var j = 1; j < rows.length; j++)
      {
        var divs = rows[j].split(" ");
        values.push([url, tiker, divs[0], divs[1]]);
        Logger.log(Utilities.formatString('Got from cache: %s %s %s', tiker, divs[0], divs[1]));
      }
      continue;
    }
    
    var url = `https://finance.yahoo.com/quote/` + tiker + `/history?period1=1549486800&period2=1849486800&filter=div&frequency=1mo`;
    Logger.log(url);
     
    var xml = UrlFetchApp.fetch(url, options).getContentText();
    var subtext = xml.substring(xml.indexOf('<tbody'), xml.indexOf('/tbody') + 7);
    var doc = XmlService.parse(subtext).getRootElement();
    var rows = doc.getChildren();
    var strcache = url;
    
    for (var j = 0; j < rows.length; j++)
    {
      var currentRow = rows[j].getChildren();
      var tDate = currentRow[0].getChildText('span');
      if (tDate == null) continue;      
      var day = tDate.substring(4,6);
      var month = tDate.substring(0,3);
      var date = tDate.substring(8,12);
      switch (month)
      {
        case "Jan":
          date += ".01.";
          break;
        case "Feb":
          date += ".02.";
          break;
        case "Mar":
          date += ".03.";
          break;
        case "Apr":
          date += ".04.";
          break;
        case "May":
          date += ".05.";
          break;
        case "Jun":
          date += ".06.";
          break;
        case "Jul":
          date += ".07.";
          break;
        case "Aug":
          date += ".08.";
          break;
        case "Sep":
          date += ".09.";
          break;
        case "Oct":
          date += ".10.";
          break;
        case "Nov":
          date += ".11.";
          break;
        case "Dec":
          date += ".12.";
          break;
      }
      date += day;
      var amount = currentRow[1].getChildText('strong').replace('.',',');
      strcache += '#' + date + ' ' + amount;
      Logger.log(Utilities.formatString('Put to cache: %s %s %s', tiker, date, amount));
      values.push([url, tiker, date, amount]);
    }
    cache.put(tiker, strcache, 88000)
  }
  return values;
}

/**
 * @customfunction
 */
function getCurrentPrice(tiker)
{
  const base_url = 'query1.finance.yahoo.com/v8/finance/chart/';
   
  var xml = UrlFetchApp.fetch(base_url + encodeURIComponent(tiker), options).getContentText();
  if (xml.substring(19, 23) == 'null') return 0;
  var price = xml.substring(xml.indexOf('regularMarketPrice') + 20, xml.indexOf('chartPreviousClose') - 2 );
  Logger.log(tiker + ', ' + +price);
  return +price;
}
