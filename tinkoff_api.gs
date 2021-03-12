function makeApiPostCall_(methodUrl, json_payload)
{
  var url = 'https://api-invest.tinkoff.ru/openapi/' + methodUrl;  
  var response = UrlFetchApp.fetch(url, {
    'headers': {'accept': 'application/json', "Authorization": `Bearer ${props.getProperty("token")}`},
    'method' : 'post',
    'escaping': false,
    'muteHttpExceptions':true,
    "payload" : json_payload});
  if (response.getResponseCode() != 200) return null;
  Logger.log(response.getContentText()); 
  return JSON.parse(response.getContentText()).payload;
}

function makeApiGetCall_(methodUrl)
{
  var url = 'https://api-invest.tinkoff.ru/openapi/' + methodUrl;
  var response = UrlFetchApp.fetch(url, {
    'headers': {'accept': 'application/json', "Authorization": `Bearer ${props.getProperty("token")}`},
    'escaping': false,
    'muteHttpExceptions':true});
  if (response.getResponseCode() != 200) return null;
  //Logger.log(response.getContentText());
  return JSON.parse(response.getContentText()).payload;
}

/**
 * @customfunction
 */
function getFigi(tiker, forceUpdate = false)
{
  if (tiker == null) return null;
  tiker = tiker.toLowerCase();
  var cached = cache.get(tiker);
  if (!forceUpdate && cached != null)
    return cached;
  var data = makeApiGetCall_(`market/search/by-ticker?ticker=${tiker}`);
  if (data == null) return null;
  var figi = data.instruments[0].figi;
  Logger.log('Put to cache: ' + tiker + ' = ' + figi);
  cache.put(tiker, figi, 86400);
  return figi;
}

/**
 * @customfunction
 */
function getTiker(figi, forceUpdate = false)
{
  if (figi == null) return null;
  figi = figi.toUpperCase();
  var cached = cache.get(figi);
  if (!forceUpdate && cached != null)
    return cached;
  var data = makeApiGetCall_(`market/search/by-figi?figi=${figi}`);
  if (data == null) return null;
  var tiker = data.ticker;
  if (tiker == 'USD000UTSTOM') tiker = '_USD';
  if (tiker == 'EUR_RUB__TOM') tiker = '_EUR';
  Logger.log('Put to cache: ' + figi + ' = ' + tiker); 
  cache.put(figi, tiker, 86400);
  return tiker;
}

/**
 * @customfunction
 */
function getPortfolio()
{
  var portfolio = makeApiGetCall_(`portfolio`).positions;
  if (portfolio.length == 0) 
    return 'PORTFOLIO IS EMPTY';
  var values = [];
  for (var i = 0; i < portfolio.length; i++)
  {
    var n = portfolio[i];
    values.push([n.figi, n.ticker, n.balance, n.averagePositionPrice.value, n.name]);
  }
  return values;
}

/**
 * @customfunction
 */
function getOrderlist()
{
  var orderlist = makeApiGetCall_(`orders`);
  if (orderlist == null || orderlist.length == 0) 
    return 'NO ORDERS';
  var values = [];
  for (var i = 0; i < orderlist.length; i++)
  {
    var n = orderlist[i];
    values.push([n.orderId, n.figi, getTiker(n.figi), n.operation, n.status, n.requestedLots, n.executedLots, n.type, n.price]);
  }
  return values;
}

/**
 * @customfunction
 */
function pushOrder(tiker, lots, price)
{
  if (tiker == null || lots == null || price == null) return;
  var figi = getFigi(tiker);
  var apicall = `orders/limit-order?figi=${figi}`;
  var data = {
    "lots" : lots,
    "operation" : "Buy",
    "price" : price};
  return makeApiPostCall_(apicall, Utilities.jsonStringify(data));
}

/**
 * @customfunction
 */
function getOperations(from_date = null, to_date = null, figiOrTiker = null)
{
  if (from_date == null)
    from_date = 'Feb 07, 2019 03:00:00';
  from_date = new Date(from_date).toISOString();
  
  if (to_date == null)
    to_date = new Date();
  else
    to_date = new Date(to_date);
  to_date = new Date(to_date.getTime() + 10800000).toISOString();
  
  var apicall = `operations?from=${from_date}&to=${to_date}`
  if (figiOrTiker != null)
  {
    if (figiOrTiker.length < 6)
      figiOrTiker = getFigi(figiOrTiker);
    apicall += `&figi=${figiOrTiker}`;
  }
  Logger.log(apicall);
  var operations = makeApiGetCall_(apicall);
  if (operations == null) return null;
  var operations = operations.operations;
  
  var values = [];
  for (var i = 0; i < operations.length; i++)
  {
    var n = operations[i];
    var tik = getTiker(n.figi);
    Logger.log(n.date + " " + tik + " " + n.price + " " + n.quantityExecuted);
    var commission = "";
    if (n.quantityExecuted > 0) commission = -n.commission.value;
    //values.push([n.id, n.date, n.status, n.figi, tik, n.currency, n.payment, n.price, n.quantity, n.quantityExecuted, n.operationType]); //old
    values.push([n.date, n.figi, tik, n.operationType, n.price, n.currency, n.quantityExecuted, n.payment, commission, n.status, n.id]);
  }
  return values;
}
 
