/** @OnlyCurrentDoc */
const mills_per_day = 1000 * 86400;
const cache = CacheService.getScriptCache();
const props = PropertiesService.getScriptProperties();

/**
 * @customfunction
 */
function tiker_sumifs(tikers, tiker_range, sum_range, criteria_range, is_greater_than_zero = true) 
{
  if (sum_range.length != tiker_range.length || sum_range.length != criteria_range.length) return null;
  
  var values = [];
  for (var t in tikers)
  {
    var tiker = tikers[t].toString();
    if (tiker == "") break;
    var sum = 0;
    for (var i = 0; i < sum_range.length; i++)
    {
      if (tiker_range[i] == "") break;
      if (tiker_range[i] == tiker)
      {
        if (is_greater_than_zero && +criteria_range[i] > 0) sum+= +sum_range[i];
        else if (!is_greater_than_zero && +criteria_range[i] < 0) sum+= +sum_range[i];
      }
    }
    values.push(Math.abs(sum)); 
  }
  return values;
}

/**
 * @customfunction
 */
function getCurrencyAndCompanyName(figisOrTikers)
{
  if (figisOrTikers == null) return;
  var values = [];
  for (var i = 0; i < figisOrTikers.length; i++)
  {
    var tiker = figisOrTikers[i];
    if (tiker.length > 6)
      tiker = getTiker(tiker);
    var data = makeApiGetCall_(`market/search/by-ticker?ticker=${tiker}`);
    if (data == null) continue;
    data = data.instruments[0];
    if (data.currency == 'RUB') data.currency = '₽';
    if (data.currency == 'USD') data.currency = '$';
    Logger.log(tiker + ' = ' + data.name);
    values.push([data.currency, data.name]);
  }
  return values;
}

function getDivs()
{
  var summaries = SpreadsheetApp.getActive().getRangeByName('Summaries');
  var trades = SpreadsheetApp.getActive().getRangeByName('Trades');
  if (summaries == null || trades == null) return null;
  
  summaries = summaries.getValues();
  var tikersString = "";
  var values = [];
  for (var i = 0; i < summaries.length; i++)
  {
    var tiker = summaries[i][0];
    if (tiker == "") break;
    if (summaries[i][5] == "₽") continue;
    tikersString += tiker + " ";
  }
  var divs = getDividends(tikersString).sort(function(a,b){
    var A = new Date(a[2]).getTime();
    var B = new Date(b[2]).getTime();
    if (B < A) return 1;
    if (A < B) return -1;
    return 0;
  });
  
  trades = trades.getValues();
  var tradesCount = 0;
  for (var i = 1; i < trades.length; i++)
  {
    if (trades[i][0] == "")
    {
      tradesCount = i - 1;
      break;
    }
  }
  if (tradesCount == 0) return;

  for (var i = 0; i < divs.length; i++)
  {
    var divTime = new Date(divs[i][2]);
    var lastN = 0;
    var count = "";
    var avgPrice = ""; 
    var currency = ""; 
    for (var k = 1; k <= tradesCount; k++)
    {
      if (new Date(trades[k][0]).getTime() <  divTime.getTime()) continue;
      lastN = k - 1;
      break;
    }
    for (; lastN > 0; lastN--)
    {
      if (divs[i][1] == trades[lastN][1])
      {
        count = trades[lastN][17];
        avgPrice = trades[lastN][18];
        currency = trades[lastN][4];
        break;
      }
    }
    values.push([divs[i][0], divs[i][1], divTime, divs[i][3], count, avgPrice, currency]);
  }
  return values;
}

/*
      new Date(lastOp[0]),  //0 date
      lastOp[2],            //1 tiker
      ++tradesCount,        //2
      lastOp[4],            //3 price
      lastOp[5],            //4 curr
      lastOp[6],            //5 qty
      volume,               //6 
      lastOp[5],            //7 
      lastOp[8],            //8 comis
      lastOp[5],            //9 
      lastOp[summary],      //10
      lastOp[5],            //11
      delta,                //12
      lastOp[5],            //13
      deltaPercent,         //14
      lastPos,              //15
      qtyAtStart,           //16
      qtyTotal,             //17
      avgPrice,             //18
      lastVolume,           //19
      daysPassed            //20
*/

function tradesOneTimeCalculations()
{
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades');
  var tradesCountRange = SpreadsheetApp.getActive().getRangeByName('TradesCount');
  if (tradesRange == null || tradesCountRange == null) return;
  
  var trades = tradesRange.getValues();
  var tradesCount = tradesCountRange.getValue();
  if (tradesCount == "") return;
  
  for (var i = 1; i <= tradesCount; i++)
  {
    var lastPos = 0;
    for (var k = i - 1; k > 0; k--)
    {
      if (trades[k][1] == trades[i][1])
      {
        lastPos = k;
        break;
      }
    }
    var avgPrice = trades[i][3];
    var qtyTotal = trades[i][5];
    var qtyAtStart = 0;
    var lastVolume = 0;
    var daysPassed = "";
    var currency = "";
    var delta = "";
    var deltaPercent = "";
    if(lastPos > 0)
    {
      qtyAtStart = trades[lastPos][17];
      qtyTotal += qtyAtStart;
      if(qtyTotal * trades[i][5] > 0)
      {
        var avg = avgPrice;
        var qty = trades[i][5];
        var lastavg = trades[lastPos][18];
        avgPrice = (avg * qty + lastavg * qtyAtStart) / qtyTotal;
        lastVolume = avgPrice * qtyAtStart;
      }
      else
      {
        avgPrice = trades[lastPos][18];
        lastVolume = avgPrice * qtyAtStart;
        delta = (avgPrice - trades[i][3]) * trades[i][5];
        deltaPercent = delta / Math.abs(lastVolume);
        currency = trades[i][4];
      }
      var time1 = new Date(trades[i][0]).getTime();
      var time2 = new Date(trades[lastPos][0]).getTime();
      daysPassed = (time1 - time2) / mills_per_day;
    }
    trades[i][17] = qtyTotal;
    trades[i][18] = avgPrice;
    tradesRange.offset(i,12,1,9).setValues([[
      delta,        //12
      currency,     //13
      deltaPercent, //14
      lastPos,      //15
      qtyAtStart,   //16
      qtyTotal,     //17
      avgPrice,     //18
      lastVolume,   //19
      daysPassed    //20
      ]]);
  }
  return;
}

function getLastTrades() 
{
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades');
  var tradesCountRange = SpreadsheetApp.getActive().getRangeByName('TradesCount');
  var tradesIdRange = SpreadsheetApp.getActive().getRangeByName('TradesIds')
  if (tradesRange == null || tradesCountRange == null || tradesIdRange == null) return;
  
  var trades = tradesRange.getValues();
  var tradesIds = tradesIdRange.getValues();

  var tradesCount = 0;
  for (var i = 1; i < trades.length; i++)
  {
    if (trades[i][0] == "")
    {
      tradesCount = i - 1;
      break;
    }
  }
  var pivotIds = [];
  for (var i = 1; i < tradesIds.length; i++)
  {
    if (tradesIds[i][0] == "") continue;
    pivotIds.push(tradesIds[i][0]);
  }

  var lastOperationDate = null;
  if (tradesCount != 0) lastOperationDate = trades[tradesCount][0];
  var lastOperations = getOperations(lastOperationDate);

  for (var i = lastOperations.length - 1; i >= 0; i--)
  {
    var lastOp = lastOperations[i];
  //[n.date0, n.figi1, tik2, n.operationType3, n.price4, n.currency5, n.quantityExecuted6, n.payment7, comission.value8, n.status9, n.id10]
    if (lastOp[6] > 0)
    {
      if (pivotIds.includes(+lastOp[10])) continue;
      if (lastOp[5] == 'RUB') lastOp[5] = '₽';
      if (lastOp[5] == 'USD') lastOp[5] = '$';
      if (lastOp[5] == 'EUR') lastOp[5] = '€';
      var volume = lastOp[4] * lastOp[6];
      if (lastOp[3] == 'Sell') lastOp[6] = -lastOp[6];
      
      var date = new Date(lastOp[0]);
      var comiss = comiss = volume * 0.0005;;
      if (date.getTime() - new Date("Apr 22, 2019 03:00:00").getTime() < 0) 
        comiss = volume * 0.0003;
      var summary = volume + comiss;

      tradesRange.offset(++tradesCount,0,1,12).setValues([[
        date,          //0 date
        lastOp[2],     //1 tiker
        tradesCount,   //2 ##
        lastOp[4],     //3 price
        lastOp[5],     //4 curr
        lastOp[6],     //5 qty
        volume,        //6 vol
        lastOp[5],     //7 
        comiss,        //8 comis
        lastOp[5],     //9 
        summary,       //10 summs
        lastOp[5],     //11
      ]]);
      tradesIdRange.getCell(tradesCount,1).setValue(lastOp[10]);
    } 
  }
  tradesCountRange.setValue(tradesCount);
  tradesOneTimeCalculations();
  return;
}

function getLastDividends()
{
  var divRange = SpreadsheetApp.getActive().getRangeByName('Dividends');
  if (divRange == null ) return;
  var dividends = divRange.getValues();

  var operations = null;
  for (var i = 0; i < dividends.length; i++)
  {
    var lastDiv = dividends[i];
    if (lastDiv[1] == "") continue;
    if (lastDiv[2] == "")
    {
      if (operations == null) operations = getOperations(lastDiv[1]); 
      var payoutDate = null;
      var valueSum = 0;
      for (var k = operations.length - 1; k >= 0; k--)
      {
        var lastOp = operations[k];
        if (payoutDate != null)
        {
          var nextPayoutDate = new Date(lastOp[0]);
          if (nextPayoutDate.getTime() - payoutDate.getTime() > mills_per_day)
          {
            divRange.offset(i,2,1,1).setValue(payoutDate);
            divRange.offset(i,13,1,1).setValue(valueSum);
            break;
          }
        }
        if (lastOp[3] != "Dividend") continue;
        if (lastOp[2] == lastDiv[0])
        {
          if (payoutDate != null)
          {
            valueSum += lastOp[7];
            lastOp[2] = "";
            continue;
          }
          payoutDate = new Date(lastOp[0]);
          valueSum += lastOp[7];
          lastOp[2] = "";
        }
      }
    }    
  }
  return;
}

function forcePushOrders()
{
  var knifesRange = SpreadsheetApp.getActive().getRangeByName('KnifeCatcher');
  if (knifesRange == null ) return;
  var knifes = knifesRange.getValues();
  var orderList = ""
  var sum = 0;

  for (var i = 0; i < knifes.length; i++)
  {
    if(knifes[i][0] == "" || knifes[i][1] == "" || knifes[i][2] == "") continue;
    var vol = knifes[i][1] * knifes[i][2];
    sum += vol;
    orderList += "\nКупить: " + knifes[i][1] + "шт. " + knifes[i][0] + " по " + knifes[i][2] + knifes[i][3] + " за " + vol + knifes[i][3];
  }
  orderList += "\nИтого: " + sum + "$";
  Logger.log(orderList);
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Вы собираетесь выставить следующие заявки:\n" + orderList + "\n\nПродолжить?",ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
  {
    for (var i = 0; i < knifes.length; i++)
    {
      if(knifes[i][0] == "" || knifes[i][1] == "" || knifes[i][2] == "") continue;
      pushOrder(knifes[i][0],knifes[i][1],knifes[i][2]);
      }
    }
  return;
}
