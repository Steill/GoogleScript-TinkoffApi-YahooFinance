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
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades');
  var divsRange = SpreadsheetApp.getActive().getRangeByName('Dividends');
  if (divsRange == null || tradesRange == null) return null;

  var tikers = [];
  var tradesCount = 0;
  var divsCount = 0;
  var tikersString = "";  
  var trades = tradesRange.getValues();
  for (var i = 1; i < trades.length; i++)
  {
    if (trades[i][0] == "")
    {
      tradesCount = i - 1;
      break;
    }
    var tiker = trades[i][1];
    if (trades[i][4] == "₽" || tikers.includes(tiker)) continue;
    tikers.push(tiker);
    tikersString += tiker + " ";
  }
  if (tradesCount == 0) return;

  var divs = getDividends(tikersString).sort(function(a,b){
    var A = new Date(a[2]).getTime();
    var B = new Date(b[2]).getTime();
    if (B < A) return 1;
    if (A < B) return -1;
    return 0;});
  for (var i = 0; i < divs.length; i++)
  {
    var divTime = new Date(divs[i][2]);
    var count = "";
    var avgPrice = ""; 
    var currency = ""; 
    for (var k = 1; k <= tradesCount; k++)
    {
      if (new Date(trades[k][0]).getTime() <  divTime.getTime()) continue;
      var lastN = k - 1;
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
      break;
    }
    if (count != 0)
    {
      var divAmount = +divs[i][3].replace(',','.');
      var nrange = divsRange.offset(++divsCount,0,1,2).setValues([[
        divs[i][1],
        divTime
      ]]);
      nrange.offset(0,3,1,10).setValues([[
        count,
        avgPrice,
        currency,
        avgPrice * count,
        currency,
        divAmount,
        currency,
        divAmount * count,
        currency,
        divAmount / avgPrice
      ]]);
      
      if (divsRange.getLastRow() - divsCount < 3)
        divsRange.getSheet().insertRowsBefore(divsRange.getLastRow(),5);
    }
  }
  return;
}

function tradesOneTimeCalculations()
{
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades');
  if (tradesRange == null ) return;
  
  var trades = tradesRange.getValues();  
  for (var i = 1; i < trades.length; i++)
  {
    if (trades[i][1] == "") return; 

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
  if (tradesRange == null) return;
  
  var pivotIds = [];
  var tradesCount = 0;
  var lastOperationDate = null;
  var trades = tradesRange.getValues();
  for (var i = 1; i < trades.length; i++)
  {
    if (trades[i][0] == "")
    {
      tradesCount = i - 1;
      var pivotTime = new Date(trades[tradesCount][0]).getTime();
      for (var k = tradesCount - 1; k > 0; k--)
      {
        var newpivotTime = new Date(trades[k][0]).getTime();;
        if (pivotTime - newpivotTime > mills_per_day) break;
        tradesCount = k;
      }
      if (tradesCount > 0) lastOperationDate = trades[tradesCount][0];
      for (var k = tradesCount; k > 0 && tradesCount - k < 10; k--) 
        pivotIds.push(trades[k][21]);
      break;
    }
  }
  
  var lastOperations = getOperations(lastOperationDate);
  if (lastOperations == null) return null;

  for (var i = lastOperations.length - 1; i >= 0; i--)
  {
    var lastOp = lastOperations[i];
    if (lastOp[6] > 0)
    {
      if (pivotIds.includes(lastOp[10])) continue;
      if (lastOp[5] == 'RUB') lastOp[5] = '₽';
      if (lastOp[5] == 'USD') lastOp[5] = '$';
      if (lastOp[5] == 'EUR') lastOp[5] = '€';
      var volume = lastOp[4] * lastOp[6];
      if (lastOp[3] == 'Sell') lastOp[6] = -lastOp[6];
      
      var date = new Date(lastOp[0]);
      var comiss = volume * props.getProperty("comiss");
      if (props.getProperty("comiss") == 0.0005)
      {
        if (date.getTime() - new Date("Apr 22, 2019").getTime() < 0) 
          comiss = volume * 0.0003;
      }
      var summary = volume + comiss;
      
      if (tradesRange.getLastRow() - tradesCount < 10)
        tradesRange.getSheet().insertRowsBefore(tradesRange.getLastRow(),10);
      
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
      tradesRange.offset(tradesCount,21,1,1).setValue(lastOp[10]);
    } 
  }
  tradesOneTimeCalculations();
  return;
}

function getLastDividends()
{
  var divsRange = SpreadsheetApp.getActive().getRangeByName('Dividends');
  if (divsRange == null) return;
  var dividends = divsRange.getValues();

  for (var i = 1; i < dividends.length; i++)
  {
    var lastDiv = dividends[i];
    if (lastDiv[1] == "") continue;
    if (lastDiv[2] == "")
    {
      var valueSum = 0;
      var payoutDate = null;
      var opTimeShifted = new Date(new Date(lastDiv[1]).getTime() + mills_per_day * 75);
      var operations = getOperations(lastDiv[1], opTimeShifted); 
      for (var k = operations.length - 1; k >= 0; k--)
      {
        var lastOp = operations[k];
        if (payoutDate != null)
        {
          var nextOpDate = new Date(lastOp[0]);
          if (nextOpDate.getTime() - payoutDate.getTime() > mills_per_day)
          {
            divsRange.offset(i,2,1,1).setValue(payoutDate);
            divsRange.offset(i,13,1,4).setValues([[valueSum, lastDiv[5], lastDiv[10] - valueSum, lastDiv[5]]]);
            break;
          }
        }
        if (lastOp[3] != "Dividend") continue;
        if (lastOp[2] == lastDiv[0])
        {
          if (payoutDate == null) payoutDate = new Date(lastOp[0]);
          valueSum += lastOp[7];
        }
      }
    }
    else 
      if (lastDiv[13] != "" && lastDiv[14] == "")
        divsRange.offset(i,14,1,3).setValues([[lastDiv[5], lastDiv[10] - lastDiv[13], lastDiv[5]]]);
  }
  return;
}

function getFullOperationList()
{
  var debugRange = SpreadsheetApp.getActive().getRangeByName('Debug');
  if (debugRange == null ) return;
  
  var pivotIds = [];
  var opCount = 0;
  var lastOperationDate = null;
  var operationList = debugRange.getValues(); 
  for (var i = 1; i < operationList.length; i++)
  {
    if (operationList[i][0] == "")
    {
      opCount = i - 3;
      if (opCount < 0) opCount = 0;
      var date = operationList[opCount][0];
      if (opCount > 0 && date != "")
      {
        lastOperationDate = date;
        for (var k = opCount; k > 0 && opCount - k < 10; k--) 
        pivotIds.push(operationList[k][10]);
      }
      break;
    }
  }
  var lastOperations = getOperations(lastOperationDate);
  if (lastOperations == null) return null;
  
  debugRange.getSheet().insertRowsBefore(debugRange.getLastRow(), lastOperations.length);
  for (var i = lastOperations.length - 1; i >= 0; i--)
  {
    var lastOp = lastOperations[i];
    if (pivotIds.includes(lastOp[10])) continue;
    debugRange.offset(++opCount,0,1,11).setValues([lastOp]);
  }
  return;  
}

function forcePushOrders()
{
  var knifesRange = SpreadsheetApp.getActive().getRangeByName('KnifeCatcher');
  if (knifesRange == null ) return;
  var knifes = knifesRange.getValues();
  var orderConfirmString = "";
  var orders = [];
  var sumUSD = 0;
  var sumRUB = 0;

  for (var i = 0; i < knifes.length; i++)
  {
    if(knifes[i][1] == "" || knifes[i][2] == "" || knifes[i][3] == "" || knifes[i][4] == "") continue;
    
    var type;
    if (knifes[i][0].toLowerCase() == "купить") type = "Buy";
    else 
      if (knifes[i][0].toLowerCase() == "продать") type = "Sell";
        else continue;

    var vol = knifes[i][2] * knifes[i][3];
    if (knifes[i][4] == "$") sumUSD += vol;
    if (knifes[i][4] == "₽") sumRUB += vol;
    orderConfirmString += "\n"+ knifes[i][0] + ": " + knifes[i][2] + "шт. " + knifes[i][1] + " по " + knifes[i][3] + knifes[i][4] + " за " + vol + knifes[i][4];
    orders.push([knifes[i][1], knifes[i][2], knifes[i][3], type]);
  }
  orderConfirmString += "\n\nИтого: ";
  if (sumUSD != 0)
  {
    orderConfirmString += sumUSD + "$";
    if (sumRUB != 0)
      orderConfirmString += "\n             " + sumRUB + "₽"
  }
  else orderConfirmString += sumRUB + "₽";
  Logger.log(orderConfirmString);
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Вы собираетесь выставить следующие заявки:\n" + orderConfirmString + "\n\nПродолжить?",ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
    for (var i = 0; i < orders.length; i++)
      pushOrder(orders[i][0],orders[i][1],orders[i][2],orders[i][3]);
  return;
}
