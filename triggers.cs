function getDividendsDaily()
{
  Logger.log(SpreadsheetApp.getActive().getId());
  var summaries = SpreadsheetApp.getActive().getRangeByName('Summaries');
  if (summaries == null) return null;

  summaries = summaries.getValues();
  var tikersString = "";
  for (var i = 0; i < summaries.length; i++)
  {
    var tiker = summaries[i][0];
    if (tiker == "") break;
    //if (tiker[0] == "_") continue;
    if (summaries[i][5] == "₽") continue;
    tikersString += tiker + " ";
  }
  Logger.log(tikersString);
  getDividends(tikersString, true);
  return null;
}

function getFigisDaily()
{
  Logger.log(SpreadsheetApp.getActive().getId());
  var summaries = SpreadsheetApp.getActive().getRangeByName('Summaries');
  if (summaries == null) return null;

  summaries = summaries.getValues();
  for (var i = 0; i < summaries.length; i++)
  {
    var tiker = summaries[i][0];
    if (tiker == "") break;
    if (tiker[0] == "_") continue;
    //Logger.log(tiker);
    getFigi(tiker,true);
    //getTiker(getFigi(tiker,true),true);
  }
  return null;
}

function getTikersDaily()
{
  Logger.log(SpreadsheetApp.getActive().getId());
  var summaries = SpreadsheetApp.getActive().getRangeByName('Summaries');
  if (summaries == null) return null;
  
  summaries = summaries.getValues();
  for (var i = 0; i < summaries.length; i++)
  {
    var tiker = summaries[i][0];
    if (tiker == "") break;
    if (tiker[0] == "_") continue;
    //Logger.log(tiker);
    getTiker(getFigi(tiker),true);
  }
  return null;
}

function pushOrdersToServerDaily()
{
  return;
}
