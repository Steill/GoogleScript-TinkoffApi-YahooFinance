function onOpen() 
{
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Tinkoff')
      .addItem('Ввести токен', 'askForNewToken')
      .addSeparator()
      .addItem('Получить список операций', 'getLastTradesUi')
      .addItem('Получить список дивидендов', 'getLastDividendsUi')
      .addItem('Пересчитать листы', 'tradesOneTimeCalculationsUi')
      .addSeparator()
      .addSeparator()
      .addItem('Принудительное обновление', 'forceUpdateAll')
      .addSeparator()
      .addItem('Выставить заявки', 'forcePushOrders')
      .addToUi();
}

function askForNewToken()
{
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Установка нового пользователя',
      'Введите токен:',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  if (button != ui.Button.OK) return null;
  var newtoken = result.getResponseText();
  props.deleteProperty("token");
  props.setProperty("token", newtoken);
  //cache.put("token", newtoken);
  SpreadsheetApp.getUi().alert('Установлен новый токен');
  return null;
}

function forceUpdateAll()
{
  getDividendsDaily();
  getFigisDaily();
  getTikersDaily();
  SpreadsheetApp.getUi().alert('Принудительное обновление кеша завершено');
}
function getLastTradesUi()
{
  getLastTrades();
  SpreadsheetApp.getUi().alert('Список операций обновлён');
}

function getLastDividendsUi()
{
  getLastDividends();
  SpreadsheetApp.getUi().alert('Список дивидендов обновлён');
}

function tradesOneTimeCalculationsUi()
{
  tradesOneTimeCalculations();
  SpreadsheetApp.getUi().alert('Листы пересчитаны');
}
