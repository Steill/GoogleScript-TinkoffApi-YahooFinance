function onOpen() 
{
  SpreadsheetApp.getUi()
      .createMenu('Tinkoff')
      .addItem('Новый пользователь', 'askForNewUser')
      .addItem('Замена токена', 'askForNewToken')
      .addSeparator()
      .addSeparator()
      .addItem('Перезапись кэша', 'forceUpdateAll')
      .addToUi();
}

function askForNewToken()
{
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Замена токена. Не используйте для замены пользователя!',
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

function askForNewUser()
{
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Внимание! Таблица будет очищена от данных! Продолжить?',
      ui.ButtonSet.YES_NO);
  var button = result.getSelectedButton();
  if (button != ui.Button.YES) return null;

  result = ui.prompt(
      'Установка нового пользователя',
      'Введите токен:',
      ui.ButtonSet.OK_CANCEL);
  button = result.getSelectedButton();
  if (button != ui.Button.OK) return null;
  var newtoken = result.getResponseText();

  result = ui.prompt(
      'Установка даты начала торговли. Используется часовой пояс таблицы.',
      'Варианты форматирования:\n30 dec 2009\n30.12.2009\n2009 30 dec\n2009 dec 30\n2009.12.30',
      ui.ButtonSet.OK_CANCEL);
  button = result.getSelectedButton();
  if (button != ui.Button.OK) return null;
  var firstDay = result.getResponseText();
  
  result = ui.prompt(
      'Установка комиссии (варианты форматирования: 0.05% / 0.0005 / 0,3% / 0,003)',
      'Для сделок до 22 апреля 2020г. при вводе комиссии 0.05% будет использована старая комиссия 0.03%',
      ui.ButtonSet.OK_CANCEL);
  button = result.getSelectedButton();
  if (button != ui.Button.OK) return null;
  var comiss = result.getResponseText();

  props.deleteAllProperties();
  props.setProperty("token", newtoken);
  props.setProperty("firstDayString", firstDay);
  props.setProperty("firstDaySecs", (new Date(firstDay).getTime() / 1000 + Math.floor(Math.random() * 5001)).toString());

  if (comiss.endsWith("%")) comiss = +comiss.substring(0, comiss.indexOf("%")) /100;
  props.setProperty("comiss", +comiss);

  SpreadsheetApp.getActive().getRangeByName('Summaries').clearContent();
  SpreadsheetApp.getActive().getRangeByName('Trades').clearContent();
  SpreadsheetApp.getActive().getRangeByName('Debug').clearContent();
  SpreadsheetApp.getActive().getRangeByName('KnifeCatcher').clearContent();
  SpreadsheetApp.getActive().getRangeByName('Dividends').clearContent();

  SpreadsheetApp.getUi().alert('Установлен новый пользователь. Обновление данных. Может занять несколько минут.');
  
  getDivs();
  getLastTrades();
  getFullOperationList();
  triggerSetup();
  return null;
}

function triggerSetup()
{
  ScriptApp.newTrigger('getDividendsDaily').timeBased().everyDays(1).atHour(5).nearMinute(30).create();
  ScriptApp.newTrigger('getTikersDaily').timeBased().everyDays(1).atHour(5).nearMinute(30).create();
  ScriptApp.newTrigger('getFigisDaily').timeBased().everyDays(1).atHour(5).nearMinute(30).create();
  ScriptApp.newTrigger('getDivs').timeBased().everyDays(1).atHour(5).nearMinute(30).create();

  ScriptApp.newTrigger('getFullOperationList').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('getLastDividends').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('getLastTrades').timeBased().everyHours(1).create();
}

function forceUpdateAll()
{
  getDividendsDaily();
  getFigisDaily();
  getTikersDaily();
  SpreadsheetApp.getUi().alert('Принудительное обновление кэша завершено');
  return null;
}
function getLastTradesUi()
{
  getLastTrades();
  SpreadsheetApp.getUi().alert('Список операций обновлён');
  return null;
}

function getLastDividendsUi()
{
  getLastDividends();
  SpreadsheetApp.getUi().alert('Список дивидендов обновлён');
  return null;
}
