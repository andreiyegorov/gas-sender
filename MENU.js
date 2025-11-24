/*******************************************************
 *  📨 SENDER POSTMAN — MENU (simple working)
 *  Версия: 2025-11 18.05
 *******************************************************/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📨 Sender Postman');

  menu.addItem('🚀 Отправить новые отчёты', 'senderPostmanRun');
  menu.addItem('🔍 Проверить новые отчёты', 'senderPostmanCheck');
  menu.addItem('🔗 Проверить новые ссылки (CSV)', 'senderFetcherCheck');

  menu.addSeparator();

  menu.addItem('📊 Открыть лог SENDER', 'openSenderLog');
  menu.addItem('📥 Открыть лог LINKS', 'openReceivedLog');

  menu.addToUi();
}

/* открытие логов */
function openSenderLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('sender-log');
  const ui = SpreadsheetApp.getUi();
  if (sh) {
    ss.setActiveSheet(sh);
    ui.alert('📊 Лог SENDER открыт.');
  } else {
    ui.alert('⚠️ Лист "sender-log" ещё не создан.');
  }
}

function openReceivedLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('received-log');
  const ui = SpreadsheetApp.getUi();
  if (sh) {
    ss.setActiveSheet(sh);
    ui.alert('📥 Лог LINKS (received-log) открыт.');
  } else {
    ui.alert('⚠️ Лист "received-log" ещё не создан.');
  }
}

/* вызов Fetcher */
function senderFetcherCheck() {
  senderFetcher_('menu-check');
}