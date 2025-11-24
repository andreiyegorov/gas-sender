function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📨 Sender Postman');
  menu.addItem('🚀 Отправить новые отчёты', 'senderPostmanRun');
  menu.addItem('🔍 Проверить новые отчёты (без отправки)', 'senderPostmanCheck');
  menu.addItem('📥 Проверить новые ссылки (CSV)', 'senderFetcherMenu');
  menu.addItem('💰 Проверить операции СберБизнес', 'checkSberOperationsMenu');
  menu.addToUi();
}