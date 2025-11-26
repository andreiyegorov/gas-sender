/*******************************************************
 *  📋 MENU — Меню таблицы
 *  📅 Версия: 2711-0505 (27 ноября 05:05)
 *  📘 Назначение:
 *     Создаёт меню "Sender Postman" в таблице
 *******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📨 Sender Postman');
  menu.addItem('🚀 Отправить новые отчёты', 'senderPostmanRun');
  menu.addItem('🔍 Проверить новые отчёты (без отправки)', 'senderPostmanCheck');
  menu.addItem('🔗 Проверить новые ссылки (CSV)', 'senderInviteCheck');
  menu.addSeparator();
  menu.addItem('📊 Обновить Dashboard', 'refreshDashboard');
  menu.addItem('🔄 Пересоздать Dashboard', 'resetDashboard');
  menu.addSeparator();
  menu.addItem('💰 Проверить операции СберБизнес', 'parseSberOperations');
  menu.addToUi();
}