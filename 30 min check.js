/*******************************************************
 *  ⏰ СОЗДАНИЕ АВТО-ТРИГГЕРА
 *  📅 Версия: 2711-0505 (27 ноября 05:05)
 *  📘 Назначение:
 *     Создаёт триггер autoCheck30min каждые 30 минут
 *******************************************************/
function createAutoTrigger() {
  // удалить старые триггеры (все варианты)
  ScriptApp.getProjectTriggers().forEach(tr => {
    const fn = tr.getHandlerFunction();
    if (fn === 'autoCheck30min' || fn === 'senderPostmanAuto' || fn === 'senderPostmanTrigger') {
      ScriptApp.deleteTrigger(tr);
    }
  });

  // создать новый — каждые 30 минут
  ScriptApp.newTrigger('autoCheck30min')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert('✅ Авто-триггер создан: autoCheck30min каждые 30 минут.');
}
