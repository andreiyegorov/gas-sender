function createSenderPostmanTrigger() {

  // удаляем старые триггеры, если есть
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'senderPostmanTrigger') {
      ScriptApp.deleteTrigger(tr);
    }
  });

  // создаём новый
  ScriptApp.newTrigger('senderPostmanTrigger')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert('✅ Триггер создан: каждые 30 минут.');
}