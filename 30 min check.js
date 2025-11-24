function createSenderPostmanTrigger() {

  // удалить старые триггеры на авто-функцию
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'senderPostmanAuto') {
      ScriptApp.deleteTrigger(tr);
    }
  });

  // создать новый — каждые 30 минут
  ScriptApp.newTrigger('senderPostmanAuto')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert('✅ Авто-триггер создан: каждые 30 минут.');
}