/*******************************************************
 *  ☑️ CHECKBOX — автопростановка даты при галочке
 *  📅 Версия: 2511-0345 (25 ноября 03:45)
 *  📘 Назначение:
 *     Автоматически проставляет дату/время в ячейку справа
 *     при включении чекбокса.
 *  📦 Источник: UK Sales Automation
 *******************************************************/

/*******************************************************
 *  1️⃣ АВТОМАТИЧЕСКАЯ ПРОСТАВКА ДАТЫ/ВРЕМЕНИ ПРИ ЧЕКБОКСЕ
 *******************************************************/
function onEdit(e) {
  try {
    var sheet, range;

    // 1.1. Определяем, как вызвана функция: автоматически или вручную
    if (e && e.range) {
      range = e.range;
      sheet = range.getSheet();
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      range = sheet.getActiveRange();
      Logger.log('onEdit: ручной запуск (test mode)');
    }

    if (!sheet) return;

    // 1.4. Обрабатываем только одиночные ячейки
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

    // 1.5. Проверяем, что редактируется чекбокс
    var dv = range.getDataValidation();
    var isCheckbox = false;
    if (dv) {
      try {
        isCheckbox =
          (dv.getCriteriaType && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) ||
          (dv.getCriteriaValues && dv.getCriteriaValues().toString().indexOf('TRUE') !== -1);
      } catch (err2) {
        // fallback — проверим значение напрямую
      }
    }

    var val = range.getValue();
    if (!isCheckbox && val !== true) return;
    if (val !== true) return; // реагируем только при включении галочки

    // 1.6. Определяем ячейку справа
    var row = range.getRow();
    var col = range.getColumn();
    var rightCol = col + 1;
    if (rightCol > sheet.getMaxColumns()) sheet.insertColumnAfter(col);
    var rightCell = sheet.getRange(row, rightCol);

    // 1.7. Если справа пусто — вставляем дату и время
    var rv = rightCell.getValue();
    if (rv === '' || rv === null) {
      var now = new Date();
      rightCell.setValue(now);
      rightCell.setNumberFormat('dd.MM HH:mm'); // формат "25.03 15:30"
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/*******************************************************
 *  2️⃣ АВТОРИЗАЦИЯ (однократный ручной запуск)
 *******************************************************/
function authorize() {
  SpreadsheetApp.getActive();
}

