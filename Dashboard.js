/*******************************************************
 *  📊 DASHBOARD
 *  📅 Версия: 2611-1230 (26 ноября 12:30)
 *  📘 Назначение:
 *     Dashboard с двумя колонками: FETCHER и SENDER
 *******************************************************/

var DASHBOARD_SHEET = 'Dashboard';

/*******************************************************
 *  1️⃣ ПОЛУЧИТЬ ИЛИ СОЗДАТЬ ЛИСТ
 *******************************************************/
function getDashboardSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DASHBOARD_SHEET);
  
  if (!sh) {
    sh = ss.insertSheet(DASHBOARD_SHEET);
    setupDashboardStructure_(sh);
  }
  
  return sh;
}

/*******************************************************
 *  2️⃣ НАСТРОЙКА СТРУКТУРЫ (только при создании)
 *******************************************************/
function setupDashboardStructure_(sh) {
  // Заголовки
  sh.getRange('B1').setValue('FETCHER').setFontWeight('bold');
  sh.getRange('C1').setValue('SENDER').setFontWeight('bold');
  
  // Статусы
  sh.getRange('A2').setValue('LAST Received/Sent');
  sh.getRange('A3').setValue('LAST CHECK AUTO');
  sh.getRange('A4').setValue('PREV CHECK AUTO');
  sh.getRange('A5').setValue('LAST CHECK MENU');
  sh.getRange('A6').setValue('PREV CHECK MENU');
  sh.getRange('A7').setValue('NEXT CHECK IN');
  
  // Статистика
  sh.getRange('A9').setValue('↘️ IDs received').setFontWeight('bold');
  sh.getRange('A10').setValue('✍️ IDs registered').setFontWeight('bold');
  sh.getRange('A11').setValue('↗️ IDs sent to client').setFontWeight('bold');
  sh.getRange('A12').setValue('↘️ REPORTS received').setFontWeight('bold');
  sh.getRange('A13').setValue('↗️ REPORTS sent').setFontWeight('bold');
  
  // Ширина колонок
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 150);
  
  // Форматируем все ячейки данных как текст
  sh.getRange('B2:C13').setNumberFormat('@');
}

/*******************************************************
 *  3️⃣ ОБНОВИТЬ СТАТИСТИКУ
 *******************************************************/
function updateDashboardStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getDashboardSheet_();
  
  var linksSheet = ss.getSheetByName('Links');
  
  // ↘️ IDs received
  var idsReceived = 0;
  if (linksSheet) {
    idsReceived = Math.max(0, linksSheet.getLastRow() - 1);
  }
  sh.getRange('B9').setValue(String(idsReceived));
  
  // ✍️ IDs registered
  var idsRegistered = countNamedRange_(ss, 'registered');
  sh.getRange('B10').setValue(String(idsRegistered));
  
  // ↗️ IDs sent to client
  var idsSentToClient = countNamedRangeChecked_(ss, 'id_sent_count');
  sh.getRange('B11').setValue(String(idsSentToClient));
  
  // ↘️ REPORTS received
  var reportsReceived = countLabeledEmails_('sender_postman_done');
  sh.getRange('B12').setValue(String(reportsReceived));
  
  // ↗️ REPORTS sent
  var reportsSent = countNamedRangeChecked_(ss, 'report_sent');
  sh.getRange('B13').setValue(String(reportsSent));
}

/*******************************************************
 *  4️⃣ ПОДСЧЁТ НЕПУСТЫХ ЯЧЕЕК В ДИАПАЗОНЕ
 *******************************************************/
function countNamedRange_(ss, rangeName) {
  try {
    var range = ss.getRangeByName(rangeName);
    if (!range) return 0;
    var values = range.getValues();
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] !== '' && values[i][0] !== null) {
        count++;
      }
    }
    return count;
  } catch(e) {
    return 0;
  }
}

/*******************************************************
 *  5️⃣ ПОДСЧЁТ ЧЕКБОКСОВ TRUE ИЛИ ТЕКСТА "SENT"
 *******************************************************/
function countNamedRangeChecked_(ss, rangeName) {
  try {
    var range = ss.getRangeByName(rangeName);
    if (!range) return 0;
    var values = range.getValues();
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      var val = values[i][0];
      if (val === true || (typeof val === 'string' && val.toLowerCase().includes('sent'))) {
        count++;
      }
    }
    return count;
  } catch(e) {
    return 0;
  }
}

/*******************************************************
 *  6️⃣ ПОДСЧЁТ ПИСЕМ С ЯРЛЫКОМ
 *******************************************************/
function countLabeledEmails_(labelName) {
  try {
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) return 0;
    return label.getThreads().length;
  } catch(e) {
    return 0;
  }
}

/*******************************************************
 *  7️⃣ FETCHER: записать результат
 *******************************************************/
function updateDashboardLastFetched_(count, runDate) {
  var sh = getDashboardSheet_();
  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');
  
  sh.getRange('B2').setValue(count + ' | ' + timeStr);
  updateDashboardStats_();
}

/*******************************************************
 *  8️⃣ FETCHER: записать время проверки
 *******************************************************/
function updateDashboardFetcherCheck_(source, runDate) {
  var sh = getDashboardSheet_();
  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');
  
  if (source === 'trigger') {
    var lastAuto = sh.getRange('B3').getValue();
    if (lastAuto) sh.getRange('B4').setValue(lastAuto);
    sh.getRange('B3').setValue(timeStr);
    
    var nextTime = new Date(runDate.getTime() + 30 * 60000);
    var nextStr = Utilities.formatDate(nextTime, tz, 'HH:mm');
    sh.getRange('B7').setValue(nextStr + ' (через 30 мин)');
  } else {
    var lastMenu = sh.getRange('B5').getValue();
    if (lastMenu) sh.getRange('B6').setValue(lastMenu);
    sh.getRange('B5').setValue(timeStr);
  }
}

/*******************************************************
 *  9️⃣ SENDER: записать отправленные ID
 *******************************************************/
function updateDashboardLastSent_(idsSent, runDate) {
  var sh = getDashboardSheet_();
  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');
  
  if (idsSent && idsSent.length > 0) {
    var lastId = idsSent[idsSent.length - 1];
    sh.getRange('C2').setValue(lastId + ' | ' + timeStr);
  }
  
  updateDashboardStats_();
}

/*******************************************************
 *  🔟 SENDER: записать авто-проверку
 *******************************************************/
function updateDashboardAutoStatus_(runDate) {
  var sh = getDashboardSheet_();
  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');
  
  var lastAuto = sh.getRange('C3').getValue();
  if (lastAuto) sh.getRange('C4').setValue(lastAuto);
  sh.getRange('C3').setValue(timeStr);
  
  var nextTime = new Date(runDate.getTime() + 30 * 60000);
  var nextStr = Utilities.formatDate(nextTime, tz, 'HH:mm');
  sh.getRange('C7').setValue(nextStr + ' (через 30 мин)');
}

/*******************************************************
 *  1️⃣1️⃣ SENDER: записать ручную проверку
 *******************************************************/
function updateDashboardManualStatus_(functionName, runDate) {
  var sh = getDashboardSheet_();
  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');
  
  var lastMenu = sh.getRange('C5').getValue();
  if (lastMenu) sh.getRange('C6').setValue(lastMenu);
  sh.getRange('C5').setValue(timeStr);
}

/*******************************************************
 *  1️⃣2️⃣ МЕНЮ: Обновить Dashboard
 *******************************************************/
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DASHBOARD_SHEET);
  
  // Если лист не существует — создаём
  if (!sh) {
    sh = ss.insertSheet(DASHBOARD_SHEET);
    setupDashboardStructure_(sh);
  }
  
  // Обновляем только статистику (не трогаем статусы)
  updateDashboardStats_();
  
  SpreadsheetApp.getUi().alert('✅ Dashboard обновлён!');
}

/*******************************************************
 *  1️⃣3️⃣ МЕНЮ: Пересоздать Dashboard с нуля
 *******************************************************/
function resetDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DASHBOARD_SHEET);
  
  if (sh) {
    sh.clear();
    setupDashboardStructure_(sh);
  } else {
    sh = ss.insertSheet(DASHBOARD_SHEET);
    setupDashboardStructure_(sh);
  }
  
  updateDashboardStats_();
  SpreadsheetApp.getUi().alert('✅ Dashboard пересоздан!');
}
