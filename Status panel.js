/*******************************************************
 *  🟦 SENDER POSTMAN — STATUS PANEL
 *  Обновляет K1–K6 на листе sender-log
 *******************************************************/

var STATUS_SHEET = 'sender-log';

/*******************************************************
 * Обновление панели статусов (автомат)
 * Вызывается из senderPostman_
 *******************************************************/
function updateSenderAutoStatusPanel_(idsSent, runDate) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STATUS_SHEET);
  if (!sh) return;

  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');

  // ----- K2 и K3 (автоматические проверки)
  var lastAuto = sh.getRange('K2').getValue();
  if (lastAuto) sh.getRange('K3').setValue(lastAuto);
  sh.getRange('K2').setValue(timeStr);

  // ----- K1 (Last sent)
  if (idsSent && idsSent.length > 0) {
    sh.getRange('K1').setValue(idsSent.join(', ') + ' | ' + timeStr);
  }

  // ----- K4 (Next check in) — расчёт вручную, т.к. Google убрал getNextRunTime
  var nextTrigger = getNextTriggerTime_(runDate);
  if (nextTrigger) {
    var tStr = Utilities.formatDate(nextTrigger, tz, 'HH:mm');
    var diffMin = Math.floor((nextTrigger - runDate) / 60000);
    if (diffMin < 1) diffMin = '<1';
    sh.getRange('K4').setValue(tStr + ' (через ' + diffMin + ' мин)');
  }
}

/*******************************************************
 * Обновление панели статусов (ручные проверки)
 * Вызывается из senderPostmanRun / senderPostmanCheck
 *******************************************************/
function updateSenderManualStatus_(functionName, runDate) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STATUS_SHEET);
  if (!sh) return;

  var tz = Session.getScriptTimeZone();
  var timeStr = Utilities.formatDate(runDate, tz, 'dd.MM HH:mm');

  // ----- K5 & K6
  var lastManual = sh.getRange('K5').getValue();
  if (lastManual) sh.getRange('K6').setValue(lastManual);

  sh.getRange('K5').setValue(functionName + ' | ' + timeStr);
}

/*******************************************************
 * ❗ ВАЖНО: Google удалил getNextRunTime()
 * Поэтому рассчитываем время следующего запуска сами:
 * триггер → каждые 30 минут
 *******************************************************/
function getNextTriggerTime_(runDate) {
  var next = new Date(runDate.getTime() + 30 * 60000);
  return next;
}