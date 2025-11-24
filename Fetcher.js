/*******************************************************
 *  📨 SENDER FETCHER — ловец CSV-приглашений
 *  📅 Версия: 2025-11-24 (восстановлено)
 *******************************************************/

/*******************
 * 1. КОНСТАНТЫ
 *******************/
var FETCHER_LABEL_NAME = 'sender_fetcher_newlinks';
var LINKS_SHEET_NAME   = 'Links';
var LOG_SHEET_NAME     = 'received-log';

/*******************
 * 2. ЗАПУСК ИЗ МЕНЮ
 *******************/
function senderInviteCheck() {
  senderFetcher_('menu-check');
}

/*******************
 * 3. ОСНОВНАЯ ФУНКЦИЯ
 *******************/
function senderFetcher_(source) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var nowStr = Utilities.formatDate(now, tz, 'dd.MM.yy HH.mm');

  // --- гарантируем наличие листов ---
  var sheetLinks = ss.getSheetByName(LINKS_SHEET_NAME);
  if (!sheetLinks) {
    sheetLinks = ss.insertSheet(LINKS_SHEET_NAME);
    sheetLinks.appendRow(['User ID','Password','Group Name','Email TS','Check TS']);
  }

  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(['Timestamp','Source','File','Rows added','Status','Error']);
  }

  // --- метка и поиск писем ---
  var label = GmailApp.getUserLabelByName(FETCHER_LABEL_NAME) || GmailApp.createLabel(FETCHER_LABEL_NAME);
  var threads = GmailApp.search('in:inbox newer_than:30d has:attachment');

  // --- обработка писем ---
  for (var t = threads.length - 1; t >= 0; t--) {
    var thread = threads[t];
    if (thread.getLabels().some(l => l.getName() === FETCHER_LABEL_NAME)) continue;
    var msgs = thread.getMessages();

    for (var m = 0; m < msgs.length; m++) {
      var msg = msgs[m];
      var dateStr = Utilities.formatDate(msg.getDate(), tz, 'dd.MM.yy HH.mm');
      var atts = msg.getAttachments();
      if (!atts || atts.length === 0) continue;

      for (var a = 0; a < atts.length; a++) {
        var att = atts[a];
        var filename = att.getName();
        if (!att.getContentType().match(/csv/i) && !filename.match(/\.csv$/i)) continue;

        var csv;
        try { csv = Utilities.parseCsv(att.getDataAsString()); } catch(e) { continue; }
        if (!csv || csv.length < 2) continue;

        var rows = [];
        for (var r = 1; r < csv.length; r++) {
          var u = (csv[r][0] || '').trim();
          var p = (csv[r][1] || '').trim();
          var g = (csv[r][2] || '').trim();
          if (u && p) rows.push({u,p,g,date:dateStr,check:nowStr});
        }
        rows.sort((a,b)=>parseInt(b.u.replace(/\D+/g,'')) - parseInt(a.u.replace(/\D+/g,'')));

        var added = 0;
        for (var i = 0; i < rows.length; i++) {
          var exists = sheetLinks.createTextFinder(rows[i].u).matchCase(false).findNext();
          if (!exists) {
            sheetLinks.insertRowBefore(2);
            sheetLinks.getRange(2,1,1,5).setValues([[rows[i].u,rows[i].p,rows[i].g,rows[i].date,rows[i].check]]);
            added++;
          }
        }

        thread.addLabel(label);
        if (added > 0) {
          logSheet.insertRowBefore(2);
          logSheet.getRange(2,1,1,6).setValues([[nowStr,source,filename,added,'✅ received','']]);
        }
      }
    }
  }
}