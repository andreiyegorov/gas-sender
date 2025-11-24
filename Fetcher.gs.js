/*******************************************************
 *  📨 SENDER FETCHER — ловец CSV-приглашений
 *  📅 Версия: 2411-2340 (24 ноября 23:40)
 *  📘 Назначение:
 *     Извлекает CSV-вложения из писем с приглашениями Hogan.
 *     Добавляет новые строки в лист "Links" и ведёт журнал изменений.
 *  🔧 Изменения:
 *     • Исправлена сортировка писем (старые → новые)
 *     • Ссылки внутри письма сортируются по убыванию ID
 *     • Добавлен пункт меню "🔗 Проверить новые ссылки"
 *******************************************************/

/*******************************************************
 *  1️⃣ КОНСТАНТЫ
 *******************************************************/
var FETCHER_LABEL_NAME = 'sender_fetcher_newlinks';
var LINKS_SHEET_NAME   = 'Links';
var LOG_SHEET_NAME     = 'received-log';

/*******************************************************
 *  2️⃣ МЕНЮ
 *******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📨 Sender Postman');
  menu.addItem('🚀 Отправить новые отчёты', 'senderPostmanRun');
  menu.addItem('🔍 Проверить новые отчёты (без отправки)', 'senderPostmanCheck');
  menu.addSeparator();
  menu.addItem('🔗 Проверить новые ссылки (CSV)', 'senderInviteCheck');
  menu.addToUi();
}

/*******************************************************
 *  3️⃣ ЗАПУСК FETCHER
 *******************************************************/
function senderInviteCheck() {
  senderFetcher_('menu-check');
}

/*******************************************************
 *  4️⃣ ОСНОВНАЯ ФУНКЦИЯ
 *******************************************************/
function senderFetcher_(source) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = Session.getScriptTimeZone();
  const now = new Date();
  const nowStr = Utilities.formatDate(now, tz, 'dd.MM.yy HH.mm');

  /*******************
   * 4.1 Гарантируем наличие листов
   *******************/
  let sheetLinks = ss.getSheetByName(LINKS_SHEET_NAME);
  if (!sheetLinks) {
    sheetLinks = ss.insertSheet(LINKS_SHEET_NAME);
    sheetLinks.appendRow(['User ID','Password','Group Name','Email TS','Check TS']);
  }

  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(['Timestamp','Source','File','Rows added','Status','Error']);
  }

  /*******************
   * 4.2 Метка и поиск писем
   *******************/
  const label = GmailApp.getUserLabelByName(FETCHER_LABEL_NAME) || GmailApp.createLabel(FETCHER_LABEL_NAME);
  const threads = GmailApp.search('in:inbox newer_than:30d has:attachment');

  /*******************
   * 4.3 Обработка писем (старые → новые)
   *******************/
  for (let t = threads.length - 1; t >= 0; t--) {
    const thread = threads[t];
    if (thread.getLabels().some(l => l.getName() === FETCHER_LABEL_NAME)) continue;
    const msgs = thread.getMessages();

    for (let m = 0; m < msgs.length; m++) {
      const msg = msgs[m];
      const dateStr = Utilities.formatDate(msg.getDate(), tz, 'dd.MM.yy HH.mm');
      const atts = msg.getAttachments();
      if (!atts || atts.length === 0) continue;

      /*******************
       * 4.4 Обработка CSV-вложений
       *******************/
      for (let a = 0; a < atts.length; a++) {
        const att = atts[a];
        const filename = att.getName();
        if (!att.getContentType().match(/csv/i) && !filename.match(/\.csv$/i)) continue;

        let csv;
        try { csv = Utilities.parseCsv(att.getDataAsString()); } catch(e) { continue; }
        if (!csv || csv.length < 2) continue;

        /*******************
         * 4.5 Подготовка строк — сортировка ID по убыванию
         *******************/
        let rows = [];
        for (let r = 1; r < csv.length; r++) {
          const u = (csv[r][0] || '').trim();
          const p = (csv[r][1] || '').trim();
          const g = (csv[r][2] || '').trim();
          if (u && p) rows.push([u,p,g,dateStr,nowStr]);
        }

        rows.sort((a,b)=>parseInt(b[0].replace(/\D+/g,'')) - parseInt(a[0].replace(/\D+/g,'')));

        /*******************
         * 4.6 Запись в Links (новые сверху)
         *******************/
        let newRows = [];
        for (let i = 0; i < rows.length; i++) {
          const exists = sheetLinks.createTextFinder(rows[i][0]).matchCase(false).findNext();
          if (!exists) newRows.push(rows[i]);
        }

        if (newRows.length > 0) {
          sheetLinks.insertRowsBefore(2, newRows.length);
          sheetLinks.getRange(2,1,newRows.length,5).setValues(newRows);
        }

        /*******************
         * 4.7 Отметка и лог
         *******************/
        thread.addLabel(label);
        if (newRows.length > 0) {
          logSheet.insertRowBefore(2);
          logSheet.getRange(2,1,1,6).setValues([[nowStr,source,filename,newRows.length,'✅ received','']]);
        }
      }
    }
  }
}