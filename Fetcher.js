/*******************************************************
 *  📨 SENDER FETCHER — ловец CSV-приглашений
 *  📅 Версия: 2711-0500 (27 ноября 05:00)
 *  📘 Назначение:
 *     Извлекает CSV-вложения из писем с приглашениями Hogan.
 *     Добавляет новые строки в лист "Links" и ведёт журнал изменений.
 *  🔧 Изменения:
 *     • Формат даты: dd.MM  HH:mm (без года, два пробела, двоеточие)
 *******************************************************/

/*******************************************************
 *  1️⃣ КОНСТАНТЫ
 *******************************************************/
var FETCHER_LABEL_NAME = 'sender_fetcher_newlinks';
var LINKS_SHEET_NAME   = 'Links';
var LOG_SHEET_NAME     = 'received-log';

/*******************************************************
 *  2️⃣ ЗАПУСК FETCHER
 *******************************************************/
function senderInviteCheck() {
  senderFetcher_('menu-check');
}

/*******************************************************
 *  3️⃣ ОСНОВНАЯ ФУНКЦИЯ
 *******************************************************/
function senderFetcher_(source) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = Session.getScriptTimeZone();
  const now = new Date();
  const nowStr = Utilities.formatDate(now, tz, 'dd.MM  HH:mm');
  
  let totalAdded = 0; // счётчик добавленных строк для Dashboard

  /*******************
   * 3.1 Гарантируем наличие листов
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
   * 3.2 Поиск писем (без меток — проверяем всё)
   *******************/
  const threads = GmailApp.search('in:inbox newer_than:30d has:attachment');

  /*******************
   * 3.3 Обработка писем (старые → новые)
   *******************/
  for (let t = threads.length - 1; t >= 0; t--) {
    const thread = threads[t];
    const msgs = thread.getMessages();

    for (let m = 0; m < msgs.length; m++) {
      const msg = msgs[m];
      const dateStr = Utilities.formatDate(msg.getDate(), tz, 'dd.MM  HH:mm');
      const atts = msg.getAttachments();
      if (!atts || atts.length === 0) continue;

      /*******************
       * 3.4 Обработка CSV-вложений
       *******************/
      for (let a = 0; a < atts.length; a++) {
        const att = atts[a];
        const filename = att.getName();
        if (!att.getContentType().match(/csv/i) && !filename.match(/\.csv$/i)) continue;

        let csv;
        try { csv = Utilities.parseCsv(att.getDataAsString()); } catch(e) { continue; }
        if (!csv || csv.length < 2) continue;

        /*******************
         * 3.5 Подготовка строк — сортировка ID по убыванию
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
         * 3.6 Запись в Links (новые сверху)
         *******************/
        let newRows = [];
        for (let i = 0; i < rows.length; i++) {
          const exists = sheetLinks.createTextFinder(rows[i][0]).matchCase(false).findNext();
          if (!exists) newRows.push(rows[i]);
        }

        if (newRows.length > 0) {
          sheetLinks.insertRowsBefore(2, newRows.length);
          sheetLinks.getRange(2,1,newRows.length,5).setValues(newRows);
          totalAdded += newRows.length;
        }

        /*******************
         * 3.7 Лог (без меток)
         *******************/
        if (newRows.length > 0) {
          logSheet.insertRowBefore(2);
          logSheet.getRange(2,1,1,6).setValues([[nowStr,source,filename,newRows.length,'✅ received','']]);
        }
      }
    }
  }
  
  // Обновить Dashboard
  updateDashboardFetcherCheck_(source, now);
  if (totalAdded > 0) {
    updateDashboardLastFetched_(totalAdded, now);
  }
}

