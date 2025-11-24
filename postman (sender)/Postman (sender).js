/*******************************************************
 *  📨 SENDER POSTMAN — REG VERSION (костыль)
 *  Обновлено: 2025-11
 *******************************************************/

var SHEET_ID   = SpreadsheetApp.getActive().getId();
var REG_SHEET  = 'REG';
var ID_COL     = 2;   // колонка B
var EMAIL_COL  = 4;   // колонка D
var LABEL_NAME = 'sender_postman_done';
var BCC_EMAIL  = 'goldensequence@proton.me';

/*******************
 * 2. ЗАПУСК
 *******************/
function senderPostmanRun() {
  updateSenderManualStatus_('senderPostmanRun', new Date());
  senderPostman_('menu');
}

function senderPostmanCheck() {
  updateSenderManualStatus_('senderPostmanCheck', new Date());
  senderPostman_('menu-check');
}

function senderPostmanTrigger() {
  senderPostman_('trigger');
}

/*******************
 * 3. ОСНОВНАЯ ФУНКЦИЯ
 *******************/
function senderPostman_(source) {

  globalThis.__idsSent = []; // список отправленных ID

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reg = ss.getSheetByName(REG_SHEET).getDataRange().getValues();
  var logSheet = ss.getSheetByName('sender-log');

  if (!logSheet) {
    logSheet = ss.insertSheet('sender-log');
    logSheet.appendRow([
      'Timestamp','Source','ID','Sent to','Original file','Renamed file','Status','Error'
    ]);
  }

  var label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);

  var searchQuery = [
    'in:inbox','newer_than:7d','(',
      'from:@hoganassessments.com',
      'OR from:@hoganassessments.co',
      'OR from:@hoganassessments.eu',
      'OR from:(Hogan Assessment)',
      'OR subject:(Hogan Report)',
    ')'
  ].join(' ');

  var threads = GmailApp.search(searchQuery);

  var tz = Session.getScriptTimeZone();
  var nowDate = new Date();
  var nowStr = Utilities.formatDate(nowDate, tz, 'dd.MM.yyyy HH:mm');

  for (var t = 0; t < threads.length; t++) {

    var thread = threads[t];
    if (thread.getLabels().some(function(l){return l.getName()===LABEL_NAME;})) continue;

    var msgs = thread.getMessages();
    for (var m = 0; m < msgs.length; m++) {

      var msg = msgs[m];
      var body = msg.getPlainBody() || '';
      var html = msg.getBody() || '';
      var subject = msg.getSubject() || '';
      var attachments = msg.getAttachments();

      if (!attachments || attachments.length === 0) continue;

      var idMatch = (body.match(/HL\d{6}/i) || html.match(/HL\d{6}/i) || subject.match(/HL\d{6}/i));
      if (!idMatch) continue;

      var id = idMatch[0].toUpperCase();

      // поиск ID в REG
      var row = reg.findIndex(function(r){ return (r[ID_COL-1]||'').toString().trim() === id; });
      if (row === -1) continue;

      var email = (reg[row][EMAIL_COL-1] || '').toString().trim();
      if (!email) email = 'yegorov@me.com';

      try {
        var renamed = renameSenderFiles_(attachments, id);

        globalThis.__idsSent.push(id);

        GmailApp.sendEmail(
          email,
          'Hogan Report: ' + id,
          'Здравствуйте! Ваш отчёт Hogan готов. Откройте вложение.',
          {
            attachments: renamed,
            htmlBody: '<p>Здравствуйте!<br>Ваш отчёт Hogan готов.<br>Откройте вложение.</p>',
            bcc: BCC_EMAIL,
            name: 'Hogan Sender Postman'
          }
        );

        // отметить в REG, колонка E
        ss.getSheetByName(REG_SHEET).getRange(row + 1, 5).setValue('✅ sent');

        // метка против повторов
        thread.addLabel(label);

        // запись в лог
        logSheet.insertRowBefore(2);
        logSheet.getRange(2,1,1,8).setValues([[
          nowStr,
          source,
          id,
          email,
          attachments[0].getName(),
          renamed[0].getName(),
          '✅ sent',
          ''
        ]]);

      } catch (err) {

        logSheet.insertRowBefore(2);
        logSheet.getRange(2,1,1,8).setValues([[
          nowStr,
          source,
          id,
          email,
          attachments[0].getName(),
          '',
          '⚠️ error',
          String(err)
        ]]);
      }
    }
  }

  // обновить панель статусов (автомат)
updateSenderAutoStatusPanel_(globalThis.__idsSent, nowDate);
globalThis.__idsSent = [];
}

/*******************
 * 4. ПЕРЕИМЕНОВАНИЕ
 *******************/
function renameSenderFiles_(atts, id) {
  var out = [];
  var base = id + ' Report.pdf';

  for (var i = 0; i < atts.length; i++) {
    var n = base;
    if (atts.length > 1) {
      var dot = base.lastIndexOf('.');
      n = base.slice(0, dot) + '-' + (i+1) + base.slice(dot);
    }
    out.push(atts[i].copyBlob().setName(n));
  }
  return out;
}