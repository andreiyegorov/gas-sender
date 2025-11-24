/*******************************************************
 *  💰 SBER ARRIVALS PARSER
 *  📅 Версия: 2611-0935 (26 ноября 09:35)
 *  📘 Назначение:
 *     Извлекает только ПРИХОДЫ ("Вам поступили средства")
 *     из писем СберБизнес и записывает в лист "SB_arrivals"
 *     (новые строки добавляются сверху).
 *******************************************************/

var SB_ARRIVALS_SHEET = 'SB_arrivals';

function parseSberArrivals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SB_ARRIVALS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SB_ARRIVALS_SHEET);
    sh.appendRow([
      'Дата проверки',
      'Дата письма',
      'Компания',
      'ИНН',
      'Р/С',
      '№ договора',
      'Дата договора',
      'Сумма',
      'Назначение'
    ]);
  }

  const threads = GmailApp.search('subject:(операции по счёту) "поступили средства" newer_than:30d');
  const tz = Session.getScriptTimeZone();
  const checkDate = Utilities.formatDate(new Date(), tz, 'dd.MM.yy HH:mm');

  for (let t = threads.length - 1; t >= 0; t--) {
    const thread = threads[t];
    const msgs = thread.getMessages();

    for (let m = 0; m < msgs.length; m++) {
      const msg = msgs[m];
      const body = msg.getPlainBody();
      if (!body || !body.includes('поступили средства')) continue;

      const dateStr = Utilities.formatDate(msg.getDate(), tz, 'dd.MM.yy HH:mm');

      const senderMatch = body.match(/Кто отправитель\?\s*([\s\S]*?)Назначение платежа:/i);
      const senderBlock = senderMatch ? senderMatch[1].trim().replace(/\n/g, ' ') : '';
      const company = (senderBlock.match(/^([^,]+)/) || [''])[0].trim();
      const inn = (senderBlock.match(/ИНН\s*([*\d]+)/) || [''])[1];
      const rs = (senderBlock.match(/р\/с\s*([*\d]+)/i) || [''])[1];

      const contractMatch = body.match(/договор[ау]*\s*№\s*([\d\-]+)\s*от\s*([\d\.]+)/i);
      const contractNum = contractMatch ? contractMatch[1] : '';
      const contractDate = contractMatch ? contractMatch[2] : '';

      const amountMatch = body.match(/\+?\s*([\d\s]+(?:,\d{2})?)\s*RUB/i);
      const amount = amountMatch ? parseFloat(amountMatch[1].replace(/\s+/g, '').replace(',', '.')) : '';

      const purposeMatch = body.match(/Назначение платежа:\s*([\s\S]*?)(?:Пожалуйста|Россия|©|$)/i);
      let purpose = purposeMatch ? purposeMatch[1].trim().replace(/\n/g, ' ') : '';
      purpose = purpose
        .replace(/Без налога\s*\(НДС\)/gi, '')
        .replace(/БЕЗ НАЛОГА\s*\(НДС\)/gi, '')
        .replace(/В\.?\s*Т\.?\s*Ч\.?\s*НДС\s*0%[^.,;]*/gi, '')
        .replace(/НДС не облагается/gi, '')
        .replace(/НДС\s*0%[^.,;]*/gi, '')
        .replace(/[>]+/g, '')
        .replace(/\s{2,}/g, ' ')
        .trim();

      sh.insertRowBefore(2);
      sh.getRange(2, 1, 1, 9).setValues([[
        checkDate, dateStr, company, inn, rs, contractNum, contractDate, amount, purpose
      ]]);
    }
  }
}