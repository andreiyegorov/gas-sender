function doPost(e) { 
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1MwuY9TvqVSlBVMCYO2LZfuLeksIO9rOTZwZM3lfT4qs/edit');
  const sh = ss.getSheetByName('REG');

  const payload = JSON.parse(e.postData.contents);
  const items = payload.items || [];

  items.forEach(item => {
    const ts = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd.MM.yy HH:mm");

    sh.appendRow([
      ts,                          // TIMESTAMP
      item.userId,                 // Login
      item.newPassword,            // NewPassword
      "https://www.assessmentlink.com/CoreParticipant/Participant/Login.aspx?lid=ru&cbid=&tkn=",
      "✅ registered"              // STATUS
    ]);
  });

  return ContentService.createTextOutput("OK")
      .setMimeType(ContentService.MimeType.TEXT);
}