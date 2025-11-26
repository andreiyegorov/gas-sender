/*******************************************************
 *  🔍 MINI PROJECT INSPECTOR
 *  📅 Версия: 2711-0505 (27 ноября 05:05)
 *  📘 Назначение:
 *     Инспекция функций в текущем Apps Script проекте
 *******************************************************/

function inspectProjectMini() {

  const SYSTEM_HANDLERS = [
    'onOpen','onEdit','onInstall','doGet','doPost','onSelectionChange'
  ];

  const ui = SpreadsheetApp.getUi();
  let report = [];

  /* ───────── СБОР ФУНКЦИЙ ИЗ ГЛОБАЛЬНОГО ОБЪЕКТА ───────── */
  const globalNames = Object.keys(this).filter(n =>
    typeof this[n] === 'function' &&
    !n.startsWith('_') &&
    n !== 'inspectProjectMini'
  );

  report.push('=== FUNCTIONS IN GLOBAL SCOPE ===');
  globalNames.forEach(n => report.push('• ' + n));
  report.push('');

  /* ───────── ХЕНДЛЕРЫ ───────── */
  report.push('=== SYSTEM HANDLERS ===');
  SYSTEM_HANDLERS.forEach(h => {
    if (globalNames.includes(h)) {
      report.push('⚙️ ' + h);
    }
  });
  report.push('');

  /* ───────── ТРИГГЕРЫ ───────── */
  report.push('=== TRIGGERS ===');
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    report.push('(none)');
  } else {
    triggers.forEach(t => {
      report.push(`• ${t.getHandlerFunction()} — ${t.getEventType()}`);
    });
  }
  report.push('');

  /* ───────── НЕПРИВЯЗАННЫЕ ФУНКЦИИ ───────── */
  report.push('=== POTENTIALLY UNUSED ===');
  let unused = [];
  globalNames.forEach(n => {
    let hasTrigger = triggers.some(t => t.getHandlerFunction() === n);
    let isHandler = SYSTEM_HANDLERS.includes(n);
    if (!hasTrigger && !isHandler) unused.push(n);
  });

  if (unused.length === 0) report.push('(none)');
  else unused.forEach(n => report.push('🟡 ' + n));

  report.push('');

  /* ───────── ВЫВОД ───────── */
  const result = report.join('\n');
  Logger.log(result);

  const html = `
<html>
<body>
<textarea id="out" style="width:100%;height:90%;font-family:monospace;">${result
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')}</textarea>
<script>
  const ta = document.getElementById('out');
  ta.select();
  document.execCommand('copy');
</script>
<div style="font-family:sans-serif;padding-top:6px;">
✅ Скопировано в буфер.
</div>
</body>
</html>
`;

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500),
    '🔍 MINI PROJECT INSPECTOR'
  );
}