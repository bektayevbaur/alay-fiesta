/**
 * ALAY Fiesta — приём заявок с сайта в Google Таблицу.
 *
 * Установка (подробно в README.md):
 * 1. Создайте пустую Google Таблицу → Расширения → Apps Script.
 * 2. Вставьте этот код вместо содержимого Code.gs и сохраните.
 * 3. Выберите функцию setup и нажмите «Выполнить» (разрешите доступ).
 * 4. «Начать развертывание» → «Новое развертывание» → тип «Веб-приложение»,
 *    «Запуск от имени»: Я, «У кого есть доступ»: Все → «Развернуть».
 * 5. Скопируйте URL веб-приложения в config.js на сайте.
 *
 * Наружу (doGet) отдаются только счётчики, без имён и телефонов.
 */

var SHEET_NAME = 'Заявки';
var PRICE = 25000;
var RECEIPTS_FOLDER = 'ALAY Fiesta — чеки';

var HEADERS = [
  'Дата заявки', 'Имя и фамилия', 'Телефон', 'Гостей', 'Имена гостей',
  'Образ', 'Оплата (со слов гостя)', 'На чьё имя перевод', 'Питание',
  'Комментарий', 'Чек', 'Человек', 'Сумма, ₸', 'Оплатил ✓', 'Заметка организатора'
];
// Колонки (1-based)
var COL = { heads: 12, sum: 13, paid: 14 };

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#cc2f1f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, HEADERS.length, 150);
  sh.setColumnWidth(2, 200);
  var rows = 600;
  sh.getRange(2, COL.paid, rows, 1).insertCheckboxes();
  sh.getRange(2, COL.sum, rows, 1).setNumberFormat('#,##0');
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$N2=TRUE')
    .setBackground('#dff3e7')
    .setRanges([sh.getRange(2, 1, rows, HEADERS.length)])
    .build();
  sh.setConditionalFormatRules([rule]);
  var first = ss.getSheets()[0];
  if (first.getName() !== SHEET_NAME && first.getLastRow() === 0) ss.deleteSheet(first);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true }); // ловушка для ботов
    var name = clean_(d.name, 80), phone = clean_(d.phone, 30);
    if (!name || !phone) return json_({ ok: false, error: 'Укажите имя и телефон' });
    var guests = Math.max(0, Math.min(3, parseInt(d.guests, 10) || 0));
    var heads = 1 + guests;

    var receipt = '';
    if (d.receipt && d.receipt.data && String(d.receipt.data).length < 4000000) {
      try {
        var bytes = Utilities.base64Decode(String(d.receipt.data));
        var type = /^image\/(jpeg|png|webp)$|^application\/pdf$/.test(d.receipt.type) ? d.receipt.type : 'image/jpeg';
        var ext = type === 'application/pdf' ? '.pdf' : '.jpg';
        var blob = Utilities.newBlob(bytes, type, 'Чек — ' + name + ext);
        receipt = folder_().createFile(blob).getUrl();
      } catch (err) { receipt = 'не удалось сохранить файл'; }
    }

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var row = [
      new Date(), txt_(name), txt_(phone), guests, txt_(clean_(d.guestNames, 200)),
      txt_(clean_(d.look, 40)), txt_(clean_(d.payment, 40)), txt_(clean_(d.payer, 80)),
      txt_(clean_(d.diet, 120)), txt_(clean_(d.comment, 500)), receipt, heads, heads * PRICE, false, ''
    ];
    var r = nextRow_(sh);
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    sh.getRange(r, COL.paid).insertCheckboxes();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'Не удалось сохранить заявку' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var res = { ok: true, records: 0, heads: 0, paidHeads: 0, paidRecords: 0, updated: new Date().toISOString() };
  var last = sh ? sh.getLastRow() : 1;
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, COL.paid).getValues();
    vals.forEach(function (v) {
      if (!v[1]) return;
      var h = Number(v[COL.heads - 1]) || 1;
      res.records++; res.heads += h;
      if (v[COL.paid - 1] === true) { res.paidRecords++; res.paidHeads += h; }
    });
  }
  return json_(res);
}

// Первая пустая строка по колонке «Имя». getLastRow() не подходит:
// заранее поставленные галочки считаются заполненными ячейками.
function nextRow_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 2;
  var names = sh.getRange(2, 2, last - 1, 1).getValues();
  for (var i = names.length - 1; i >= 0; i--) {
    if (String(names[i][0]).trim() !== '') return i + 3;
  }
  return 2;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clean_(s, max) {
  return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}
// Защита от формул: всё текстовое пишем как текст.
function txt_(s) { return s ? "'" + s : ''; }
function folder_() {
  var it = DriveApp.getFoldersByName(RECEIPTS_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPTS_FOLDER);
}
