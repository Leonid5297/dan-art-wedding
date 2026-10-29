/**
 * Приём анкет гостей в Google Таблицу.
 * Колонки: Дата | ФИО | Количество человек | Телефон
 *
 * Обновление (адрес /exec остаётся прежним):
 * 1. Таблица → Расширения → Apps Script → замените содержимое Code.gs этим кодом, сохраните.
 * 2. Развернуть → Управление развёртываниями → значок карандаша → Версия: «Новая версия» → Развернуть.
 */

var TZ = 'Asia/Tomsk';
var HEADERS = ['Дата', 'ФИО', 'Количество человек', 'Телефон'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return reply_({ ok: false, error: 'busy' });

  try {
    var p = e.parameter || {};
    var fio = clean_(p.fio || p.name, 60);
    var phone = clean_(p.phone, 18);
    var count = parseInt(p.count || p.guests, 10);
    if (isNaN(count) || count < 0) count = 0;
    if (count > 10) count = 10;

    if (fio.length < 2) return reply_({ ok: false, error: 'empty fio' });

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var row = sheet.getLastRow() + 1;
    // ФИО и телефон пишем как текст, чтобы «+7 (999)…» не превратилось в формулу
    sheet.getRange(row, 2).setNumberFormat('@');
    sheet.getRange(row, 4).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 4).setValues([[
      Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'),
      fio,
      count,
      phone
    ]]);

    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Открыв адрес /exec в браузере, можно убедиться, что скрипт жив
function doGet() {
  return reply_({ ok: true, status: 'alive' });
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
