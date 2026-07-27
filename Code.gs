/**
 * Бэкенд для сайта отчёта по текущему ремонту.
 * Хранит данные в листе "Storage" этой же Google Таблицы (ключ / значение).
 *
 * НАСТРОЙКА:
 * 1. Создайте новую Google Таблицу (sheets.google.com).
 * 2. Расширения -> Apps Script.
 * 3. Удалите весь код-заглушку и вставьте сюда содержимое этого файла.
 * 4. Сохраните проект (значок дискеты).
 * 5. Развернуть -> Новое развёртывание.
 *    - Тип: Веб-приложение
 *    - Выполнять от имени: Меня
 *    - У кого есть доступ: Все (Anyone)
 * 6. Разрешите доступ (Google покажет предупреждение "непроверенное приложение" —
 *    это нормально для собственного скрипта: "Дополнительно" -> "Перейти на страницу (небезопасно)" -> "Разрешить").
 * 7. Скопируйте полученную ссылку (заканчивается на /exec).
 * 8. Вставьте эту ссылку в файл сайта (otchet_tr_2026.html) в переменную API_URL.
 */

function doGet(e) {
  var action = e.parameter.action;
  var sheet = getSheet_();

  if (action === 'get') {
    var key = e.parameter.key;
    var row = findRow_(sheet, key);
    return jsonResponse_({ key: key, value: row ? row[1] : null });
  }

  if (action === 'bulkGet') {
    var prefix = e.parameter.prefix || '';
    var data = sheet.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < data.length; i++) {
      var k = data[i][0];
      if (k && String(k).indexOf(prefix) === 0) {
        result[k] = data[i][1];
      }
    }
    return jsonResponse_({ data: result });
  }

  return jsonResponse_({ error: 'unknown action' });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var sheet = getSheet_();

  if (body.action === 'set') {
    upsert_(sheet, body.key, body.value);
    return jsonResponse_({ ok: true });
  }

  if (body.action === 'delete') {
    deleteKey_(sheet, body.key);
    return jsonResponse_({ ok: true });
  }

  return jsonResponse_({ error: 'unknown action' });
}

// ---- вспомогательные функции ----

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Storage');
  if (!sheet) {
    sheet = ss.insertSheet('Storage');
    sheet.appendRow(['key', 'value', 'updatedAt']);
  }
  return sheet;
}

function findRow_(sheet, key) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i];
  }
  return null;
}

function findRowIndex_(sheet, key) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return i + 1; // номер строки (1-индексация)
  }
  return -1;
}

function upsert_(sheet, key, value) {
  var idx = findRowIndex_(sheet, key);
  var now = new Date().toISOString();
  if (idx === -1) {
    sheet.appendRow([key, value, now]);
  } else {
    sheet.getRange(idx, 2).setValue(value);
    sheet.getRange(idx, 3).setValue(now);
  }
}

function deleteKey_(sheet, key) {
  var idx = findRowIndex_(sheet, key);
  if (idx !== -1) sheet.deleteRow(idx);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
