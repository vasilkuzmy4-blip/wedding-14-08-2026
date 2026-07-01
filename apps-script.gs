/**
 * Google Apps Script для приёма ответов формы RSVP свадебного сайта.
 * Вставь этот код в Расширения → Apps Script своей Google-таблицы,
 * затем разверни как веб-приложение (см. RSVP-SETUP.md).
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    var data = {};
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = {};
    }

    var drinks = Array.isArray(data.drinks) ? data.drinks.join(', ') : (data.drinks || '');

    sheet.appendRow([
      new Date(),          // Дата
      data.name || '',     // Имя
      data.attend || '',   // Статус (Буду / Не смогу)
      drinks,              // Напитки
      data.stay || '',     // Размещение (Да / Нет / Не знаю)
      data.msg || '',      // Сообщение
      data.page || ''      // Страница
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Позволяет открыть URL /exec в браузере для быстрой проверки, что скрипт жив.
function doGet() {
  return ContentService.createTextOutput('RSVP endpoint is running.');
}
