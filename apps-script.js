// =====================================================
// Вставь этот код в Google Apps Script
// Инструкция: см. README или сообщение от Claude
// =====================================================

const SHEET_NAME = 'Дневник';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Дата-время', 'Имя', 'Команда', 'Заметка']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET — загрузить записи участника
function doGet(e) {
  const name = (e.parameter.name || '').trim();
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues().slice(1); // skip header

  const entries = rows
    .filter(r => r[1] === name)
    .map(r => ({
      datetime: r[0] ? new Date(r[0]).toISOString() : '',
      name:     r[1],
      team:     r[2],
      text:     r[3]
    }))
    .reverse(); // newest first

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, entries }))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST — сохранить новую запись
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    sheet.appendRow([
      new Date(),
      (data.name || '').trim(),
      (data.team || '').trim(),
      (data.text || '').trim()
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
