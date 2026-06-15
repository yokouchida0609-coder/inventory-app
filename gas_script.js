// Google Apps Script - 全アプリ共通データ保存スクリプト
// このコードをGoogleスプレッドシートのApps Scriptに貼り付けて、
// Webアプリとしてデプロイしてください。

function doGet(e) {
  const sheetName = e.parameter.sheet || 'inventory';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  const data = sheet.getRange(1, 1).getValue();
  const output = ContentService.createTextOutput(data || '[]');
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const sheetName = payload.sheet || 'inventory';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1).setValue(JSON.stringify(payload.items));
  const output = ContentService.createTextOutput('ok');
  output.setMimeType(ContentService.MimeType.TEXT);
  return output;
}
