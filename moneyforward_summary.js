// ==================================================
// マネーフォワード カード引落まとめ for Google Apps Script
// 使い方: setupTrigger を1回実行するだけ！
// ==================================================

// 新着メールチェック（6時間ごと自動実行）
function checkNewEmails() {
  var props = PropertiesService.getScriptProperties();
  var saved = props.getProperty('PROCESSED_IDS');
  var processedArr = saved ? JSON.parse(saved) : [];
  var processedIds = {};
  for (var i = 0; i < processedArr.length; i++) {
    processedIds[processedArr[i]] = true;
  }

  var query = 'from:feedback@moneyforward.com subject:' + encodeURIComponent('引き落としのお知らせ') + ' newer_than:35d';
  query = 'from:feedback@moneyforward.com subject:引き落としのお知らせ newer_than:35d';
  var threads = GmailApp.search(query);
  var newRows = [];
  var newIds = [];

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      if (processedIds[msgId]) continue;

      var subject = msg.getSubject();
      var body = msg.getBody();
      var month = extractMonth(subject);
      var cardName = extractCardName(subject);
      var billingDate = extractBillingDate(body);
      var amount = extractAmount(body);

      newRows.push([month, cardName, billingDate, amount, '三井住友銀行', new Date()]);
      newIds.push(msgId);
      processedIds[msgId] = true;
    }
  }

  if (newRows.length === 0) {
    Logger.log('新着なし: ' + new Date());
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  appendToSheet(ss, newRows);

  var allIds = processedArr.concat(newIds);
  props.setProperty('PROCESSED_IDS', JSON.stringify(allIds));
  Logger.log('新着 ' + newRows.length + ' 件追加');
}

// 全件リフレッシュ（手動実行 または 毎月1日）
function summarizeMoneyForward() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_IDS');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('引落まとめ');
  if (!sheet) {
    sheet = ss.insertSheet('引落まとめ');
  }
  sheet.clearContents();
  setSheetHeader(sheet);

  var query = 'from:feedback@moneyforward.com subject:引き落としのお知らせ newer_than:90d';
  var threads = GmailApp.search(query);
  var rows = [];
  var processedIds = {};
  var idArr = [];

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      if (processedIds[msgId]) continue;
      processedIds[msgId] = true;
      idArr.push(msgId);

      var subject = msg.getSubject();
      var body = msg.getBody();
      rows.push([
        extractMonth(subject),
        extractCardName(subject),
        extractBillingDate(body),
        extractAmount(body),
        '三井住友銀行',
        new Date()
      ]);
    }
  }

  if (rows.length === 0) {
    sheet.getRange(2, 1).setValue('データなし（直近90日のメールが見つかりませんでした）');
    return;
  }

  rows.sort(function(a, b) {
    if (a[0] !== b[0]) return b[0] > a[0] ? 1 : -1;
    return a[1] > b[1] ? 1 : -1;
  });

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('¥#,##0');
  sheet.autoResizeColumns(1, 6);

  PropertiesService.getScriptProperties().setProperty('PROCESSED_IDS', JSON.stringify(idArr));

  createMonthlySummary(ss, rows);
  SpreadsheetApp.getUi().alert('完了！\n月別合計シートの黄色セルに残高を入力してください。');
}

// シートへの差分追記
function appendToSheet(ss, newRows) {
  var sheet = ss.getSheetByName('引落まとめ');
  if (!sheet) {
    sheet = ss.insertSheet('引落まとめ');
    setSheetHeader(sheet);
  }
  if (sheet.getLastRow() < 1) {
    setSheetHeader(sheet);
  }

  var existingKeys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < existing.length; i++) {
      existingKeys[existing[i][0] + '|' + existing[i][1]] = true;
    }
  }

  var toAdd = [];
  for (var j = 0; j < newRows.length; j++) {
    var key = newRows[j][0] + '|' + newRows[j][1];
    if (!existingKeys[key]) toAdd.push(newRows[j]);
  }
  if (toAdd.length === 0) return;

  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, toAdd.length, 6).setValues(toAdd);
  sheet.getRange(start, 4, toAdd.length, 1).setNumberFormat('¥#,##0');
  sheet.autoResizeColumns(1, 6);

  var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var filtered = [];
  for (var k = 0; k < allData.length; k++) {
    if (allData[k][0] !== '') filtered.push(allData[k]);
  }
  createMonthlySummary(ss, filtered);
}

// ヘッダー設定
function setSheetHeader(sheet) {
  sheet.getRange(1, 1, 1, 6).setValues([['年月', 'カード名', '引落予定日', '引落予定額', '口座', '取得日時']]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  sheet.getRange(1, 1, 1, 6).setBackground('#4a86e8');
  sheet.getRange(1, 1, 1, 6).setFontColor('#ffffff');
}

// 件名から年月を抽出
function extractMonth(subject) {
  var m = subject.match(/（(\d{4}年\d{2}月)）/);
  return m ? m[1] : '不明';
}

// 件名からカード名を抽出
function extractCardName(subject) {
  var m = subject.match(/】(.+?) 引き落とし/);
  return m ? m[1].trim() : '不明';
}

// 引落予定日を抽出
function extractBillingDate(html) {
  var text = html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
  var p1 = text.match(/引き落とし予定日\s*(\d{4}年\d{2}月\d{2}日)/);
  if (p1) return p1[1];
  var p2 = text.match(/引落[^\d]*(\d{4}\/\d{2}\/\d{2})/);
  if (p2) return p2[1];
  var p3 = text.match(/引き落とし予定日[^\d]*(\d{2}月\d{2}日)/);
  if (p3) return p3[1];
  return '不明';
}

// 引落予定額を抽出
function extractAmount(html) {
  var text = html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
  var p1 = text.match(/引き落とし予定額\s*[¥￥]?([\d,]+)\s*円?/);
  if (p1) return parseInt(p1[1].replace(/,/g, ''), 10);
  var p2 = text.match(/[¥￥]([\d,]+)/);
  if (p2) return parseInt(p2[1].replace(/,/g, ''), 10);
  var p3 = text.match(/([\d,]+)円/);
  if (p3) return parseInt(p3[1].replace(/,/g, ''), 10);
  return 0;
}

// 月別合計シート作成（残高チェックつき）
function createMonthlySummary(ss, rows) {
  var summarySheet = ss.getSheetByName('月別合計');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('月別合計');
  }

  // 既存の残高を保持
  var savedBalances = {};
  var lastRow = summarySheet.getLastRow();
  if (lastRow > 1) {
    var existing = summarySheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i][1] === '三井住友銀行 残高（手入力）' && existing[i][2]) {
        savedBalances[existing[i][0]] = existing[i][2];
      }
    }
  }

  summarySheet.clearContents();
  summarySheet.clearFormats();
  summarySheet.getRange(1, 1, 1, 5).setValues([['年月', '項目', '金額', '', '判定']]);
  summarySheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  summarySheet.getRange(1, 1, 1, 5).setBackground('#34a853');
  summarySheet.getRange(1, 1, 1, 5).setFontColor('#ffffff');

  // 月別集計
  var monthMap = {};
  var monthOrder = [];
  for (var r = 0; r < rows.length; r++) {
    var mn = rows[r][0];
    var amt = typeof rows[r][3] === 'number' ? rows[r][3] : 0;
    if (!monthMap[mn]) {
      monthMap[mn] = { cards: [], total: 0 };
      monthOrder.push(mn);
    }
    monthMap[mn].cards.push({ name: rows[r][1], date: rows[r][2], amount: rows[r][3] });
    monthMap[mn].total += amt;
  }
  monthOrder.sort(function(a, b) { return b > a ? 1 : -1; });

  var rowNum = 2;

  for (var mi = 0; mi < monthOrder.length; mi++) {
    var month = monthOrder[mi];
    var data = monthMap[month];

    // カード別明細
    for (var ci = 0; ci < data.cards.length; ci++) {
      var card = data.cards[ci];
      summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, card.name, card.amount]]);
      summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
      rowNum++;
    }

    // 引落合計
    summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, '【引落合計】', data.total]]);
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#e8f5e9');
    summarySheet.getRange(rowNum, 1, 1, 5).setFontWeight('bold');
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    var totalRow = rowNum;
    rowNum++;

    // 残高入力欄（黄色）
    var bal = savedBalances[month] || '';
    summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, '三井住友銀行 残高（手入力）', bal]]);
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#fff9c4');
    summarySheet.getRange(rowNum, 2).setFontWeight('bold');
    summarySheet.getRange(rowNum, 3).setNote('三井住友銀行の残高をここに入力してください');
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    var balanceRow = rowNum;
    rowNum++;

    // 差引残高
    summarySheet.getRange(rowNum, 1).setValue(month);
    summarySheet.getRange(rowNum, 2).setValue('差引残高（引落後）');
    summarySheet.getRange(rowNum, 3).setFormula(
      '=IF(C' + balanceRow + '="","残高を入力",C' + balanceRow + '-C' + totalRow + ')'
    );
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    summarySheet.getRange(rowNum, 5).setFormula(
      '=IF(C' + balanceRow + '="","－",IF(C' + rowNum + '>=0,"OK 足りる！","NG 不足！"))'
    );
    summarySheet.getRange(rowNum, 5).setFontWeight('bold');
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#e3f2fd');

    // 条件付き書式
    var cfRules = summarySheet.getConditionalFormatRules();
    cfRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#ffcdd2')
        .setFontColor('#c62828')
        .setRanges([summarySheet.getRange(rowNum, 1, 1, 5)])
        .build()
    );
    cfRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0)
        .setBackground('#c8e6c9')
        .setFontColor('#1b5e20')
        .setRanges([summarySheet.getRange(rowNum, 3, 1, 1)])
        .build()
    );
    summarySheet.setConditionalFormatRules(cfRules);

    rowNum += 2;
  }

  summarySheet.autoResizeColumns(1, 5);
}

// トリガー設定（初回に1回だけ実行）
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // 6時間ごとに新着チェック
  ScriptApp.newTrigger('checkNewEmails')
    .timeBased()
    .everyHours(6)
    .create();

  // 毎月1日9時に全件リフレッシュ
  ScriptApp.newTrigger('summarizeMoneyForward')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  // 毎月1日9時にサマリーメール
  ScriptApp.newTrigger('sendMonthlySummaryEmail')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  SpreadsheetApp.getUi().alert(
    'トリガー設定完了！\n6時間ごとに新着チェック＆自動更新されます。\n次は summarizeMoneyForward を実行して初回データを取得してください。'
  );
}

// サマリーメール送信
function sendMonthlySummaryEmail() {
  var query = 'from:feedback@moneyforward.com subject:引き落としのお知らせ newer_than:35d';
  var threads = GmailApp.search(query);
  var now = new Date();
  var currentYM = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月');
  var cardData = [];
  var seen = {};

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      if (seen[msgId]) continue;
      seen[msgId] = true;

      var subject = msg.getSubject();
      var month = extractMonth(subject);
      if (month.indexOf(currentYM.slice(0, 7)) < 0) continue;

      var body = msg.getBody();
      cardData.push({
        cardName: extractCardName(subject),
        billingDate: extractBillingDate(body),
        amount: extractAmount(body)
      });
    }
  }

  if (cardData.length === 0) return;

  var total = 0;
  var bodyText = '【' + currentYM + ' カード引落まとめ】\n口座：三井住友銀行\n-----------------\n';
  for (var i = 0; i < cardData.length; i++) {
    var d = cardData[i];
    var amtStr = d.amount > 0 ? '\\' + d.amount.toLocaleString() : '要確認';
    bodyText += d.cardName + '\n  引落日：' + d.billingDate + '\n  金額：' + amtStr + '\n\n';
    total += d.amount;
  }
  bodyText += '-----------------\n合計：\\' + total.toLocaleString() + '\n';

  var userEmail = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(userEmail, '【' + currentYM + '】カード引落まとめ', bodyText);
  Logger.log('サマリーメール送信: ' + userEmail);
}
