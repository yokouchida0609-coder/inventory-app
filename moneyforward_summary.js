/**
 * マネーフォワード カード引落まとめスクリプト
 *
 * 【使い方】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 > Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. 「setupTrigger」を一度だけ実行（初回はGmail権限の許可が必要）
 *    → 以降は6時間ごとに自動チェック＆更新される！
 * 5. 手動で今すぐ更新したいときは「summarizeMoneyForward」を実行
 */

// ==============================
// 新着チェック（6時間ごとに自動実行）
// ==============================
function checkNewEmails() {
  const props = PropertiesService.getScriptProperties();
  const processedKey = 'PROCESSED_IDS';
  const processedIds = new Set(JSON.parse(props.getProperty(processedKey) || '[]'));

  // 直近35日のマネーフォワードメールを検索
  const query = 'from:feedback@moneyforward.com subject:引き落としのお知らせ newer_than:35d';
  const threads = GmailApp.search(query);

  const newRows = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const msgId = message.getId();
      if (processedIds.has(msgId)) return; // 処理済みはスキップ

      const subject = message.getSubject();
      const body = message.getBody();

      const monthMatch = subject.match(/（(\d{4}年\d{2}月)）/);
      const month = monthMatch ? monthMatch[1] : '不明';

      const cardMatch = subject.match(/】(.+?) 引き落とし/);
      const cardName = cardMatch ? cardMatch[1].trim() : '不明';

      const billingDate = extractBillingDate(body);
      const amount = extractAmount(body);

      newRows.push({ msgId, row: [month, cardName, billingDate, amount, '三井住友銀行', new Date()] });
      processedIds.add(msgId);
    });
  });

  if (newRows.length === 0) {
    Logger.log('新着なし：' + new Date());
    return;
  }

  // スプレッドシートに追記
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  appendToSheet(ss, newRows.map(n => n.row));

  // 処理済みIDを保存
  props.setProperty(processedKey, JSON.stringify([...processedIds]));

  Logger.log('新着 ' + newRows.length + '件を追加：' + newRows.map(n => n.row[1]).join(', '));
}

// ==============================
// メイン処理：全件リフレッシュ（手動用）
// ==============================
function summarizeMoneyForward() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 処理済みIDをリセット
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_IDS');

  let sheet = ss.getSheetByName('引落まとめ');
  if (!sheet) {
    sheet = ss.insertSheet('引落まとめ');
  }
  sheet.clearContents();
  setupSheetHeader(sheet);

  const query = 'from:feedback@moneyforward.com subject:引き落としのお知らせ newer_than:90d';
  const threads = GmailApp.search(query);

  const rows = [];
  const processedIds = new Set();

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const msgId = message.getId();
      if (processedIds.has(msgId)) return;
      processedIds.add(msgId);

      const subject = message.getSubject();
      const body = message.getBody();

      const monthMatch = subject.match(/（(\d{4}年\d{2}月)）/);
      const month = monthMatch ? monthMatch[1] : '不明';

      const cardMatch = subject.match(/】(.+?) 引き落とし/);
      const cardName = cardMatch ? cardMatch[1].trim() : '不明';

      const billingDate = extractBillingDate(body);
      const amount = extractAmount(body);

      rows.push([month, cardName, billingDate, amount, '三井住友銀行', new Date()]);
    });
  });

  if (rows.length === 0) {
    sheet.getRange(2, 1).setValue('データなし（直近90日のメールが見つかりませんでした）');
    return;
  }

  rows.sort((a, b) => {
    if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
    return a[1].localeCompare(b[1]);
  });

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('¥#,##0');
  sheet.autoResizeColumns(1, 6);

  // 処理済みIDを保存（次回以降の差分チェックに使う）
  PropertiesService.getScriptProperties().setProperty(
    'PROCESSED_IDS',
    JSON.stringify([...processedIds])
  );

  createMonthlySummary(ss, rows);

  SpreadsheetApp.getUi().alert('完了！引落まとめを更新しました。\n「月別合計」シートの黄色セルに残高を入力してください。');
  Logger.log('完了：' + rows.length + '件処理しました');
}

// ==============================
// シートへの差分追記
// ==============================
function appendToSheet(ss, newRows) {
  let sheet = ss.getSheetByName('引落まとめ');
  if (!sheet) {
    sheet = ss.insertSheet('引落まとめ');
    setupSheetHeader(sheet);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) setupSheetHeader(sheet);

  // 既存データを取得して重複確認（年月+カード名で判定）
  const existingData = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    : [];
  const existingKeys = new Set(existingData.map(r => r[0] + '|' + r[1]));

  const rowsToAdd = newRows.filter(r => !existingKeys.has(r[0] + '|' + r[1]));
  if (rowsToAdd.length === 0) return;

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
  sheet.getRange(startRow, 4, rowsToAdd.length, 1).setNumberFormat('¥#,##0');
  sheet.autoResizeColumns(1, 6);

  // 月別合計も再生成
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .filter(r => r[0] !== '');
  createMonthlySummary(ss, allData);
}

// ==============================
// シートのヘッダーを設定
// ==============================
function setupSheetHeader(sheet) {
  sheet.getRange(1, 1, 1, 6).setValues([
    ['年月', 'カード名', '引落予定日', '引落予定額', '口座', '取得日時']
  ]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  sheet.getRange(1, 1, 1, 6).setBackground('#4a86e8');
  sheet.getRange(1, 1, 1, 6).setFontColor('#ffffff');
}

// ==============================
// 引落予定日を抽出
// ==============================
function extractBillingDate(html) {
  const text = html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

  const p1 = text.match(/引き落とし予定日\s*(\d{4}年\d{2}月\d{2}日)/);
  if (p1) return p1[1];

  const p2 = text.match(/引落[^\d]*(\d{4}\/\d{2}\/\d{2})/);
  if (p2) return p2[1];

  const p3 = text.match(/引き落とし予定日[^\d]*(\d{2}月\d{2}日)/);
  if (p3) return p3[1];

  return '不明';
}

// ==============================
// 引落予定額を抽出
// ==============================
function extractAmount(html) {
  const text = html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

  const p1 = text.match(/引き落とし予定額\s*[¥￥]?([\d,]+)\s*円?/);
  if (p1) return parseInt(p1[1].replace(/,/g, ''), 10);

  const p2 = text.match(/[¥￥]([\d,]+)/);
  if (p2) return parseInt(p2[1].replace(/,/g, ''), 10);

  const p3 = text.match(/([\d,]+)円/);
  if (p3) return parseInt(p3[1].replace(/,/g, ''), 10);

  return '要確認';
}

// ==============================
// 月別合計シートを作成（残高チェックつき）
// ==============================
function createMonthlySummary(ss, rows) {
  let summarySheet = ss.getSheetByName('月別合計');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('月別合計');
  }

  // 既存の残高入力値を保持（更新時に消えないように）
  const savedBalances = {};
  const lastRow = summarySheet.getLastRow();
  if (lastRow > 1) {
    const existingData = summarySheet.getRange(2, 1, lastRow - 1, 5).getValues();
    existingData.forEach(row => {
      if (row[1] === '三井住友銀行 残高（手入力）' && row[2]) {
        savedBalances[row[0]] = row[2];
      }
    });
  }

  summarySheet.clearContents();
  summarySheet.clearFormats();

  summarySheet.getRange(1, 1, 1, 5).setValues([['年月', '項目', '金額', '', '判定']]);
  summarySheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  summarySheet.getRange(1, 1, 1, 5).setBackground('#34a853');
  summarySheet.getRange(1, 1, 1, 5).setFontColor('#ffffff');

  const monthMap = {};
  rows.forEach(row => {
    const month = row[0];
    const amount = typeof row[3] === 'number' ? row[3] : 0;
    if (!monthMap[month]) monthMap[month] = { cards: [], total: 0 };
    monthMap[month].cards.push({ name: row[1], date: row[2], amount: row[3] });
    monthMap[month].total += amount;
  });

  let rowNum = 2;
  const months = Object.keys(monthMap).sort((a, b) => b.localeCompare(a));

  months.forEach(month => {
    const data = monthMap[month];

    // カード別明細
    data.cards.forEach(card => {
      summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, card.name, card.amount]]);
      summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
      rowNum++;
    });

    // 引落合計行
    summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, '【引落合計】', data.total]]);
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#e8f5e9');
    summarySheet.getRange(rowNum, 1, 1, 5).setFontWeight('bold');
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    const totalRow = rowNum;
    rowNum++;

    // 残高入力行（黄色）
    const savedBalance = savedBalances[month] || '';
    summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, '三井住友銀行 残高（手入力）', savedBalance]]);
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#fff9c4');
    summarySheet.getRange(rowNum, 2).setFontWeight('bold');
    summarySheet.getRange(rowNum, 3).setNote('三井住友銀行の現在残高をここに入力してください');
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    const balanceRow = rowNum;
    rowNum++;

    // 差引残高行
    summarySheet.getRange(rowNum, 1, 1, 2).setValues([[month, '差引残高（引落後）']]);
    summarySheet.getRange(rowNum, 3).setFormula(
      `=IF(C${balanceRow}="","残高を入力してください",C${balanceRow}-C${totalRow})`
    );
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    summarySheet.getRange(rowNum, 5).setFormula(
      `=IF(C${balanceRow}="","－",IF(C${rowNum}>=0,"✅ 足りる！","⚠️ 不足！"))`
    );
    summarySheet.getRange(rowNum, 5).setFontWeight('bold');
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#e3f2fd');

    // 条件付き書式（不足→赤、足りる→緑）
    const rules = summarySheet.getConditionalFormatRules();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#ffcdd2').setFontColor('#c62828')
        .setRanges([summarySheet.getRange(rowNum, 1, 1, 5)])
        .build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0)
        .setBackground('#c8e6c9').setFontColor('#1b5e20')
        .setRanges([summarySheet.getRange(rowNum, 3, 1, 1)])
        .build()
    );
    summarySheet.setConditionalFormatRules(rules);

    rowNum += 2;
  });

  summarySheet.autoResizeColumns(1, 5);
  summarySheet.activate();
}

// ==============================
// 月次サマリーをメールで送信
// ==============================
function sendMonthlySummaryEmail() {
  const query = 'from:feedback@moneyforward.com subject:引き落としのお知らせ newer_than:35d';
  const threads = GmailApp.search(query);

  const currentMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月');
  const cardData = [];
  const processedIds = new Set();

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const msgId = message.getId();
      if (processedIds.has(msgId)) return;
      processedIds.add(msgId);

      const subject = message.getSubject();
      const body = message.getBody();

      const monthMatch = subject.match(/（(\d{4}年\d{2}月)）/);
      const month = monthMatch ? monthMatch[1] : '';
      if (!month.startsWith(currentMonth.slice(0, 7))) return;

      const cardMatch = subject.match(/】(.+?) 引き落とし/);
      const cardName = cardMatch ? cardMatch[1].trim() : '不明';
      cardData.push({ cardName, billingDate: extractBillingDate(body), amount: extractAmount(body) });
    });
  });

  if (cardData.length === 0) return;

  const total = cardData.reduce((sum, d) => sum + (typeof d.amount === 'number' ? d.amount : 0), 0);

  let emailBody = `【${currentMonth} カード引落まとめ】\n口座：三井住友銀行\n─────────────────\n`;
  cardData.forEach(d => {
    const amtStr = typeof d.amount === 'number' ? '¥' + d.amount.toLocaleString() : d.amount;
    emailBody += `${d.cardName}\n  引落日：${d.billingDate}\n  金 額：${amtStr}\n\n`;
  });
  emailBody += `─────────────────\n合計：¥${total.toLocaleString()}\n`;

  const userEmail = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(userEmail, `【${currentMonth}】カード引落まとめ`, emailBody);
  Logger.log('サマリーメール送信完了：' + userEmail);
}

// ==============================
// トリガー設定（初回に一度だけ実行）
// ==============================
function setupTrigger() {
  // 既存トリガーをすべて削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 6時間ごとに新着チェック（メール到着後 最大6時間以内に自動更新）
  ScriptApp.newTrigger('checkNewEmails')
    .timeBased()
    .everyHours(6)
    .create();

  // 毎月1日の朝9時に全件リフレッシュ＋サマリーメール送信
  ScriptApp.newTrigger('summarizeMoneyForward')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  ScriptApp.newTrigger('sendMonthlySummaryEmail')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('トリガー設定完了！');
  SpreadsheetApp.getUi().alert(
    'トリガー設定完了！\n\n' +
    '✅ 6時間ごと：新着メールを自動チェック＆更新\n' +
    '✅ 毎月1日 9:00：全件リフレッシュ＋サマリーメール送信'
  );
}
