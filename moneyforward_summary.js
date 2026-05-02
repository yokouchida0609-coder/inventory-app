/**
 * マネーフォワード カード引落まとめスクリプト
 *
 * 【使い方】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 > Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. 「summarizeMoneyForward」を実行（初回はGmail権限の許可が必要）
 * 5. 毎月自動実行したい場合：「setupTrigger」を一度実行
 */

// ==============================
// メイン処理：引落メールをまとめる
// ==============================
function summarizeMoneyForward() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シートの準備
  let sheet = ss.getSheetByName('引落まとめ');
  if (!sheet) {
    sheet = ss.insertSheet('引落まとめ');
  }
  sheet.clearContents();

  // ヘッダー
  sheet.getRange(1, 1, 1, 6).setValues([
    ['年月', 'カード名', '引落予定日', '引落予定額', '口座', '取得日時']
  ]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  sheet.getRange(1, 1, 1, 6).setBackground('#4a86e8');
  sheet.getRange(1, 1, 1, 6).setFontColor('#ffffff');

  // 直近3ヶ月分のマネーフォワードメールを検索
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
      const body = message.getBody(); // HTML本文

      // 件名から年月を抽出：「（2026年05月）」
      const monthMatch = subject.match(/（(\d{4}年\d{2}月)）/);
      const month = monthMatch ? monthMatch[1] : '不明';

      // 件名からカード名を抽出：「】カード名 引き落とし」
      const cardMatch = subject.match(/】(.+?) 引き落とし/);
      const cardName = cardMatch ? cardMatch[1].trim() : '不明';

      // HTML本文から引落予定日を抽出
      const billingDate = extractBillingDate(body);

      // HTML本文から引落予定額を抽出
      const amount = extractAmount(body);

      rows.push([month, cardName, billingDate, amount, '三井住友銀行', new Date()]);
    });
  });

  if (rows.length === 0) {
    sheet.getRange(2, 1).setValue('データなし（直近90日のメールが見つかりませんでした）');
    return;
  }

  // 年月→カード名 の順でソート
  rows.sort((a, b) => {
    if (a[0] !== b[0]) return b[0].localeCompare(a[0]); // 年月降順
    return a[1].localeCompare(b[1]); // カード名昇順
  });

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);

  // 金額列を数値書式に（抽出できた場合）
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('¥#,##0');

  // 列幅を自動調整
  sheet.autoResizeColumns(1, 6);

  // 月別合計シートを作成（残高チェックつき）
  createMonthlySummary(ss, rows);

  SpreadsheetApp.getUi().alert('完了！引落まとめを更新しました。\n「月別合計」シートの黄色セルに残高を入力してください。');
  Logger.log('完了：' + rows.length + '件処理しました');
}

// ==============================
// 引落予定日を抽出
// ==============================
function extractBillingDate(html) {
  // HTMLタグを除去してテキスト化
  const text = html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

  // パターン1: 「引き落とし予定日 2026年05月27日」
  const p1 = text.match(/引き落とし予定日\s*(\d{4}年\d{2}月\d{2}日)/);
  if (p1) return p1[1];

  // パターン2: 「引落日 2026/05/27」
  const p2 = text.match(/引落[^\d]*(\d{4}\/\d{2}\/\d{2})/);
  if (p2) return p2[1];

  // パターン3: 「05月27日」だけの場合
  const p3 = text.match(/引き落とし予定日[^\d]*(\d{2}月\d{2}日)/);
  if (p3) return p3[1];

  return '不明';
}

// ==============================
// 引落予定額を抽出
// ==============================
function extractAmount(html) {
  // HTMLタグを除去してテキスト化
  const text = html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');

  // パターン1: 「引き落とし予定額 ¥12,345」
  const p1 = text.match(/引き落とし予定額\s*[¥￥]?([\d,]+)\s*円?/);
  if (p1) return parseInt(p1[1].replace(/,/g, ''), 10);

  // パターン2: 「¥12,345」
  const p2 = text.match(/[¥￥]([\d,]+)/);
  if (p2) return parseInt(p2[1].replace(/,/g, ''), 10);

  // パターン3: 「12,345円」
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

  // 既存の残高入力値を保持（再実行時に消えないように）
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

  // ヘッダー
  summarySheet.getRange(1, 1, 1, 5).setValues([['年月', '項目', '金額', '', '判定']]);
  summarySheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  summarySheet.getRange(1, 1, 1, 5).setBackground('#34a853');
  summarySheet.getRange(1, 1, 1, 5).setFontColor('#ffffff');

  // 月別にグループ化
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
    const startRow = rowNum;

    // カード別明細
    data.cards.forEach(card => {
      const amtCell = summarySheet.getRange(rowNum, 3);
      summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, card.name, card.amount]]);
      amtCell.setNumberFormat('¥#,##0');
      rowNum++;
    });

    // 引落合計行
    summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, '【引落合計】', data.total]]);
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#e8f5e9');
    summarySheet.getRange(rowNum, 1, 1, 5).setFontWeight('bold');
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    const totalRow = rowNum;
    rowNum++;

    // 残高入力行（黄色・手入力）
    const savedBalance = savedBalances[month] || '';
    summarySheet.getRange(rowNum, 1, 1, 3).setValues([[month, '三井住友銀行 残高（手入力）', savedBalance]]);
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#fff9c4'); // 黄色
    summarySheet.getRange(rowNum, 2).setFontWeight('bold');
    summarySheet.getRange(rowNum, 3).setNote('三井住友銀行の現在残高をここに入力してください');
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');
    const balanceRow = rowNum;
    rowNum++;

    // 差引残高行（残高 - 引落合計）
    const diffFormula = `=IF(C${balanceRow}="","残高を入力してください",C${balanceRow}-C${totalRow})`;
    summarySheet.getRange(rowNum, 1, 1, 2).setValues([[month, '差引残高（引落後）']]);
    summarySheet.getRange(rowNum, 3).setFormula(diffFormula);
    summarySheet.getRange(rowNum, 3).setNumberFormat('¥#,##0');

    // 判定（足りる／不足）
    const judgmentFormula = `=IF(C${balanceRow}="","－",IF(C${rowNum}>=0,"✅ 足りる！","⚠️ 不足！"))`;
    summarySheet.getRange(rowNum, 5).setFormula(judgmentFormula);
    summarySheet.getRange(rowNum, 5).setFontWeight('bold');
    summarySheet.getRange(rowNum, 1, 1, 5).setBackground('#e3f2fd');

    // 不足のとき赤・足りるとき緑の条件付き書式
    const diffRange = summarySheet.getRange(rowNum, 3);
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#ffcdd2')
      .setFontColor('#c62828')
      .setRanges([summarySheet.getRange(rowNum, 1, 1, 5)])
      .build();
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0)
      .setBackground('#c8e6c9')
      .setFontColor('#1b5e20')
      .setRanges([summarySheet.getRange(rowNum, 3, 1, 1)])
      .build();
    const rules = summarySheet.getConditionalFormatRules();
    rules.push(rule1);
    rules.push(rule2);
    summarySheet.setConditionalFormatRules(rules);

    rowNum += 2; // 月間の空行
  });

  summarySheet.autoResizeColumns(1, 5);

  // 残高入力セルをアクティブに
  summarySheet.activate();
}

// ==============================
// 月次サマリーをメールで受信
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
      if (!month.includes(currentMonth.slice(0, 7))) return; // 当月以外はスキップ

      const cardMatch = subject.match(/】(.+?) 引き落とし/);
      const cardName = cardMatch ? cardMatch[1].trim() : '不明';
      const billingDate = extractBillingDate(body);
      const amount = extractAmount(body);

      cardData.push({ cardName, billingDate, amount });
    });
  });

  if (cardData.length === 0) {
    Logger.log('当月の引落メールなし');
    return;
  }

  // 合計計算
  const total = cardData.reduce((sum, d) => {
    return sum + (typeof d.amount === 'number' ? d.amount : 0);
  }, 0);

  // メール本文作成
  let body = `【${currentMonth} カード引落まとめ】\n`;
  body += `口座：三井住友銀行\n`;
  body += `─────────────────\n`;
  cardData.forEach(d => {
    const amtStr = typeof d.amount === 'number'
      ? '¥' + d.amount.toLocaleString()
      : d.amount;
    body += `${d.cardName}\n`;
    body += `  引落日：${d.billingDate}\n`;
    body += `  金 額：${amtStr}\n\n`;
  });
  body += `─────────────────\n`;
  body += `合計：¥${total.toLocaleString()}\n`;

  // 自分宛にメール送信
  const userEmail = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(userEmail, `【${currentMonth}】カード引落まとめ`, body);
  Logger.log('サマリーメール送信完了：' + userEmail);
}

// ==============================
// 毎月自動実行トリガーを設定
// ==============================
function setupTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 毎月1日の午前9時に実行
  ScriptApp.newTrigger('summarizeMoneyForward')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  // メール送信も毎月1日
  ScriptApp.newTrigger('sendMonthlySummaryEmail')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('トリガー設定完了！毎月1日 9:00 に自動実行されます');
  SpreadsheetApp.getUi().alert('トリガー設定完了！毎月1日 9:00 に自動実行されます。');
}
