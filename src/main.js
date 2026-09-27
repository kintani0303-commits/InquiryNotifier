/**
 * 未対応の一覧をまとめてメールで通知する(毎朝の時間トリガー・メニューから実行)
 */
function sendDailySummary() {
  writeLog('INFO', '処理開始: 未対応まとめ通知');
  try {
    const config = getConfig();
    const sheetName = config['回答シート名'];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    const pending = getPendingRows(sheet);
    writeLog('INFO', `未対応: ${pending.length}件`);

    if (pending.length === 0) {
      writeLog('INFO', '未対応なし。メールは送りません');
    } else {
      const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d');
      const subject = `【要対応】${today} 未対応の問い合わせ ${pending.length}件`;
      const body = buildSummaryBody(pending, getHeader(sheet));
      sendMail(config, config['通知先'], subject, body);
    }
  } catch (e) {
    writeLog('ERROR', e.message);
  }
  writeLog('INFO', '処理終了');
}

/**
 * フォーム送信時に、ステータスに「未対応」を記入して新着通知メールを送る(フォーム送信時トリガー)
 * ※ エディタから直接実行しないこと(e が渡されないためエラーになる)
 */
function notifyOnSubmit(e) {
  writeLog('INFO', '処理開始: フォーム送信通知');
  try {
    const config = getConfig();
    const answers = e.namedValues;
    const name = getAnswer(answers, '氏名');
    const type = getAnswer(answers, '種別');
    const content = getAnswer(answers, '内容');

    // 追加された行のステータスに「未対応」を記入
    const formSheet = e.range.getSheet();
    const row = e.range.getRow();
    const col = getHeader(formSheet).indexOf('ステータス');
    if (col === -1) {
      throw new Error('「ステータス」列が見つかりません');
    }
    formSheet.getRange(row, col + 1).setValue('未対応');

    const to = getRecipientByType(type, config);
    const subject = `【新着】${type}:${name} 様`;
    sendMail(config, to, subject, buildNotifyBody(name, type, content));
    sendAutoReply(config, answers);
  } catch (err) {
    writeLog('ERROR', err.message);
  }
}

/**
 * フォームの回答から、指定した質問の答えを取り出す
 * 質問が見つからなければエラーにする
 */
function getAnswer(answers, title) {
  if (!answers[title]) {
    throw new Error(`フォームに「${title}」という質問が見つかりません`);
  }
  return answers[title][0];
}
