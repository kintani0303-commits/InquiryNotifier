/**
 * 「自動返信」シートから件名と本文のひな形を読み込む
 */
function getReplyTemplate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動返信');
  if (!sheet) {
    throw new Error('「自動返信」シートが見つかりません');
  }
  return {
    subject: sheet.getRange('B1').getValue(),
    body: sheet.getRange('B2').getValue(),
  };
}

/**
 * ひな形の {{項目名}} を実際の値に置き換える
 * 例: fillTemplate('{{氏名}} 様', { '氏名': '山田' }) → '山田 様'
 */
function fillTemplate(text, values) {
  let result = text;
  for (const key in values) {
    result = result.split(`{{${key}}}`).join(values[key]);
  }
  return result;
}

/**
 * メールアドレスの形として正しいかを調べる
 */
function isValidEmail(address) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address);
}

/**
 * フォームに回答した人へ、受付完了の自動返信メールを送る
 */
function sendAutoReply(config, answers) {
  if (config['自動返信'] !== true) {
    writeLog('INFO', '自動返信: オフのため送信しません');
    return;
  }

  const to = getAnswer(answers, 'メール');
  if (!isValidEmail(to)) {
    writeLog('WARN', `自動返信: メールアドレスの形式が正しくないため送信しません(${to})`);
    return;
  }

  // フォームの回答を { '氏名': '山田', '種別': '質問', ... } の形にする
  const values = {};
  for (const key in answers) {
    values[key] = answers[key][0];
  }

  const template = getReplyTemplate();
  const subject = fillTemplate(template.subject, values);
  const body = fillTemplate(template.body, values);
  sendMail(config, to, subject, body, false);   // 回答者宛てなので CC は付けない
}
