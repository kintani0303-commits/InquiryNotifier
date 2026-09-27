/**
 * 未対応まとめメールの本文を作る
 */
function buildSummaryBody(pending, header) {
  const nameCol = header.indexOf('氏名');
  const typeCol = header.indexOf('種別');

  const lines = [];
  for (let i = 0; i < pending.length; i++) {
    lines.push(`・${pending[i][nameCol]}(${pending[i][typeCol]})`);
  }
  return `未対応の問い合わせが ${pending.length} 件あります。\n\n` + lines.join('\n');
}

/**
 * 新着通知メールの本文を作る
 */
function buildNotifyBody(name, type, content) {
  return `フォームから新しい問い合わせが届きました。\n\n氏名:${name}\n種別:${type}\n内容:\n${content}`;
}

/**
 * メールを送信する。設定でドライランがオンなら送らずにログだけ残す
 * CC が設定されていれば CC にも送る
 */
function sendMail(config, to, subject, body) {
  const cc = config['CC'] || '';

  const options = {};
  if (cc) {
    options.cc = cc;
  }

  const ccText = cc || 'なし';
  if (config['ドライラン'] === true) {
    writeLog('INFO', `[ドライラン] 宛先: ${to} / CC: ${ccText} / 件名: ${subject}`);
    console.log(body);
  } else {
    MailApp.sendEmail(to, subject, body, options);
    writeLog('INFO', `メール送信: ${to} / CC: ${ccText} / 件名: ${subject}`);
  }
}
