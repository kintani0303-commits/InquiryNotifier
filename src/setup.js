/**
 * 初期設定:設定シートの作成とトリガーの登録をまとめて行う
 * ※ メニューから実行すること(エディタから実行すると alert でエラーになる)
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createConfigSheetIfMissing(ss);
  createRoutingSheetIfMissing(ss);
  createReplySheetIfMissing(ss);
  deleteOwnTriggers();   // 二重登録を防ぐため、先に古いトリガーを消す

  ScriptApp.newTrigger('notifyOnSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  writeLog('INFO', '初期設定完了: トリガーを登録しました');
  SpreadsheetApp.getUi().alert('初期設定が完了しました。\n\n「設定」シートの通知先を確認し、動作確認ができたらドライランのチェックを外してください。');
}

/**
 * このツールが使うトリガーだけを削除する
 */
function deleteOwnTriggers() {
  const targets = ['notifyOnSubmit', 'sendDailySummary'];
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (targets.includes(triggers[i].getHandlerFunction())) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 「設定」シートがなければ、初期値入りで作成する
 */
function createConfigSheetIfMissing(ss) {
  if (ss.getSheetByName('設定')) {
    return;   // すでにあれば何もしない(ユーザーが入力した設定を消さない)
  }
  const sheet = ss.insertSheet('設定');
  sheet.getRange(1, 1, 7, 2).setValues([
    ['項目', '値'],
    ['通知先', Session.getEffectiveUser().getEmail()],
    ['ドライラン', true],
    ['回答シート名', 'フォームの回答 1'],
    ['CC', ''],
    ['自動返信', false],
    ['差出人名', ''],
  ]);
  sheet.getRange('B3').insertCheckboxes();
  sheet.getRange('B6').insertCheckboxes();
}

/**
 * 「振り分け」シートがなければ、見出し入りで作成する
 */
function createRoutingSheetIfMissing(ss) {
  if (ss.getSheetByName('振り分け')) {
    return;   // すでにあれば何もしない(ユーザーが入力した設定を消さない)
  }
  const sheet = ss.insertSheet('振り分け');
  sheet.getRange(1, 1, 4, 2).setValues([
    ['種別', '宛先'],
    ['見積依頼', ''],
    ['質問', ''],
    ['その他', ''],
  ]);
}

/**
 * 「自動返信」シートがなければ、ひな形入りで作成する
 */
function createReplySheetIfMissing(ss) {
  if (ss.getSheetByName('自動返信')) {
    return;   // すでにあれば何もしない(ユーザーが入力した設定を消さない)
  }
  const sheet = ss.insertSheet('自動返信');
  sheet.getRange(1, 1, 2, 2).setValues([
    ['件名', '【お問い合わせ受付】{{種別}}について'],
    ['本文', '{{氏名}} 様\n\nこのたびはお問い合わせいただき、ありがとうございます。\n「{{種別}}」についてのお問い合わせを受け付けました。\n担当者より2営業日以内にご連絡いたします。\n\n※このメールは送信専用です。ご返信いただいてもお答えできません。'],
  ]);
}

/**
 * このツールのトリガーをすべて解除する(アンインストール)
 */
function uninstall() {
  deleteOwnTriggers();
  writeLog('INFO', 'トリガーを解除しました');
  SpreadsheetApp.getUi().alert('トリガーを解除しました。\n自動通知は停止しています。');
}
