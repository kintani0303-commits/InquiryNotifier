/**
 * スプレッドシートを開いたときに独自メニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('問い合わせ管理')
    .addItem('未対応を今すぐ通知', 'runSummaryFromMenu')
    .addSeparator()
    .addItem('初期設定(トリガー登録)', 'setup')
    .addItem('トリガーを解除', 'uninstall')
    .addToUi();
}

/**
 * メニューから実行したとき用:処理後に画面右下へ完了メッセージを出す
 */
function runSummaryFromMenu() {
  sendDailySummary();
  SpreadsheetApp.getActiveSpreadsheet().toast('処理が終わりました。「ログ」シートで結果を確認できます。', '問い合わせ管理');
}
