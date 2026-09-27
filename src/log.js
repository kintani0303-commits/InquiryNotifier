/**
 * 「ログ」シートの最終行に1行追記する(シートがなければ作成)
 * @param {string} level   INFO / ERROR など
 * @param {string} message 内容
 */
function writeLog(level, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('ログ');
    logSheet.appendRow(['日時', 'レベル', '内容']);
  }
  logSheet.appendRow([new Date(), level, message]);
}
