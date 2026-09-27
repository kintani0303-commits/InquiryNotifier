/**
 * 「設定」シートを読み込み、項目名をキーにしたオブジェクトで返す
 * 例: { '通知先': 'a@example.com', 'ドライラン': true, '回答シート名': 'フォームの回答 1', 'CC': '', ... }
 */
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) {
    throw new Error('「設定」シートが見つかりません');
  }
  const values = sheet.getDataRange().getValues();

  const config = {};
  for (let i = 1; i < values.length; i++) {
    config[values[i][0]] = values[i][1];
  }

  if (!config['通知先']) {
    throw new Error('設定シートの「通知先」が空です');
  }
  return config;
}

/**
 * 「振り分け」シートから、種別に対応する宛先を返す
 * 見つからないとき・シートがないときは設定シートの「通知先」を返す
 */
function getRecipientByType(type, config) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('振り分け');
  if (!sheet) {
    return config['通知先'];
  }

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === type && values[i][1]) {   // 種別が一致して、宛先が空でない
      return values[i][1];                          // 見つかった時点で返して終了
    }
  }
  return config['通知先'];                          // 最後まで見つからなかった
}
