/**
 * シートの1行目(見出し)を配列で返す
 */
function getHeader(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * ステータスが「未対応」の行だけを2次元配列で返す(見出し行は含まない)
 */
function getPendingRows(sheet) {
  const values = sheet.getDataRange().getValues();
  const col = values[0].indexOf('ステータス');
  if (col === -1) {
    throw new Error('「ステータス」列が見つかりません');
  }

  const pending = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][col] === '未対応') {
      pending.push(values[i]);
    }
  }
  return pending;
}
