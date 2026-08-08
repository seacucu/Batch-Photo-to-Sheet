/**
 * 排版輔助工具。不屬於正式管線，選單上以分隔線隔開。
 */


/**
 * 刪掉目前這個分頁裡 A–D 四欄全空的列，把排版時留下的空隙收攏。
 * 只處理標題列以下、最後一列以上的範圍。刪錯了 Ctrl+Z 就回來了。
 */
function removeBlankRows() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = sheet.getName();

  if (name.indexOf(CFG.CONFIG_SHEET) === 0) {
    ui.alert('這是設定分頁，不是照片分頁。請先切換到要整理的相簿分頁。');
    return;
  }

  var last = sheet.getLastRow();
  if (last <= CFG.FIRST_DATA_ROW - 1) {
    ui.alert('「' + name + '」沒有資料列。');
    return;
  }

  var first = CFG.FIRST_DATA_ROW;
  var values = sheet.getRange(first, 1, last - first + 1, 4).getValues();
  var blanks = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var empty = true;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null) { empty = false; break; }
    }
    if (empty) blanks.push(first + i);
  }

  if (!blanks.length) {
    ui.alert('「' + name + '」沒有空白列。');
    return;
  }

  var answer = ui.alert(
    '刪除空白列',
    '「' + name + '」裡有 ' + blanks.length + ' 列是空的（第 ' +
    blanks[0] + ' 列起）。\n\n刪掉它們？下面的列會往上遞補。',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  // 由下往上刪，否則刪過的位移會讓後面的列號全部失準。
  // 連續的列合併成一次 deleteRows，少幾十次 API 呼叫。
  var removed = 0;
  var i2 = blanks.length - 1;
  while (i2 >= 0) {
    var end = blanks[i2];
    var start = end;
    while (i2 > 0 && blanks[i2 - 1] === start - 1) { i2--; start = blanks[i2]; }
    sheet.deleteRows(start, end - start + 1);
    removed += end - start + 1;
    i2--;
  }

  SpreadsheetApp.flush();
  ui.alert('已從「' + name + '」刪除 ' + removed + ' 列。');
}
