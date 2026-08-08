/**
 * staging 分頁：直接建一個排好格式的分頁，一步到位。
 *
 * 不複製模板、不需要「項目」分頁。staging 只是你自己排版的草稿區，
 * 沒有理由長得像老闆的檔案。A 欄下拉吃程式裡的字串清單，分頁自給自足，
 * 複製到任何活頁簿都還能用。
 */


/**
 * 建立（或補完格式）staging 分頁。可重複執行。
 *
 * 不必顧慮「不要動到使用者排過版的列」——②按下去時會先把舊分頁整個刪掉，
 * 所以這裡碰到的分頁一定是本次執行自己建的。
 */
function ensureStagingSheet_(ss, name, photoCount) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  sheet.setColumnWidth(CFG.C_ITEM, CFG.COL_W_ITEM);
  sheet.setColumnWidth(CFG.C_BEFORE, CFG.COL_W_PHOTO);
  sheet.setColumnWidth(CFG.C_AFTER, CFG.COL_W_PHOTO);
  sheet.setColumnWidth(CFG.C_NOTE, CFG.COL_W_NOTE);

  sheet.getRange(CFG.HEADER_ROW, 1, 1, 4)
    .setValues([CFG.HEADERS_SHEET])
    .setFontWeight('bold');
  sheet.setRowHeight(CFG.HEADER_ROW, CFG.ROW_HEADER);
  sheet.setFrozenRows(CFG.HEADER_ROW);

  var lastNeeded = CFG.FIRST_DATA_ROW + photoCount - 1;
  ensureRowCount_(sheet, lastNeeded);

  if (photoCount > 0) {
    sheet.setRowHeights(CFG.FIRST_DATA_ROW, photoCount, CFG.ROW_DATA);
    sheet.getRange(CFG.FIRST_DATA_ROW, 1, photoCount, 4)
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  applyItemValidation_(sheet, CFG.FIRST_DATA_ROW, photoCount);
  return sheet;
}


/** A 欄的項目下拉。用程式裡的字串清單，不依賴任何「項目」分頁。 */
function applyItemValidation_(sheet, startRow, numRows) {
  if (numRows < 1) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CFG.ITEMS, true)
    .setAllowInvalid(true)        // 原檔 allowBlank=1 且不強制，保持一致
    .build();
  sheet.getRange(startRow, CFG.C_ITEM, numRows, 1).setDataValidation(rule);
}


function ensureRowCount_(sheet, needLastRow) {
  var have = sheet.getMaxRows();
  if (have < needLastRow) sheet.insertRowsAfter(have, needLastRow - have);
}


/**
 * 分頁名以相簿資料夾名為預設值，但要能真的當分頁名用：
 * Google Sheets 禁止 [ ] * / \ ? : ，上限 100 字元，且同名分頁不能並存
 * （Drive 允許兩個同名子資料夾，所以「天生唯一」並不成立）。
 */
function sanitizeSheetName_(raw, taken) {
  var name = String(raw).replace(/[\[\]\*\/\\\?:]/g, '_').trim();
  if (!name) name = 'album';
  if (name.length > 90) name = name.substring(0, 90);
  var base = name, i = 2;
  while (taken.indexOf(name) !== -1) {
    name = base + '_' + i;
    i++;
  }
  return name;
}


/**
 * 發布分頁名 = 相簿名前四碼（MMDD）+ Luna。
 * 沒有四位數字開頭就回空字串——寧可留白讓使用者自己填，也不要寫一個爛預設值
 * 進去，因為「只填空白」的規則會讓爛值再也改不掉。
 */
function derivePublishName_(albumName) {
  var m = String(albumName).match(/^(\d{4})/);
  return m ? m[1] + 'Luna' : '';
}
