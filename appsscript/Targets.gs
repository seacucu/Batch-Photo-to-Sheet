/**
 * 相簿 → 老闆的巡房維護活頁簿。
 *
 * 對應關係是常數（一個物件永遠對同一份活頁簿），所以寫在 Config.gs 的 TARGETS，
 * 由程式帶進設定分頁，使用者不必也不該手填——那一欄是輸出，不是輸入。
 *
 * 設定分頁顯示的是「檔案智慧型方塊」，檔名與圖示由 Drive 自己解析，所以程式
 * 不必去抓檔名，也不會顯示過期的名字。
 */

// Sheets API 要求方塊所在的儲存格內容是一個佔位字元，方塊本身掛在 chipRuns 上。
var CHIP_PLACEHOLDER = '@';


/**
 * 把相簿名與老闆的檔名收斂到同一個鍵。
 *
 * 兩邊的落差只有兩處，而且是穩定的慣例：
 *   相簿名多一個 MMDD 前綴  0727唐物a  → 唐物a
 *   老闆檔名多一個「町」     唐物町a    → 唐物a
 * 兩邊都過這個函式，就不用在對照表裡替同一個物件寫兩種寫法。
 */
function normalizeName_(s) {
  return String(s || '')
    .replace(/^\d{4}/, '')
    .replace(/[町\s　]/g, '')
    .trim();
}


function targetUrl_(id) {
  return 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
}


/**
 * 每次呼叫都重建索引（34 次迭代）。不快取——沒有快取就沒有過期的快取。
 *
 * 兩個鍵正規化後撞在一起會直接拋錯。這種情況只可能來自日後手動加行，
 * 而它的後遺症是「安靜地對到錯的活頁簿」，是這套系統最不能容忍的失敗。
 */
function targetIndex_() {
  var out = {};
  for (var raw in CFG.TARGETS) {
    var key = normalizeName_(raw);
    if (out.hasOwnProperty(key) && out[key] !== CFG.TARGETS[raw]) {
      throw new Error('Config.gs 的 TARGETS 有兩個物件正規化後都變成「' + key +
                      '」，無法判斷該用哪一份，請改掉其中一個鍵。');
    }
    out[key] = CFG.TARGETS[raw];
  }
  return out;
}


/** 找到回 { key, id, url }，找不到回 null。絕不猜。 */
function resolveTarget_(albumName) {
  var key = normalizeName_(albumName);
  if (!key) return null;
  var id = targetIndex_()[key];
  return id ? { key: key, id: id, url: targetUrl_(id) } : null;
}


/**
 * 把整個「發布目標」欄重寫一次。
 *
 * 成功與失敗的列走同一條寫入路徑（同一個 fields），否則會出現「文字換掉了
 * 但舊方塊還黏在儲存格上」這種半套狀態。
 *
 * batchUpdate 丟錯就退回純文字，讓①不會因為顯示問題整個失敗——發布時的依據
 * 是 TARGETS 而不是這一欄。但退到文字代表有東西要修，所以會回報。
 */
function writeTargetChips_(cfgSheet, rows) {
  var result = { chips: 0, misses: [], fallback: false, error: '' };
  if (!rows.length) return result;

  var sheetId = cfgSheet.getSheetId();
  var col = CFG.COL_PUBLISH_TARGET - 1;
  var requests = [];
  var plain = [];

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var t = resolveTarget_(r.album);
    var text, chipRuns;
    if (t) {
      text = CHIP_PLACEHOLDER;
      chipRuns = [{ startIndex: 0, chip: { richLinkProperties: { uri: t.url } } }];
    } else {
      result.misses.push(r.album);
      text = '⚠ 對照表裡沒有「' + normalizeName_(r.album) + '」';
      chipRuns = [];
    }
    plain.push({ rowIndex: r.rowIndex, text: t ? t.key : text });
    requests.push({
      updateCells: {
        rows: [{ values: [{ userEnteredValue: { stringValue: text }, chipRuns: chipRuns }] }],
        range: {
          sheetId: sheetId,
          startRowIndex: r.rowIndex - 1, endRowIndex: r.rowIndex,
          startColumnIndex: col, endColumnIndex: col + 1
        },
        fields: 'userEnteredValue,chipRuns'
      }
    });
  }

  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, cfgSheet.getParent().getId());
    result.chips = rows.length - result.misses.length;
  } catch (e) {
    result.fallback = true;
    result.error = e.message || String(e);
    for (var j = 0; j < plain.length; j++) {
      cfgSheet.getRange(plain[j].rowIndex, CFG.COL_PUBLISH_TARGET).setValue(plain[j].text);
    }
  }
  return result;
}
