/**
 * 把排好的 staging 分頁複製到老闆的活頁簿。
 *
 * 為什麼是複製而不是重跑：批次產生只會依檔名排序把照片塞進一欄，
 * 項目、順序、前後欄位分配都是人工在那之後才做的。重跑等於回到手排之前。
 *
 * 這是整套系統裡唯一會碰到共編檔的操作，也是唯一不可逆的一步。
 */

function publishToTarget() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet;
  try {
    cfgSheet = requireConfigSheet_(ss);
  } catch (e) {
    ui.alert('✗ ' + (e.message || e));
    return;
  }

  var rows = readConfigRows_(cfgSheet).filter(function (r) {
    return r.publishOn && r.publishId;
  });
  if (!rows.length) {
    ui.alert('沒有可發布的項目。\n\n' +
             '要發布的那一列需要兩件事：\n' +
             '　1. 「發布目標試算表ID」欄填了目標活頁簿 ID\n' +
             '　2. 「發布」欄已勾選（新相簿預設不勾）');
    return;
  }

  var report = [];
  for (var i = 0; i < rows.length; i++) {
    try {
      report.push(publishOne_(ui, ss, rows[i]));
    } catch (e) {
      report.push('✗ ' + rows[i].album + '：' + (e.message || e));
    }
  }
  ui.alert('發布結果', report.join('\n\n'), ui.ButtonSet.OK);
}


/**
 * 目標活頁簿加上這批照片之後會不會大到打不開。
 *
 * Google 試算表無法開啟超過 100MB 的 Office 檔案，而目標全是 .xlsx。
 * 每張照片的估值來自實測（49 張 ≈ 7.1MB）。抓不到檔案大小就安靜略過——
 * 這只是提醒，不該擋住發布。
 */
function sizeWarning_(fileId, photoCount) {
  var current;
  try {
    current = DriveApp.getFileById(fileId).getSize();
  } catch (e) {
    return '';
  }
  var added = photoCount * CFG.EST_BYTES_PER_PHOTO;
  var after = current + added;
  var mb = function (b) { return (b / 1024 / 1024).toFixed(1) + ' MB'; };
  var line = '　大小：' + mb(current) + ' ＋ 約 ' + mb(added) + ' ＝ 約 ' + mb(after) + '\n';

  if (after > CFG.OFFICE_LIMIT_BYTES) {
    return line +
      '\n⚠⚠ 會超過 100 MB。Google 試算表打不開超過 100 MB 的 Office 檔案，\n' +
      '　　發布之後這份檔案很可能就開不起來了。\n' +
      '　　建議先把舊分頁搬到封存檔，或調低 Config.gs 的 THUMB_WIDTH。\n';
  }
  if (after > CFG.OFFICE_LIMIT_BYTES * CFG.SIZE_WARN_RATIO) {
    return line +
      '\n⚠ 已逼近 100 MB 的開啟上限，再幾次巡檢就會滿。\n';
  }
  return line;
}


function publishOne_(ui, ss, job) {
  if (!job.publishName) {
    throw new Error('「發布分頁名」是空的（相簿名沒有 MMDD 開頭時不會自動產生），請自己填。');
  }
  var src = ss.getSheetByName(job.staging);
  if (!src) throw new Error('找不到 staging 分頁「' + job.staging + '」。');

  // 半成品不該進共編檔。進度是 0 表示沒有執行紀錄（例如換過 schema，
  // 或分頁是很久以前產的），那就不擋——寧可放行也不要誤擋一個好好的分頁。
  var photos = listAlbumPhotos_(job.folderId);
  if (job.written > 0 && job.written < photos.length) {
    throw new Error('這個相簿還沒產完（' + job.written + ' / ' + photos.length +
                    ' 張），先跑完再發布。');
  }

  // openById 失敗會在這裡拋錯，此時還沒對目標做過任何事
  var dst = SpreadsheetApp.openById(job.publishId);
  if (dst.getSheetByName(job.publishName)) {
    throw new Error('目標「' + dst.getName() + '」已經有分頁「' + job.publishName +
                    '」了，為避免覆寫已中止。');
  }

  var answer = ui.alert(
    '確認發布',
    '即將把分頁複製進共用檔案：\n\n' +
    '　來源：' + job.staging + '（' + photos.length + ' 張）\n' +
    '　目標：' + dst.getName() + '\n' +
    '　分頁名：' + job.publishName + '\n' +
    sizeWarning_(job.publishId, photos.length) + '\n' +
    '該檔案的所有協作者都會看到這個新分頁。確定嗎？',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return '－ ' + job.album + '：已取消';

  var copied = src.copyTo(dst);
  try {
    copied.setName(job.publishName);
    dst.setActiveSheet(copied);
    dst.moveActiveSheet(1);
  } catch (e) {
    // copyTo 成功之後才失敗，會在共編檔留下一個叫「xxx 的副本」的分頁。
    // 這是唯一不可逆的操作，收乾淨再把錯誤丟出去。
    try { dst.deleteSheet(copied); } catch (ignored) {}
    throw new Error('複製後設定失敗，已把殘留分頁清掉：' + (e.message || e));
  }

  // A 欄下拉是內嵌字串清單，複製過去本來就自給自足，不需要重新綁定
  // 目標檔的「項目」分頁——那反而會讓下拉依賴一個十幾個人都能改的分頁。
  return '✓ ' + job.album + '：已複製到「' + dst.getName() + '」的「' + job.publishName + '」\n' +
         '　' + dst.getUrl() + '#gid=' + copied.getSheetId();
}
