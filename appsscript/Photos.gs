/**
 * 相簿列舉與照片寫入。
 *
 * 照片用 SpreadsheetApp.newCellImage() 建立儲存格內圖片。實測結果：
 * 產出的 CellImage 的 getUrl() 是空的、getContentUrl() 指向
 * googleusercontent — 圖片已烘進試算表，不依賴相簿的共用權限。
 *
 * 不要走舊腳本那套 =IMAGE() → copyTo(PASTE_VALUES) → 刪來源：實測在現在的
 * Google Sheets 上會得到字串 #REF!，同分頁內操作亦然。
 *
 * Drive 的 thumbnail 端點會在伺服器端把 HEIC 解成 JPEG 並套好 irot 旋轉，
 * 所以完全不需要在別處處理 HEIC。
 */


/** 列出相簿裡的圖片，依檔名自然排序（本語料的檔名序 == 拍攝時間序）。 */
function listAlbumPhotos_(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var it = folder.getFiles();
  var out = [];
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType().indexOf('image/') === 0) {
      out.push({ name: f.getName(), id: f.getId() });
    }
  }
  out.sort(function (a, b) {
    return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
  });
  return out;
}


function thumbUrl_(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w' + CFG.THUMB_WIDTH;
}


function cellImageRow_(photo) {
  return [
    SpreadsheetApp.newCellImage()
      .setSourceUrl(thumbUrl_(photo.id))
      .setAltTextTitle(photo.name)      // 版面上沒有欄位放檔名，存在替代文字裡
      .build()
  ];
}


/**
 * 從第 alreadyWritten 張開始，把照片寫進指定欄。
 *
 * setValues 會在寫入當下逐一去抓 URL，整批只要有一張抓不到就整批拋例外、
 * 整批都不會寫入（曾經因此 49 張全滅）。所以批次失敗時退回逐格寫，
 * 讓損害侷限在真正壞掉的那一張。
 *
 * 回傳 { written, failures }。written 是「這一欄現在總共有幾張」，
 * 抓不到的那些**也算進去**——它們的格子留有可見的佔位文字，是一個已處理的
 * 位置，不會讓進度卡在那裡永遠前進不了。
 */
function writePhotos_(sheet, photos, col, alreadyWritten, deadline) {
  var done = Math.max(0, alreadyWritten || 0);
  var failures = [];

  while (done < photos.length) {
    if (new Date().getTime() > deadline) break;

    var n = Math.min(CFG.BATCH_SIZE, photos.length - done);
    var slice = photos.slice(done, done + n);
    var startRow = CFG.FIRST_DATA_ROW + done;

    try {
      sheet.getRange(startRow, col, n, 1).setValues(slice.map(cellImageRow_));
    } catch (batchErr) {
      // 整批失敗最常見的原因是節流。先喘一口氣再逐格試，
      // 否則四十次連發只會繼續被擋。
      Utilities.sleep(CFG.RETRY_BASE_MS);
      for (var i = 0; i < n; i++) {
        writeOnePhoto_(sheet, slice[i], startRow + i, col, failures);
      }
    }
    done += n;
  }
  SpreadsheetApp.flush();
  return { written: done, failures: failures };
}


/** 單格寫入，指數退避重試。真的救不回就留下可見的佔位文字。 */
function writeOnePhoto_(sheet, photo, row, col, failures) {
  var cell = sheet.getRange(row, col, 1, 1);
  var lastErr = '';
  for (var attempt = 0; attempt < CFG.RETRY_MAX; attempt++) {
    try {
      cell.setValues([cellImageRow_(photo)]);
      return true;
    } catch (e) {
      lastErr = String(e.message || e);
      if (attempt < CFG.RETRY_MAX - 1) {
        Utilities.sleep(CFG.RETRY_BASE_MS * Math.pow(2, attempt));
      }
    }
  }
  // 佔位文字而不是清空：空格子會讓「這個相簿做完了沒」永遠判定為否，
  // 整個相簿就此卡死。留字則是一個已處理的位置，而且你一眼看得到要補哪一列。
  cell.setValue('⚠ 抓不到 ' + photo.name);
  failures.push({ name: photo.name, row: row, error: lastErr });
  return false;
}
