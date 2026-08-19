/**
 * 巡檢照片上傳 — 主流程
 *
 * 一顆按鈕：掃描相簿 → 每個相簿建一個 staging 分頁 → 塞照片 →
 * 撞到執行上限自動續跑。全部發生在腳本所在的這份試算表裡。
 *
 * 「匯入相片」勾了就重新產生、沒勾就完全不碰，沒有任何「幫你判斷是否已完成」
 * 的隱藏邏輯——那會讓勾選框控制不了它看起來在控制的東西。
 *
 * 續跑進度存在 DocumentProperties，使用者看不到也不需要看；他們要看的是
 * 設定分頁的「狀態」欄。刻意不從分頁內容推導，因為範圍讀取看不看得見
 * 儲存格內圖片並未經證實。按②會把進度歸零，所以不需要「重置」按鈕。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Chloe的工具')
    .addItem('① 建立／更新設定分頁', 'setupConfigSheet')
    .addItem('② 產生相簿分頁（匯入照片）', 'generateAll')
    .addItem('③ 發布到正式目標', 'publishToTarget')
    .addSeparator()
    .addItem('刪除本分頁的空白列', 'removeBlankRows')
    .addToUi();
}


/**
 * ① 掃描 Drive，把相簿寫進設定分頁。不碰照片。
 *
 * 跟②分開的用意：這一步會動設定分頁（含 schema 遷移），而②會被觸發器
 * 反覆呼叫。混在一起的話，每一輪續跑都要重掃七個 Drive 資料夾，光掃描就
 * 可能吃掉整輪的時間預算。
 */
function setupConfigSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var cfgSheet = ensureConfigSheet_(ss);
    var before = readConfigRows_(cfgSheet).length;
    var chips = syncConfigFromDrive_(ss, cfgSheet);
    var after = readConfigRows_(cfgSheet).length;

    var msg = '設定分頁已更新。\n\n' +
              '相簿共 ' + after + ' 個' +
              (after > before ? '（新增 ' + (after - before) + ' 個）' : '（沒有新增）') + '。\n';

    if (chips.misses.length) {
      msg += '\n⚠ 這 ' + chips.misses.length + ' 個相簿找不到對應的發布目標，' +
             '不會被發布：\n　' + chips.misses.join('\n　') +
             '\n（要在 Config.gs 的 TARGETS 對照表補上）\n';
    }
    if (chips.fallback) {
      msg += '\n⚠ 檔案方塊寫不進去，「發布目標」欄改成純文字。\n　' + chips.error +
             '\n　發布功能不受影響，但這件事要修。\n';
    }
    ui.alert(msg + '\n接著按「② 產生相簿分頁」。');
  } catch (e) {
    ui.alert('✗ ' + (e.message || e));
  }
}


/**
 * ② 選單進入點（有 UI）。
 *
 * 語意就是字面上的意思：**勾了「匯入相片」就重新產生，沒勾就完全不碰。**
 * 不做「已經產完所以幫你跳過」這種聰明事——那會讓勾選框控制不了它看起來
 * 在控制的東西。排好版之後把該列取消勾選，它就永遠不會再被動到。
 *
 * 按下去等於「重來」：先把要處理的分頁刪掉、進度歸零，再從頭產生。
 * 動手前會把清單攤開讓你確認。
 */
function generateAll() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet;
  try {
    cfgSheet = requireConfigSheet_(ss);
  } catch (e) {
    ui.alert('✗ ' + (e.message || e));
    return;
  }

  var targets = readConfigRows_(cfgSheet).filter(function (r) { return r.importOn; });
  if (!targets.length) {
    ui.alert('「匯入相片」欄一個都沒勾，沒有要處理的相簿。');
    return;
  }

  var lines = targets.map(function (r) {
    var exists = ss.getSheetByName(r.staging) ? '（既有分頁會被換掉）' : '（新建）';
    return '　・' + r.album + ' → ' + r.staging + ' ' + exists;
  });
  var answer = ui.alert(
    '產生相簿分頁',
    '會重新產生以下 ' + targets.length + ' 個分頁：\n\n' + lines.join('\n') +
    '\n\n沒有勾選「匯入相片」的相簿完全不會被碰到。\n繼續嗎？',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  // 重來：刪掉舊分頁、進度歸零。之後續跑才有乾淨的起點。
  var folderIds = [];
  for (var i = 0; i < targets.length; i++) {
    var old = ss.getSheetByName(targets[i].staging);
    if (old) ss.deleteSheet(old);
    folderIds.push(targets[i].folderId);
    setCell_(cfgSheet, targets[i].rowIndex, CFG.COL_NOTE, '');
  }
  clearProgress_(folderIds);
  SpreadsheetApp.flush();

  runCycle_(true);
}


/** 觸發器進入點（沒有 UI，任何 getUi() 都會拋錯）。 */
function resumeGenerate() {
  runCycle_(false);
}


function runCycle_(fromMenu) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    // 選單執行與觸發器執行不會自動互斥。兩邊同時跑會重複寫入、
    // 而且各自再排一個觸發器。
    if (fromMenu) {
      SpreadsheetApp.getUi().alert('另一個執行正在進行中，請等它跑完。');
    }
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = null;
  try {
    // 這一步不建立也不遷移設定分頁——那是①的職責。這裡只讀。
    cfgSheet = requireConfigSheet_(ss);

    // 保險觸發器排在任何執行都不可能還活著之後。正常結束會刪掉它；
    // 若這次執行被 6 分鐘硬上限砍死，排下一輪的程式碼根本來不及跑，
    // 就靠它把鏈子接回來。
    deleteResumeTriggers_();
    createResumeTrigger_(CFG.SAFETY_TRIGGER_MS);

    var deadline = new Date().getTime() + CFG.SOFT_TIME_LIMIT_MS;
    var r = processAlbums_(ss, cfgSheet, deadline);

    deleteResumeTriggers_();
    var msg;
    if (r.remaining === 0) {
      msg = '✓ 全部完成。' + summarize_(r);
    } else if (r.hitDeadline || r.written > 0) {
      createResumeTrigger_(CFG.NEXT_CYCLE_MS);
      msg = '執行中，' + CFG.NEXT_CYCLE_MS / 1000 + ' 秒後自動續跑（不用再按）。'
          + summarize_(r);
    } else {
      // 沒撞到時限、卻一張也沒寫進去，表示卡住了。再排下去只是空轉。
      msg = '⚠ 已停止：這一輪沒有任何進展，還有 ' + r.remaining
          + ' 個相簿未完成。請看各列的「狀態」欄。';
    }
    setStatus_(cfgSheet, msg);
    if (fromMenu) SpreadsheetApp.getUi().alert(msg);

  } catch (e) {
    // 先前建的保險觸發器必須拆掉。留著的話，一個永久性的錯誤（例如相簿根
    // 資料夾不見了）會變成每 7 分鐘重試一次、永遠不停。
    try { deleteResumeTriggers_(); } catch (ignored) {}
    var err = '✗ 已停止：' + (e.message || e);
    try { if (cfgSheet) setStatus_(cfgSheet, err); } catch (ignored2) {}
    if (fromMenu) SpreadsheetApp.getUi().alert(err);
    else throw e;              // 讓 Google 寄出觸發器失敗通知信
  } finally {
    lock.releaseLock();
  }
}


function summarize_(r) {
  return ' 本輪寫入 ' + r.written + ' 張，未完成相簿 ' + r.remaining + ' 個。';
}


function processAlbums_(ss, cfgSheet, deadline) {
  var rows = readConfigRows_(cfgSheet);
  var written = 0, remaining = 0, hitDeadline = false;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r.importOn) continue;

    if (new Date().getTime() > deadline) { hitDeadline = true; remaining++; continue; }

    try {
      var photos = listAlbumPhotos_(r.folderId);
      if (!photos.length) {
        setCell_(cfgSheet, r.rowIndex, CFG.COL_NOTE, '相簿裡沒有圖片');
        continue;
      }

      // 分頁不在就從頭寫。進度是「本次執行」的位置——②按下去時已經歸零，
      // 所以這裡不為零只可能是同一次執行被時限打斷後的續跑。
      var sheet = ss.getSheetByName(r.staging);
      var start = sheet ? Math.min(r.written, photos.length) : 0;

      if (start >= photos.length) {     // 這一輪之前就寫完了，續跑時不必重來
        setCell_(cfgSheet, r.rowIndex, CFG.COL_NOTE, '完成 ' + photos.length + ' 張');
        continue;
      }

      sheet = ensureStagingSheet_(ss, r.staging, photos.length);
      var res = writePhotos_(sheet, photos, CFG.PHOTO_COL, start, deadline);

      written += res.written - start;
      setProgress_(r.folderId, res.written);

      if (res.written < photos.length) {
        hitDeadline = true;
        remaining++;
        setCell_(cfgSheet, r.rowIndex, CFG.COL_NOTE,
                 '進行中 ' + res.written + ' / ' + photos.length);
      } else {
        setCell_(cfgSheet, r.rowIndex, CFG.COL_NOTE, describeDone_(photos.length, res.failures));
      }
    } catch (e) {
      setCell_(cfgSheet, r.rowIndex, CFG.COL_NOTE, '✗ ' + (e.message || e));
      remaining++;
    }
  }
  return { written: written, remaining: remaining, hitDeadline: hitDeadline };
}


function describeDone_(total, failures) {
  if (!failures.length) return '完成 ' + total + ' 張';
  var list = failures.slice(0, 5).map(function (f) { return '第' + f.row + '列 ' + f.name; });
  if (failures.length > 5) list.push('等 ' + failures.length + ' 張');
  return '完成 ' + total + ' 張，其中 ' + failures.length + ' 張抓不到：' + list.join('、');
}


// ---------------------------------------------------------------- 設定分頁

/** 只讀取，不建立也不遷移。給②和續跑用——它們不該偷偷改 schema。 */
function requireConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CFG.CONFIG_SHEET);
  if (!sheet) {
    throw new Error('找不到「' + CFG.CONFIG_SHEET + '」分頁，請先執行「① 建立／更新設定分頁」。');
  }
  var head = sheet.getRange(1, 1, 1, CFG.HEADERS_CONFIG.length).getValues()[0]
    .map(function (v) { return String(v).trim(); });
  if (head.join('') !== CFG.HEADERS_CONFIG.join('')) {
    throw new Error('「' + CFG.CONFIG_SHEET + '」分頁的欄位格式不符，請先執行「① 建立／更新設定分頁」。');
  }
  return sheet;
}


function ensureConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CFG.CONFIG_SHEET);
  if (!sheet) return buildConfigSheet_(ss, []);

  var width = CFG.HEADERS_CONFIG.length;
  var head = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (v) {
    return String(v).trim();
  });
  var ok = head.join('') === CFG.HEADERS_CONFIG.join('');
  if (ok) return sheet;

  // Schema 變了：把舊表留著（加時間戳，第二次遷移才不會撞名），
  // 並把兩個核取欄的狀態依「來源資料夾ID」帶到新表。
  var carried = harvestOldRows_(sheet);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  sheet.setName(CFG.CONFIG_SHEET + '_舊_' + stamp);
  return buildConfigSheet_(ss, carried);
}


/**
 * 換 schema 時，把兩個核取欄的狀態從舊表帶到新表，以 folderId 為鍵。
 *
 * 為什麼非帶不可：排好版的相簿，「匯入相片」是被**取消**勾選的。重建時若一律
 * 回到預設的勾選狀態，之後誤按②就會把排好的版整個換掉。這件事發生過一次。
 *
 * 靠**標題名稱**定位欄位，不靠欄號——遷移的前提就是欄序變了。找不到標題就不帶，
 * 那時新表會用預設值（匯入=勾、發布=不勾）。
 */
function harvestOldRows_(oldSheet) {
  var out = [];
  try {
    var last = oldSheet.getLastRow();
    var wide = oldSheet.getLastColumn();
    if (last < 2 || wide < 1) return out;

    var head = oldSheet.getRange(1, 1, 1, wide).getValues()[0].map(function (v) {
      return String(v).trim();
    });
    var fCol = head.indexOf('來源資料夾ID');
    var iCol = head.indexOf('匯入相片');
    var pCol = head.indexOf('發布');
    if (fCol === -1) return out;

    var vals = oldSheet.getRange(2, 1, last - 1, wide).getValues();
    for (var i = 0; i < vals.length; i++) {
      var folderId = String(vals[i][fCol]).trim();
      if (!folderId) continue;
      out.push({
        folderId: folderId,
        importOn: iCol === -1 ? null : isChecked_(vals[i][iCol]),
        publishOn: pCol === -1 ? null : isChecked_(vals[i][pCol])
      });
    }
  } catch (ignored) {}
  return out;
}


function buildConfigSheet_(ss, carried) {
  var sheet = ss.insertSheet(CFG.CONFIG_SHEET, 0);
  var w = CFG.HEADERS_CONFIG.length;
  sheet.getRange(1, 1, 1, w).setValues([CFG.HEADERS_CONFIG]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(CFG.COL_ALBUM, 130);
  sheet.setColumnWidth(CFG.COL_FOLDER_ID, 300);
  sheet.setColumnWidth(CFG.COL_STAGING, 130);
  sheet.setColumnWidth(CFG.COL_PUBLISH_TARGET, 220);   // 放檔案方塊，不再是 44 字元的 ID
  sheet.setColumnWidth(CFG.COL_PUBLISH_NAME, 110);
  sheet.setColumnWidth(CFG.COL_NOTE, 380);
  // 需要解釋的欄位把說明寫在標題儲存格的備註裡，滑鼠移上去就看得到
  sheet.getRange(1, CFG.COL_PUBLISH_TARGET).setNote(CFG.NOTE_TARGET);
  sheet.getRange(1, CFG.COL_IMPORT).setNote(CFG.NOTE_IMPORT);
  sheet.getRange(1, CFG.COL_PUBLISH_ON).setNote(CFG.NOTE_PUBLISH);
  sheet.getRange(CFG.STATUS_CELL).setValue('尚未執行');
  if (carried && carried.length) {
    PropertiesService.getDocumentProperties()
      .setProperty('carriedRows', JSON.stringify(carried));
  }
  return sheet;
}


/**
 * 掃描根資料夾的子資料夾，把缺的列補進設定分頁。
 * 配對依據是來源資料夾 ID，不是名稱——資料夾改名不會變成重複列。
 *
 * 規則：**只填空白，絕不覆蓋既有值**，包含兩個核取欄在內。
 * **唯一的例外是「發布目標」欄**——它是從相簿名推導出來的衍生值，不是使用者
 * 的資料，所以每次都整欄重寫。若對它也套用「不覆蓋」，一個過期的 ⚠ 就會
 * 永遠黏在那裡，即使對照表早就補好了。
 */
function syncConfigFromDrive_(ss, cfgSheet) {
  var existing = readConfigRows_(cfgSheet);
  var byFolder = {};
  var takenNames = [];
  for (var i = 0; i < existing.length; i++) {
    byFolder[existing[i].folderId] = existing[i];
    if (existing[i].staging) takenNames.push(existing[i].staging);
  }

  var carried = {};
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('carriedRows');
    if (raw) {
      JSON.parse(raw).forEach(function (c) { carried[c.folderId] = c; });
      PropertiesService.getDocumentProperties().deleteProperty('carriedRows');
    }
  } catch (ignored) {}

  var it = DriveApp.getFolderById(CFG.ROOT_FOLDER_ID).getFolders();
  var found = [];
  while (it.hasNext()) {
    var f = it.next();
    found.push({ id: f.getId(), name: f.getName() });
  }
  found.sort(function (a, b) {
    return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
  });

  for (var j = 0; j < found.length; j++) {
    var alb = found[j];
    var row = byFolder[alb.id];
    if (row) {
      // 只補空白，不動任何既有值
      if (!row.staging) {
        var nm = sanitizeSheetName_(alb.name, takenNames);
        takenNames.push(nm);
        setCell_(cfgSheet, row.rowIndex, CFG.COL_STAGING, nm);
      }
      if (!row.publishName) {
        var pn = derivePublishName_(alb.name);
        if (pn) setCell_(cfgSheet, row.rowIndex, CFG.COL_PUBLISH_NAME, pn);
      }
      continue;
    }
    var name = sanitizeSheetName_(alb.name, takenNames);
    takenNames.push(name);
    var prev = carried[alb.id] || {};
    var newRow = cfgSheet.getLastRow() + 1;
    // 先插核取方塊再寫值——反過來的話 insertCheckboxes 可能把值蓋掉
    cfgSheet.getRange(newRow, CFG.COL_IMPORT).insertCheckboxes();
    cfgSheet.getRange(newRow, CFG.COL_PUBLISH_ON).insertCheckboxes();
    cfgSheet.getRange(newRow, 1, 1, CFG.HEADERS_CONFIG.length).setValues([[
      alb.name, alb.id, name,
      '',        // 發布目標：留白，等下由 writeTargetChips_ 填方塊
      derivePublishName_(alb.name),
      // 遷移過來的列沿用舊狀態；真正的新相簿才用預設（匯入=開、發布=關）
      prev.importOn === null || prev.importOn === undefined ? true : prev.importOn,
      prev.publishOn === null || prev.publishOn === undefined ? false : prev.publishOn,
      ''
    ]]);
  }

  // 整欄重寫，新舊列都一樣。放在最後一次做，只呼叫一次 API。
  return writeTargetChips_(cfgSheet, readConfigRows_(cfgSheet));
}


function readConfigRows_(cfgSheet) {
  var last = cfgSheet.getLastRow();
  if (last < 2) return [];
  var w = CFG.HEADERS_CONFIG.length;
  var vals = cfgSheet.getRange(2, 1, last - 1, w).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    var folderId = String(v[CFG.COL_FOLDER_ID - 1]).trim();
    if (!folderId) continue;
    out.push({
      rowIndex: i + 2,
      album: String(v[CFG.COL_ALBUM - 1]).trim(),
      folderId: folderId,
      staging: String(v[CFG.COL_STAGING - 1]).trim(),
      // 刻意不讀「發布目標」欄。它是顯示用的方塊，發布時的依據是 resolveTarget_()，
      // 這樣就不存在「欄位內容跟實際寫入目標不一致」的可能。
      publishName: String(v[CFG.COL_PUBLISH_NAME - 1]).trim(),
      importOn: isChecked_(v[CFG.COL_IMPORT - 1]),
      publishOn: isChecked_(v[CFG.COL_PUBLISH_ON - 1]),
      written: getProgress_(folderId)
    });
  }
  return out;
}


function isChecked_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}


// ------------------------------------------------- 續跑進度（使用者看不到）

/**
 * 「寫到第幾張」是純粹的實作細節，使用者沒有理由要看它，所以不放在設定分頁，
 * 存進腳本專用的 DocumentProperties。人看得懂的進度由「狀態」欄負責顯示。
 *
 * 以相簿資料夾 ID 為鍵。②按下去時會把要處理的相簿清成 0，所以殘留的舊值
 * 最多只活到下一次按②為止。
 */
function getProgress_(folderId) {
  var v = PropertiesService.getDocumentProperties().getProperty(CFG.PROGRESS_PREFIX + folderId);
  return Number(v) || 0;
}


function setProgress_(folderId, n) {
  PropertiesService.getDocumentProperties()
    .setProperty(CFG.PROGRESS_PREFIX + folderId, String(n));
}


function clearProgress_(folderIds) {
  var props = PropertiesService.getDocumentProperties();
  for (var i = 0; i < folderIds.length; i++) {
    props.deleteProperty(CFG.PROGRESS_PREFIX + folderIds[i]);
  }
}


function setCell_(cfgSheet, rowIndex, col, value) {
  cfgSheet.getRange(rowIndex, col).setValue(value);
}


function setStatus_(cfgSheet, text) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm:ss');
  cfgSheet.getRange(CFG.STATUS_CELL).setValue('[' + stamp + '] ' + text);
}


// ---------------------------------------------------------------- 觸發器

/**
 * 依處理常式名稱刪，不用 e.triggerUid——從選單或編輯器呼叫時根本沒有事件物件。
 */
function deleteResumeTriggers_() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === CFG.RESUME_HANDLER) {
      ScriptApp.deleteTrigger(all[i]);
    }
  }
}


function createResumeTrigger_(afterMs) {
  ScriptApp.newTrigger(CFG.RESUME_HANDLER).timeBased().after(afterMs).create();
}
