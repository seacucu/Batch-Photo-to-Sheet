/**
 * 把排好的 staging 分頁複製到老闆的活頁簿。
 *
 * 為什麼是複製而不是重跑：批次產生只會依檔名排序把照片塞進一欄，
 * 項目、順序、前後欄位分配都是人工在那之後才做的。重跑等於回到手排之前。
 *
 * 這是整套系統裡唯一會碰到共編檔的操作，也是唯一不可逆的一步。
 *
 * 刻意分成「全部檢查 → 一次確認 → 全部執行」三段，而不是一份一份問。
 * 一份一份問的問題不只是手痠：問到第五份時，前四份已經寫進共編檔了，
 * 而你是在那之後才看到第五份有問題。先全部攤開再決定，才有「全看完再按」
 * 這個選項。
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

  var rows = readConfigRows_(cfgSheet).filter(function (r) { return r.publishOn; });
  if (!rows.length) {
    ui.alert('沒有可發布的項目。\n\n' +
             '設定分頁裡要發布的那一列，「發布」欄要勾起來（新相簿預設不勾）。');
    return;
  }

  // ---- 1. 全部檢查。這一段只讀，不對任何檔案動手。
  var plans = [];
  var blocked = [];
  for (var i = 0; i < rows.length; i++) {
    var p;
    try {
      p = planPublish_(ss, rows[i]);
    } catch (e) {
      p = { ok: false, album: rows[i].album, reason: e.message || String(e) };
    }
    if (p.ok) plans.push(p); else blocked.push(p);
  }

  markDuplicates_(plans, blocked);
  plans = plans.filter(function (p) { return p.ok; });

  if (!plans.length) {
    ui.alert('發布', '沒有任何一列可以發布：\n\n' + blockedLines_(blocked), ui.ButtonSet.OK);
    return;
  }

  // ---- 2. 一次確認，之後不再問。
  if (!confirmAll_(ui, plans, blocked)) {
    ui.alert('已取消，什麼都沒有動。');
    return;
  }

  // ---- 3. 全部執行。每一筆獨立成敗。
  var report = [];
  for (var k = 0; k < plans.length; k++) {
    try {
      report.push(executePublish_(plans[k]));
    } catch (e) {
      report.push('✗ ' + plans[k].album + '：' + (e.message || e));
    }
  }
  for (var m = 0; m < blocked.length; m++) {
    report.push('－ ' + blocked[m].album + '（略過）：' + blocked[m].reason);
  }
  ui.alert('發布結果', report.join('\n\n'), ui.ButtonSet.OK);
}


/**
 * 檢查一列能不能發布，順便把執行時要用的東西都查好。
 *
 * 只讀不寫。回傳 { ok: true, ... } 或 { ok: false, album, reason }，
 * 因為「這一列不能發布」是預期內的結果而不是例外——它要能跟其他列並列
 * 顯示在同一個確認框裡。
 */
function planPublish_(ss, job) {
  var fail = function (reason) {
    return { ok: false, album: job.album, reason: reason };
  };

  if (!job.publishName) {
    return fail('「發布分頁名」是空的（相簿名沒有 MMDD 開頭時不會自動產生），請自己填。');
  }
  var src = ss.getSheetByName(job.staging);
  if (!src) return fail('找不到 staging 分頁「' + job.staging + '」。');

  // 目標在這裡重新解析，不讀設定分頁的「發布目標」欄——那一欄只是給人看的方塊。
  // 唯一的真相來源是 Config.gs 的對照表。
  var target = resolveTarget_(job.album);
  if (!target) {
    return fail('對應到物件名「' + normalizeName_(job.album) +
                '」，但 Config.gs 的 TARGETS 對照表裡沒有這一項，不知道要發布到哪。');
  }

  // 半成品不該進共編檔。進度是 0 表示沒有執行紀錄（例如換過 schema，
  // 或分頁是很久以前產的），那就不擋——寧可放行也不要誤擋一個好好的分頁。
  var photos = listAlbumPhotos_(job.folderId);
  if (job.written > 0 && job.written < photos.length) {
    return fail('還沒產完（' + job.written + ' / ' + photos.length + ' 張），先跑完再發布。');
  }

  // openById 失敗會在這裡拋錯，此時還沒對目標做過任何事
  var dst = SpreadsheetApp.openById(target.id);
  if (dst.getSheetByName(job.publishName)) {
    return fail('目標「' + dst.getName() + '」已經有分頁「' + job.publishName +
                '」了，為避免覆寫已略過。');
  }

  return {
    ok: true,
    album: job.album,
    staging: job.staging,
    publishName: job.publishName,
    src: src,
    dst: dst,
    targetId: target.id,
    targetName: dst.getName(),
    photoCount: photos.length
  };
}


/**
 * 兩列要寫進同一份目標檔的同一個分頁名時，只有第一筆會成功。
 *
 * 這種撞名本來要等 copyTo 之後才會發現，而那時已經不可逆了。既然現在
 * 整批計畫都攤在手上，就在確認之前先攔掉後來的那些。
 */
function markDuplicates_(plans, blocked) {
  var seen = {};
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    if (!p.ok) continue;
    var key = p.targetId + '\n' + p.publishName;
    if (seen[key]) {
      p.ok = false;
      p.reason = '跟「' + seen[key] + '」都要寫進「' + p.targetName + '」的分頁「' +
                 p.publishName + '」，撞名了。請改掉其中一個的「發布分頁名」。';
      blocked.push(p);
    } else {
      seen[key] = p.album;
    }
  }
}


function blockedLines_(blocked) {
  var out = [];
  for (var i = 0; i < blocked.length; i++) {
    out.push('　・' + blocked[i].album + '：' + blocked[i].reason);
  }
  return out.join('\n');
}


function mb_(bytes) {
  return (bytes / 1024 / 1024).toFixed(1) + ' MB';
}


/**
 * 目標活頁簿加上這批照片之後會不會大到打不開。
 *
 * Google 試算表無法開啟超過 100MB 的 Office 檔案，而目標全是 .xlsx。
 * 每張照片的估值來自實測（49 張 ≈ 7.1MB）。抓不到檔案大小就安靜略過——
 * 這只是提醒，不該擋住發布。
 *
 * 按目標檔分組累加：兩個相簿發到同一份檔案時，要看的是加完兩批之後的大小。
 * 一份一份算會各自都說「還沒滿」，合起來卻已經爆了。
 */
function sizeReport_(plans) {
  var byTarget = {};
  var order = [];
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    if (!byTarget[p.targetId]) {
      byTarget[p.targetId] = { name: p.targetName, photos: 0 };
      order.push(p.targetId);
    }
    byTarget[p.targetId].photos += p.photoCount;
  }

  var lines = [];
  var worst = 0;
  for (var j = 0; j < order.length; j++) {
    var id = order[j];
    var t = byTarget[id];
    var current;
    try {
      current = DriveApp.getFileById(id).getSize();
    } catch (e) {
      continue;                         // 抓不到大小就不提這一份
    }
    var added = t.photos * CFG.EST_BYTES_PER_PHOTO;
    var after = current + added;
    var ratio = after / CFG.OFFICE_LIMIT_BYTES;
    if (ratio > worst) worst = ratio;

    var mark = '';
    if (ratio > 1) mark = '⚠⚠ ';
    else if (ratio > CFG.SIZE_WARN_RATIO) mark = '⚠ ';
    lines.push('　' + mark + t.name + '：' + mb_(current) + ' ＋ 約 ' + mb_(added) +
               ' ＝ 約 ' + mb_(after));
  }
  return { lines: lines, worst: worst };
}


/**
 * 整批的唯一一個確認框。
 *
 * 要讓人按得下去，它必須回答「會動哪些共編檔、各加幾張、加完多大、
 * 哪幾列不會動」。按下「是」之後就不再問任何問題了，所以該看的東西
 * 都得在這裡看得到。
 */
function confirmAll_(ui, plans, blocked) {
  var body = '即將把 ' + plans.length + ' 個分頁複製進共用檔案：\n\n';
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    body += '　・' + p.staging + '（' + p.photoCount + ' 張）\n' +
            '　　→ ' + p.targetName + ' ／ 分頁「' + p.publishName + '」\n';
  }

  var size = sizeReport_(plans);
  if (size.lines.length) {
    body += '\n檔案大小預估：\n' + size.lines.join('\n') + '\n';
    if (size.worst > 1) {
      body += '\n⚠⚠ 標了 ⚠⚠ 的會超過 100 MB。Google 試算表打不開超過 100 MB 的\n' +
              '　　Office 檔案，發布之後那份檔案很可能就開不起來了。\n' +
              '　　建議先把舊分頁搬到封存檔，或調低 Config.gs 的 THUMB_WIDTH。\n';
    } else if (size.worst > CFG.SIZE_WARN_RATIO) {
      body += '\n⚠ 標了 ⚠ 的已逼近 100 MB 的開啟上限，再幾次巡檢就會滿。\n';
    }
  }

  if (blocked.length) {
    body += '\n以下 ' + blocked.length + ' 列不會被發布：\n' + blockedLines_(blocked) + '\n';
  }

  body += '\n這些檔案的所有協作者都會看到新分頁。\n' +
          '按「是」就一次全部發布，不會再逐一詢問。';

  return ui.alert('確認發布', body, ui.ButtonSet.YES_NO) === ui.Button.YES;
}


/**
 * 真正動手的一段。到這裡所有檢查都過了、使用者也確認過了。
 *
 * 每一筆獨立成敗：第三筆失敗不該回滾前兩筆（它們是各自獨立的分頁，
 * 而且已經在共編檔裡了），也不該讓第四筆不做。
 */
function executePublish_(p) {
  var copied = p.src.copyTo(p.dst);
  try {
    copied.setName(p.publishName);
    p.dst.setActiveSheet(copied);
    p.dst.moveActiveSheet(1);
  } catch (e) {
    // copyTo 成功之後才失敗，會在共編檔留下一個叫「xxx 的副本」的分頁。
    // 這是唯一不可逆的操作，收乾淨再把錯誤丟出去。
    try { p.dst.deleteSheet(copied); } catch (ignored) {}
    throw new Error('複製後設定失敗，已把殘留分頁清掉：' + (e.message || e));
  }

  // A 欄下拉是內嵌字串清單，複製過去本來就自給自足，不需要重新綁定
  // 目標檔的「項目」分頁——那反而會讓下拉依賴一個十幾個人都能改的分頁。
  return '✓ ' + p.album + '：已複製到「' + p.targetName + '」的「' + p.publishName + '」\n' +
         '　' + p.dst.getUrl() + '#gid=' + copied.getSheetId();
}
