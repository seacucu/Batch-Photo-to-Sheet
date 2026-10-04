/**
 * 巡檢照片上傳 — 常數
 *
 * 版面數值全部從老闆那份「唐物町a巡房維護.xlsx」反解出來。
 * xlsx 用「點」記列高、「字元」記欄寬，Google Sheets 兩者都用像素，
 * 換算比例取自 Google 自己匯出時寫的預設值：
 *   defaultRowHeight="15.0" ↔ Sheets 預設 21px  → 1pt = 1.4px
 *   defaultColWidth="14.43" ↔ Sheets 預設 100px → 1字元 = 6.93px
 */

var CFG = {
  // 「維護照片相簿」。相簿都是它的子資料夾，靠掃描取得，
  // 新增相簿不必改程式。這是整個專案唯一寫死的 ID。
  ROOT_FOLDER_ID: '1QseLWqp_RANH-zAi9x2OLKjt801HZXGg',

  CONFIG_SHEET: '設定',
  STATUS_CELL: 'J1',          // 觸發器沒有 UI，整體狀態寫這裡。在表格八欄之外，
                              // 附加列不會把它推走。
  RESUME_HANDLER: 'resumeGenerate',
  PROGRESS_PREFIX: 'written:', // 續跑進度存在 DocumentProperties，不佔設定分頁的欄位

  // 設定分頁欄位。順序即欄號，改這裡等於改 schema（會觸發遷移）。
  HEADERS_CONFIG: ['相簿名稱', '來源資料夾ID', 'staging分頁名',
                   '發布目標', '發布分頁名',
                   '匯入相片', '發布', '狀態'],
  COL_ALBUM: 1,
  COL_FOLDER_ID: 2,
  COL_STAGING: 3,
  COL_PUBLISH_TARGET: 4,      // 程式自動帶入的檔案方塊。是輸出不是輸入，沒人手填
  COL_PUBLISH_NAME: 5,
  COL_IMPORT: 6,              // 核取：要不要產生分頁／匯入照片
  COL_PUBLISH_ON: 7,          // 核取：要不要發布。新列預設不勾——發布是唯一
                              // 對外且不可逆的動作，應該是主動開啟而非預設武裝
  COL_NOTE: 8,

  NOTE_IMPORT: '勾選：按②時，這個相簿的分頁會被重新產生，照片重新匯入。\n' +
               '　　　　既有的分頁會整個換掉，排好的版也會沒有。\n' +
               '取消：完全不碰這一列。\n\n' +
               '排版完成後請把這一欄取消勾選，之後按②就不會再動到它。',
  NOTE_PUBLISH: '勾選：「發布到正式目標」會把這一列的分頁複製進「發布目標」欄那份活頁簿。\n' +
                '取消：不發布。\n\n' +
                '新相簿預設不勾——排版完成、確認無誤之後再自己勾起來。',
  NOTE_TARGET: '這一欄由程式自動帶入，不用也不要手填。\n\n' +
               '對應關係是「相簿名去掉日期 → 老闆的巡房維護活頁簿」，\n' +
               '寫在 Config.gs 的 TARGETS 對照表裡。\n' +
               '顯示 ⚠ 表示對照表裡沒有這個物件，那一列不會被發布。',

  // staging / 目標分頁版面
  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2,
  ROW_HEADER: 18,             //  12.75pt
  COL_W_ITEM: 127,            // 18.29 字元
  COL_W_PHOTO: 171,           // 想要照片大一點／小一點，只要改這個數字，列高會跟著算
  COL_W_NOTE: 182,            // 26.29 字元
  PHOTO_ASPECT: 3 / 4,        // 來源 4032×3024，縮圖 1000×750，一律 4:3
  HEADERS_SHEET: ['項目', '写真(前)', '写真(後)', '備註'],
  C_ITEM: 1,
  C_BEFORE: 2,
  C_AFTER: 3,
  C_NOTE: 4,
  PHOTO_COL: 3,               // 照片一律進 写真(後)；前後配對不在範圍內

  // 照片抓取
  THUMB_WIDTH: 1000,          // 少了這個參數只會拿到約 200px 的預設縮圖
  BATCH_SIZE: 20,
  RETRY_BASE_MS: 1500,        // 逐格退避的指數退避基數
  RETRY_MAX: 3,

  // 檔案大小警告。Google 試算表打不開超過 100MB 的 Office 檔案，而目標活頁簿
  // 全是 .xlsx，其中好幾份已經逼近或越過這條線。
  // 每張照片的估值取自實測：49 張的試算表是 7,100,865 bytes，約 145KB/張。
  OFFICE_LIMIT_BYTES: 100 * 1024 * 1024,
  SIZE_WARN_RATIO: 0.9,
  EST_BYTES_PER_PHOTO: 149000,

  // 續跑
  SOFT_TIME_LIMIT_MS: 4 * 60 * 1000,
  SAFETY_TRIGGER_MS: 7 * 60 * 1000,   // 保險觸發器：排在任何執行都不可能還活著之後
  NEXT_CYCLE_MS: 30 * 1000,

  // 發布目標對照表 — 老闆「3.巡房維護」資料夾底下的 34 份活頁簿。
  //
  // 鍵是檔名去掉「巡房維護.xlsx」剩下的部分，**照抄 Drive 上的寫法**（含「町」），
  // 這樣才能跟 Drive 逐字對照、確認沒抄錯。比對時兩邊都會過 normalizeName_()，
  // 所以相簿叫「0727唐物a」也對得上「唐物町a」。
  //
  // 這些 ID 是常數，不是每輪巡檢會變的東西，所以放在程式裡而不是讓人每次手打——
  // 打錯一個字元的後果是把照片寫進別人物件的共編檔。
  // 老闆新增物件時在這裡加一行；註解是它在 Drive 上的位置。
  TARGETS: {
    '唐物町a':      '1yS-rV2BVnCBDu5WobI8MFr80tVtvevCC',  // 京都房源/唐物區
    '唐物町b':      '1Y-FKo2pJ1kqMevzkbN4nq9cU_DDBxmel',
    '唐物町c':      '1o3kyAu5CQPoH1yAy7oRWDJBC4h25Ra-c',
    '寶柿町':       '1L6_vz5On6DszHupJs4yhWYWObeP5xK3f',
    '京吹雪':       '1i1esSXS4Qx4DZINg_WRKlovhcaoHkBym',  // 京都房源/東山區
    '緣宿':         '1bf-dXUVpoxHAqKXagql2F1VZ-7io25Y1',
    '風舍':         '1-pji6kkqgR1culPthfLSUSCphMZpMmrZ',
    '念慈庵':       '1_zm4ZvWpm5onkQNS2KMKL3IOjLJE7I72',
    '東福寺':       '1FhQJYXeU0cBTco70ZRWVMjSkq-q7rqMX',
    '有緣舍':       '1alZLoU6PipxrfLhxjs6JdedgcD-jv4sZ',
    '無憂舍':       '1swJv-W4weYKnTcHeKZTzb3mk-AQC7mcT',
    // 針小路a 在同一個資料夾裡另有一份 290MB 的 .gsheet（原生格式的轉檔產物）。
    // 這裡取 .xlsx，跟其餘 33 份一致。
    '針小路a':      '1A6GPxDCLU4kcEXWldjOGWtaUJ0NOU9Oq',  // 京都房源/京都站區
    '針小路b':      '1P38hdXQxxJvgAvipiR_owMDfUaem5l21',
    '藤木町':       '1AOh2dKjCua82Kzcafyk5sy7z7vCM9EH8',
    '洛思崖':       '1TkH5JC6SlJ5dKb9UsFyfAmcgCS4hJELD',
    '鞍馬口':       '1UdeICgANgyhdq48NrtXSkIyNN_Wa-p-Y',
    '季京都1a':     '1wt3noJldwEZWyd54OiwEdUadwlet_ZOd',
    '季京都2b':     '1usoBQhaWoYHFbSBbweqbA26GlATDQGVI',
    '季京都2c':     '12-N957GYmyvlsbujrMmv8XnylRzTvRgp',
    '五条a':        '1Agx8iVloNtUvfZ5G35aykVQZ-0NEFj83',  // 京都房源/五條區
    '五条b':        '1lNOSWuB0qyTLg_rTadiO-SVwZw6Xr1pk',
    '柿本町':       '1hEECRCHwEYZnA30-t9_KPGjNSOzi_U0L',
    '松原':         '1Pj_RDDM8xyGKD1J6SYDux0toqovg3orP',
    'モコ深江橋':   '1oQ1s2qFYiDsbu4PaaC1znxGQhOXhN8NI',  // 大阪房源
    'モコ中道':     '1uvHMDFXhrdpJz5AYVp_dpeVzDWAk7iE5',
    'モコ九條':     '1ha5pydkPawigNe4brSUTHB4iymW2OKQx',
    'モコ森ノ宮':   '1gftT_Yg1D-LHU1To7M-TnEVFFf-GeJPT',
    'ジブラ天六':   '1NYPA1b05hegnSl6z2A1v47UFNrw7cyHb',
    'ジブラ天王寺': '1hxiuS5H5U50an8JtvxUbhBXgc01-1iig',
    'ジブラ九條':   '1B6c1tYEg5Elf64Njb6_VnHct7UysTmVp',
    'ジブラ森ノ宮': '1b16RFt6yi9G8hLaIXKt4ySHL5rV8hwg1',
    'ジブラ浮田':   '1Le3MW8xQacEJQHhRzgn61vslX3Xp7qY7',
    'ジブラ阿倍野': '1ZIw0idOAo3Qy6nuVZ7oD9FZ1lDpJ3m2q',
    '小魚':         '1O4_g04K0DmVkdFOBtIO1gdYtx4lAo5sP'
  },

  // 「項目」下拉清單。37 項，逗號串接約 274 字元，遠低於 500 字元上限。
  ITEMS: [
    '牆壁"明顯刮痕"',
    '地板"明顯刮痕"',
    '燈具/燈罩',
    '掛衣框/掛衣桿',
    '滅火器灰塵',
    '門鎖電池',
    '畫框',
    '裝飾物/小盆栽/花瓶',
    '裝飾框',
    '門框/紙門框',
    '窗框',
    '樓梯扶手',
    '障子紙維護',
    '櫃子角落污漬',
    '鞋櫃上灰塵',
    '開關/面板灰塵',
    '抽屜污漬',
    '抱枕/坐墊污漬',
    '投影機灰塵',
    '投影機濾網',
    '投影機設定確認',
    '冷氣出風口',
    '冷氣機濾網',
    '冰箱除霜',
    '冰箱上灰塵',
    '水龍頭/水龍頭把手"確認是否鎖緊"',
    '浴室"污漬黴菌"',
    '浴缸金屬環',
    '廁所抽風機',
    '浴室抽風機',
    '廚房抽風機',
    '前後庭院紗窗/玻璃窗擦拭',
    '前後庭院落葉',
    '門口雜草/蜘蛛網清除',
    '庭院灑水確認',
    '外面走道清理',
    '床品狀態確認(是否有污漬/破損...)'
  ]
};

/**
 * 資料列高由照片欄寬推導，而不是另外寫一個常數。
 *
 * 儲存格內圖片是「符合儲存格大小」等比縮放，所以 COL_W_PHOTO 寬的欄位裡，圖片
 * 只會用到 COL_W_PHOTO × PHOTO_ASPECT 的高度（171px → 128px）。列高設得比這個
 * 大，多出來的就是純粹的死空間，五十列就多捲一兩千像素；設得比這個小則換成
 * 左右留白。改 COL_W_PHOTO 時這裡不必跟著改，列高是算出來的。
 *
 * 原本抄的 161px 來自 xlsx 模板的 114.75pt，但那是給固定 2.0×1.5 吋的
 * 浮動圖片用的，跟儲存格內圖片無關。
 */
CFG.ROW_DATA = Math.round(CFG.COL_W_PHOTO * CFG.PHOTO_ASPECT);
