# legacy

`original-prototype.gs` 是這個專案最早的版本：一顆「批次上傳照片」選單項、
問你 Folder ID、把整個資料夾的照片倒進當前分頁。

留著它是因為它記錄了一條**現在走不通的路**，而那條路看起來很合理、很容易再被想到：

```js
// 寫 =IMAGE() 公式 → flush → copyTo(PASTE_VALUES) 轉成靜態圖 → 刪掉暫存欄
var formula = '=IMAGE("' + imageUrl + '")';
formulaRange.copyTo(targetStaticRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
```

實測在現在的 Google Sheets 上，`PASTE_VALUES` 會把那一欄變成字串 `#REF!`，
同一個分頁內操作也一樣。現行版本改用 `SpreadsheetApp.newCellImage()`
（見 `appsscript/Photos.gs` 的檔頭註解）。

其他跟現行版本的差異：

| | 原型 | 現行 |
|---|---|---|
| 來源 | 手打一個 Folder ID | 掃「維護照片相簿」的子資料夾 |
| 目標 | 當前分頁，`sheet.clear()` | 每個相簿一個 staging 分頁 |
| 圖片 | `=IMAGE()` 公式（連結，會失效） | `newCellImage()`（烘進檔案） |
| 執行上限 | 沒處理，大相簿會中途死掉 | 軟時限 + 觸發器續跑 |
| 版面 | 無 | 欄寬列高、凍結標題、37 項下拉 |
| 發布 | 無 | ③ copyTo 進目標活頁簿 |

這個檔案不會被 `clasp push` 上傳（`.claspignore` 只放行 `appsscript/` 底下的 `*.gs`），
放在這裡純粹是備查。
