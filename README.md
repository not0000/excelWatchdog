# excelWatchdog
**自動監控指定資料夾內的 Excel (`.xlsx`) 檔案**，當 Excel 檔案被修改並存檔時，系統會： 1. **記錄變更內容**（儲存格的修改、刪除、新增） 2. **儲存快照**（Excel 資料內容轉換為 JSON，並存入 `snapshots` 資料夾） 3. **輸出變更紀錄至 `changes.txt`**（詳細記錄 Excel 內容的異動）
