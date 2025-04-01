# excelWatchdog 📌 Excel 變更監控系統 – 使用說明  

## **🔧 功能概述**
本系統會**自動監控指定資料夾內的 Excel (`.xlsx`) 檔案**，當 Excel 檔案被修改並存檔時，系統會：
1. **記錄變更內容**（儲存格的修改、刪除、新增）
2. **儲存快照**（Excel 資料內容轉換為 JSON，並存入 `snapshots` 資料夾）
3. **輸出變更紀錄至 `changes.txt`**（詳細記錄 Excel 內容的異動）

---

## **📂 目錄結構**
```plaintext
D:\excel修改python\
│── watcher.py          # 監控程式
│── changes.txt         # 變更紀錄檔案
│── data.xlsx           # 監控的 Excel 檔案
└── snapshots\          # 存放 Excel 快照
    ├── data.xlsx_20250401153045.json
    ├── data.xlsx_20250401153110.json
    ├── data.xlsx_20250401153230.json
```

---

## **🚀 如何使用**
### **1️⃣ 安裝 Python 環境**
請先確認已安裝 **Python 3.10 以上**，並安裝必要的套件：
```sh
pip install pandas openpyxl watchdog
```

---

### **2️⃣ 設定監控資料夾**
請修改 `watcher.py` 內的 `WATCH_FOLDER` 變數，設定要監控的 Excel 檔案所在資料夾：
```python
WATCH_FOLDER = r"D:\excel修改python"
```

---

### **3️⃣ 啟動監控程式**
執行以下指令啟動監控：
```sh
python watcher.py
```
成功執行後，終端機會顯示：
```plaintext
開始監控資料夾: D:\excel修改python
```

---

### **4️⃣ 觸發監控機制**
當 `data.xlsx` 被修改並存檔時：
1. **`snapshots/` 資料夾內會產生新的 JSON 檔案**
   ```plaintext
   snapshots/
   ├── data.xlsx_20250401153045.json
   ├── data.xlsx_20250401153110.json
   ├── data.xlsx_20250401153230.json
   ```
2. **`changes.txt` 內會記錄 Excel 的變更**
   ```plaintext
   Sheet1 儲存格 1,1 | 原值: '100' → 新值: '150'
   Sheet1 儲存格 2,3 | 原值: '500' → 新值: '600'
   Sheet1 儲存格 3,1 | 新增: 'NEW'
   Sheet1 儲存格 3,2 | 新增: 'TEST'
   ```

---

## **🛑 如何停止監控**
監控程式會持續執行，若要停止：
1. 在終端機按 **`Ctrl + C`**
2. 或 **關閉命令提示字元（CMD/PowerShell）視窗**

---

## **📌 功能特色**
✅ **自動偵測 Excel 變更，無需手動執行**  
✅ **每次變更都產生新的 JSON 檔案，方便追蹤歷史版本**  
✅ **變更內容詳細記錄在 `changes.txt`，清楚掌握 Excel 的異動**  
✅ **忽略 Excel 暫存檔 (`~$data.xlsx`)，避免錯誤發生**  

📌 **適用場景**：資料審核、Excel 變更追蹤、版本管理

---

試試執行看看吧！🚀
