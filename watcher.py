import os
import json
import time
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

WATCH_FOLDER = r"D:\excel修改python"
SNAPSHOT_FOLDER = os.path.join(WATCH_FOLDER, "snapshots")
CHANGES_FOLDER = os.path.join(WATCH_FOLDER, "changes")

# 確保資料夾存在
os.makedirs(SNAPSHOT_FOLDER, exist_ok=True)
os.makedirs(CHANGES_FOLDER, exist_ok=True)

def get_timestamp():
    return time.strftime("%Y%m%d%H%M%S")

def read_excel(file_path):
    try:
        df_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
        data = {sheet: df.astype(str).fillna("").to_dict() for sheet, df in df_sheets.items()}
        return data
    except Exception as e:
        print(f"讀取 Excel 錯誤: {e}")
        return None

def save_snapshot(file_name, data, timestamp):
    snapshot_path = os.path.join(SNAPSHOT_FOLDER, f"{file_name}_{timestamp}.json")
    with open(snapshot_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    return snapshot_path

def compare_snapshots(old_data, new_data):
    changes = {}
    for sheet in new_data:
        if sheet not in old_data:
            changes[sheet] = {"added": new_data[sheet]}
        else:
            sheet_changes = {}
            for row_key, new_row in new_data[sheet].items():
                if row_key not in old_data[sheet]:
                    sheet_changes[row_key] = {"added": new_row}
                else:
                    row_changes = {}
                    for col_key, new_value in new_row.items():
                        old_value = old_data[sheet][row_key].get(col_key, "")
                        if new_value != old_value:
                            row_changes[col_key] = {"old": old_value, "new": new_value}
                    if row_changes:
                        sheet_changes[row_key] = row_changes
            if sheet_changes:
                changes[sheet] = sheet_changes
    return changes

def save_changes(file_name, changes, timestamp):
    if changes:
        change_path = os.path.join(CHANGES_FOLDER, f"{file_name}_{timestamp}.json")
        with open(change_path, "w", encoding="utf-8") as f:
            json.dump(changes, f, ensure_ascii=False, indent=4)

def get_latest_snapshot(file_name):
    files = [f for f in os.listdir(SNAPSHOT_FOLDER) if f.startswith(file_name) and f.endswith(".json")]
    if not files:
        return None
    latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(SNAPSHOT_FOLDER, x)))
    with open(os.path.join(SNAPSHOT_FOLDER, latest_file), "r", encoding="utf-8") as f:
        return json.load(f)

class ExcelChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.is_directory or not event.src_path.endswith(".xlsx"):
            return
        
        file_name = os.path.basename(event.src_path)
        if file_name.startswith("~$"):
            return  # 忽略 Excel 產生的暫存檔

        time.sleep(2)  # 增加等待時間，確保檔案寫入完成
        print(f"偵測到 Excel 變更: {event.src_path}")
        
        new_data = read_excel(event.src_path)
        if new_data:
            old_data = get_latest_snapshot(file_name)
            timestamp = get_timestamp()
            if old_data:
                changes = compare_snapshots(old_data, new_data)
                save_changes(file_name, changes, timestamp)
            save_snapshot(file_name, new_data, timestamp)

if __name__ == "__main__":
    event_handler = ExcelChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()
    print(f"開始監控資料夾: {WATCH_FOLDER}")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
