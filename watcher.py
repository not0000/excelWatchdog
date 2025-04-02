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
        df_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl", header=None)
        data = {}

        for sheet, df in df_sheets.items():
            df = df.applymap(lambda x: str(x) if pd.notna(x) else "")  # **確保所有數據轉成字串**
            data[sheet] = df.to_dict(orient="index")

        return data
    except Exception as e:
        print(f"讀取 Excel 錯誤: {e}")
        return None
    
        
def save_snapshot(file_name, data):
    snapshot_path = os.path.join(SNAPSHOT_FOLDER, f"{file_name}_{int(time.time())}.json")

    # **確保所有值都轉為字串**
    data_str = json.loads(json.dumps(data, default=str))

    with open(snapshot_path, "w", encoding="utf-8") as f:
        json.dump(data_str, f, ensure_ascii=False, indent=4)


def compare_snapshots(old_data, new_data):
    changes = {}

    for sheet in new_data:
        if sheet not in old_data:
            changes[sheet] = {"added": new_data[sheet]}
        else:
            sheet_changes = {}

            for row_idx, new_row in new_data[sheet].items():
                old_row = old_data[sheet].get(str(row_idx), {})

                row_changes = {}
                for col_idx, new_value in new_row.items():
                    old_value = old_row.get(str(col_idx), None)
                    
                    # **確保 old_value 和 new_value 都是字串**
                    old_value = str(old_value) if old_value is not None else ""
                    new_value = str(new_value)

                    if new_value != old_value:
                        row_changes[col_idx] = {"old": old_value, "new": new_value}

                if row_changes:
                    sheet_changes[row_idx] = row_changes

            if sheet_changes:
                changes[sheet] = sheet_changes

    return changes

def save_changes(file_name, changes, timestamp):
    if changes:
        change_path = os.path.join(CHANGES_FOLDER, f"{file_name}_{timestamp}.json")
        with open(change_path, "w", encoding="utf-8") as f:
            json.dump(changes, f, ensure_ascii=False, indent=4)

def get_latest_snapshot(file_name):
    snapshot_files = [
        f for f in os.listdir(SNAPSHOT_FOLDER)
        if f.startswith(file_name) and f.endswith(".json")
    ]

    if not snapshot_files:
        return None  # 沒有快照時直接返回 None

    latest_file = max(
        snapshot_files, key=lambda x: os.path.getmtime(os.path.join(SNAPSHOT_FOLDER, x))
    )

    latest_path = os.path.join(SNAPSHOT_FOLDER, latest_file)
    with open(latest_path, "r", encoding="utf-8") as f:
        return json.load(f)



class ExcelChangeHandler(FileSystemEventHandler):
    def __init__(self):
        self.last_modified_times = {}

    def on_created(self, event):
        """當有新 Excel 檔案新增時，建立初始快照"""
        if event.is_directory or not event.src_path.endswith(".xlsx"):
            return
        file_name = os.path.basename(event.src_path)
        if file_name.startswith("~$"):
            return  # 忽略 Excel 產生的暫存檔

        time.sleep(2)  # 等待檔案寫入完成
        print(f"偵測到新增 Excel 檔案: {file_name}")

        new_data = read_excel(event.src_path)
        if new_data:
            timestamp = get_timestamp()
            save_snapshot(file_name, new_data, timestamp)

    def on_modified(self, event):
        """當 Excel 檔案被修改時，建立變更紀錄"""
        if event.is_directory or not event.src_path.endswith(".xlsx"):
            return
        file_name = os.path.basename(event.src_path)
        if file_name.startswith("~$"):
            return  # 忽略 Excel 產生的暫存檔

        # 避免短時間內重複觸發
        now = time.time()
        last_time = self.last_modified_times.get(file_name, 0)
        if now - last_time < 5:  # 設定 5 秒的防呆機制
            return
        self.last_modified_times[file_name] = now

        time.sleep(2)  # 等待檔案寫入完成
        print(f"偵測到 Excel 變更: {file_name}")

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
