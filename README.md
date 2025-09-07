# Excel to JSON (GUI)

- 單檔 Python 工具：讀 Excel (.xls/.xlsx) 輸出 `[{drug,total}]` JSON。
- .xls：優先用 xlrd；若不可用，Windows 可走 Excel COM fallback 轉成 .xlsx 再讀。
- 打包：`pyinstaller --onefile --noconsole ...`

## 使用
1. 雙擊 `ExcelToJSON.exe`（或 `python excel_to_json_gui.py`）
2. 選 Excel 檔 → 選輸出路徑 → 產生 JSON。

## 需求
見 `requirements.txt`。
