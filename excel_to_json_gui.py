# -*- coding: utf-8 -*-
import json, sys, os, re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# ---- 欄位名稱候選（可依你表頭調整）----
DRUG_COL_CANDS  = ["藥名成","藥名","藥名/成分","藥名(成)","品名","商品名","藥品名稱","名稱"]
TOTAL_COL_CANDS = ["總耗量","總耗用量","總用量","耗量","合計","總計","Total","total"]

def find_col_name(columns, candidates):
    cols = [str(c).strip() for c in columns]
    # 先做等值
    for cand in candidates:
        if cand in cols: return cand
        # 大小寫/去空白比對
        lowmap = {c.lower().replace(" ", ""): c for c in cols}
        key = cand.lower().replace(" ", "")
        if key in lowmap: return lowmap[key]
    # 再做包含
    for c in cols:
        cl = c.lower().replace(" ", "")
        for cand in candidates:
            if cand.lower().replace(" ", "") in cl:
                return c
    return None

def read_one_sheet(df):
    # 嘗試在這張表找兩個欄
    drug_col  = find_col_name(df.columns, DRUG_COL_CANDS)
    total_col = find_col_name(df.columns, TOTAL_COL_CANDS)
    if not drug_col or not total_col:
        return None

    # 清理與轉型
    df = df[[drug_col, total_col]].copy()
    df[drug_col]  = df[drug_col].astype(str).str.strip()
    # total 轉成數字；字串中的逗點去掉
    df[total_col] = pd.to_numeric(
        df[total_col].astype(str).str.replace(",", ""),
        errors="coerce"
    )

    df = df[df[drug_col].astype(bool) & df[total_col].notna()]
    out = []
    for _, row in df.iterrows():
        name = str(row[drug_col]).strip()
        val  = float(row[total_col])
        # 如果是整數就轉 int；否則保留小數
        total = int(val) if abs(val - int(val)) < 1e-9 else round(val, 4)
        out.append({"drug": name, "total": total})
    return out

def read_excel_to_json(path: Path):
    suffix = path.suffix.lower()
    if suffix == ".xls":
        engine = "xlrd"   # 需 xlrd==1.2.0
    else:
        engine = None     # 由 pandas 自動用 openpyxl 讀 xlsx

    # 逐張工作表尋找
    try:
        with pd.ExcelFile(path, engine=engine) as xf:
            best = None
            best_len = -1
            for sh in xf.sheet_names:
                df = pd.read_excel(xf, sheet_name=sh)
                res = read_one_sheet(df)
                if res and len(res) > best_len:
                    best, best_len = res, len(res)
            if best is None:
                raise ValueError("在任何工作表都找不到【藥名成】與【總耗量】兩欄，請確認表頭。")
            return best
    except ImportError as e:
        if "xlrd" in str(e).lower() and suffix == ".xls":
            raise RuntimeError("需要安裝 xlrd==1.2.0 才能讀 .xls：\n  pip install xlrd==1.2.0")
        if "openpyxl" in str(e).lower() and suffix == ".xlsx":
            raise RuntimeError("需要安裝 openpyxl 才能讀 .xlsx：\n  pip install openpyxl")
        raise
    except Exception:
        raise

def save_json(data, default_path: Path):
    # 讓使用者選儲存位置
    root = tk.Tk(); root.withdraw()
    out_path = filedialog.asksaveasfilename(
        title="儲存 JSON",
        initialfile=default_path.name,
        defaultextension=".json",
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
    )
    root.destroy()
    if not out_path:
        return None
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return Path(out_path)

def main():
    # 選擇 Excel
    root = tk.Tk(); root.withdraw()
    path = filedialog.askopenfilename(
        title="選擇 Excel（.xls/.xlsx）",
        filetypes=[("Excel files","*.xlsx;*.xls"), ("All files","*.*")]
    )
    root.destroy()
    if not path:
        return

    path = Path(path)
    try:
        data = read_excel_to_json(path)
    except Exception as e:
        # 用視窗顯示錯誤，避免 cp950 主控台亂碼
        messagebox.showerror("轉換失敗", str(e))
        return

    default_out = path.with_suffix(".json")
    out_path = save_json(data, default_out)
    if out_path:
        messagebox.showinfo("完成", f"已輸出：{out_path}\n共 {len(data)} 筆")

if __name__ == "__main__":
    # 避免 Windows 主控台編碼問題，整體不印訊息，只用對話框
    try:
        main()
    except Exception as e:
        try:
            messagebox.showerror("錯誤", str(e))
        except:
            pass
        sys.exit(1)
