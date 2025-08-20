import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# 讀取對照表 (固定 ./data/tantof.txt)，並清掉分隔線那一列
def load_mapping():
    try:
        mapping_df = pd.read_csv("./data/tantof.txt", sep="|", engine="python", header=0)
        mapping_df = mapping_df.dropna(axis=1, how="all")                # 移除全空欄
        mapping_df = mapping_df.rename(columns=lambda x: str(x).strip()) # 欄名去空白
        # 去除「-----」分隔線列
        for col in mapping_df.columns:
            mapping_df[col] = mapping_df[col].astype(str).str.strip()
        dash_row_mask = mapping_df.apply(lambda r: r.astype(str).str.fullmatch(r"-+").any(), axis=1)
        mapping_df = mapping_df[~dash_row_mask].copy()

        if not {"TNAME", "TCODE"}.issubset(mapping_df.columns):
            raise ValueError("對照表缺少必要欄位 TNAME 或 TCODE")

        # 建 TNAME → TCODE 對照
        mapping_dict = dict(zip(mapping_df["TNAME"].astype(str).str.strip(),
                                mapping_df["TCODE"].astype(str).str.strip()))
        return mapping_dict
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取對照表失敗: {e}")
        return {}

def make_unique(cols):
    """避免重複欄名：a, a -> a, a.1, a.2 ..."""
    seen = {}
    out = []
    for c in cols:
        c = ("" if pd.isna(c) else str(c)).strip()
        if c == "":
            c = "Unnamed"
        if c in seen:
            seen[c] += 1
            out.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out

# 全域狀態
excel_file = None
df = None
mapping_dict = load_mapping()

def choose_excel():
    """選擇 Excel 檔案，每次重置狀態，並自動載入第一個工作表"""
    global excel_file, df
    excel_file = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not excel_file:
        return
    try:
        # reset 狀態
        df = None
        sheet_combo.set("")
        sheet_combo["values"] = []
        col_combo.set("")
        col_combo["values"] = []

        xl = pd.ExcelFile(excel_file)
        sheet_combo["values"] = xl.sheet_names
        if xl.sheet_names:
            sheet_combo.current(0)
            # 自動載入第一個工作表
            load_sheet()
        else:
            messagebox.showwarning("警告", "此 Excel 沒有工作表")
    except Exception as e:
        messagebox.showerror("錯誤", str(e))

def load_sheet(event=None):
    """讀取選定工作表，強制使用第一列作為欄位標頭"""
    global df
    if not excel_file:
        messagebox.showerror("錯誤", "請先選擇 Excel 檔案")
        return
    sheet_name = sheet_combo.get()
    if not sheet_name:
        messagebox.showerror("錯誤", "請選擇工作表")
        return
    try:
        raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        if raw.shape[0] == 0:
            messagebox.showerror("錯誤", "此工作表沒有資料")
            return

        header = raw.iloc[0].astype(str).str.strip().tolist()
        header = make_unique(header)              # 保障欄名唯一
        df = raw.iloc[1:].reset_index(drop=True)  # 資料從第二列開始
        df.columns = header

        col_combo["values"] = df.columns.tolist()
        if len(df.columns) > 0:
            col_combo.current(0)
        messagebox.showinfo("完成", f"已載入工作表：{sheet_name}\n(第一列已作為欄位標頭)")
    except Exception as e:
        messagebox.showerror("錯誤", f"載入工作表失敗：{e}")

def export_file():
    global df, mapping_dict
    if df is None:
        messagebox.showerror("錯誤", "請先選擇並載入工作表")
        return

    sel_col = col_combo.get()
    if not sel_col:
        messagebox.showerror("錯誤", "請選擇對應欄位")
        return

    if not mapping_dict:
        messagebox.showerror("錯誤", "對照表讀取失敗，無法對應")
        return

    # 以選定欄位的值去對應 TCODE
    temp_series = df[sel_col].astype(str).str.strip()
    tcode_series = temp_series.map(mapping_dict)

    out_df = df.copy()
    out_df["TCODE"] = tcode_series.fillna("無對應")

    # 去除「無對應」的版本
    no_missing_df = out_df[out_df["TCODE"] != "無對應"].copy()

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="另存新檔"
    )
    if save_path:
        try:
            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                # 原始輸出
                out_df.to_excel(writer, index=False, sheet_name="Result")
                # 去除無對應的版本
                no_missing_df.to_excel(writer, index=False, sheet_name="Result_NoMissing")

                # 設定格式
                workbook  = writer.book
                worksheet = writer.sheets["Result"]

                red_format = workbook.add_format({"font_color": "red", "bold": True})

                tcode_col_idx = out_df.columns.get_loc("TCODE")

                worksheet.conditional_format(
                    1, tcode_col_idx, len(out_df), tcode_col_idx,
                    {
                        "type": "cell",
                        "criteria": "==",
                        "value": '"無對應"',
                        "format": red_format,
                    }
                )

            missing_count = (out_df["TCODE"] == "無對應").sum()
            messagebox.showinfo("完成", f"已產生新檔案：\n{save_path}\n\n⚠️ 無對應筆數：{missing_count}")

        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗：{e}")

# 介面
root = tk.Tk()
root.title("Excel 對應轉換工具")
root.geometry("430x300")

frm = ttk.Frame(root, padding=10)
frm.pack(fill="both", expand=True)

ttk.Button(frm, text="選擇 Excel 檔案", command=choose_excel).pack(pady=8)

ttk.Label(frm, text="選擇工作表").pack()
sheet_combo = ttk.Combobox(frm, state="readonly")
sheet_combo.pack(pady=5, fill="x")
sheet_combo.bind("<<ComboboxSelected>>", load_sheet)

ttk.Label(frm, text="選擇對應欄位（將用此欄的值對應到 TCODE）").pack()
col_combo = ttk.Combobox(frm, state="readonly")
col_combo.pack(pady=5, fill="x")

ttk.Button(frm, text="輸出 Excel", command=export_file).pack(pady=18)

root.mainloop()
