import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class ExcelViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sheet Viewer v.1.0.0       Author: SealRumi")
        self.root.geometry("1200x700")

        self.file_path = None
        self.excel_file = None
        self.sheet_vars = {}
        self.column_vars = {}

        # === 按鈕區 ===
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", pady=5)

        # 生成文字檔（放左邊）
        self.btn_run = tk.Button(btn_frame, text="生成文字檔", command=self.run)
        self.btn_run.pack(side="left", padx=5)

        # 子容器：選擇 Excel 與檔案名稱
        file_frame = tk.Frame(btn_frame)
        file_frame.pack(side="left", padx=5)

        self.btn_open = tk.Button(file_frame, text="選擇 Excel 檔", command=self.open_file)
        self.btn_open.pack(side="left")

        self.lbl_filename = tk.Label(file_frame, text="（未選擇檔案）", anchor="w")
        self.lbl_filename.pack(side="left", padx=10)

        # === 工作表區 ===
        self.sheet_frame = tk.Frame(root)
        self.sheet_frame.pack(fill="x", pady=5)

        # === 預覽表格 ===
        self.tree = ttk.Treeview(root)
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

    def open_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        if not self.file_path:
            return

        # 顯示檔案名稱（只顯示檔名，不含路徑）
        filename = self.file_path.split("/")[-1]
        self.lbl_filename.config(text=filename)

        # === Reset 狀態 ===
        self.sheet_vars.clear()
        self.column_vars.clear()
        for widget in self.sheet_frame.winfo_children():
            widget.destroy()
        self.tree.delete(*self.tree.get_children())   # 清空預覽表格
        self.tree["columns"] = ()

        # === 載入新檔案 ===
        self.excel_file = pd.ExcelFile(self.file_path)

        for i, sheet_name in enumerate(self.excel_file.sheet_names):
            row_frame = tk.Frame(self.sheet_frame)
            row_frame.pack(fill="x", pady=2, anchor="w")

            # 勾選工作表
            var = tk.BooleanVar()
            cb = tk.Checkbutton(
                row_frame,
                text=f"{i}: {sheet_name}",
                variable=var,
                command=lambda s=sheet_name: self.toggle_sheet_columns(s)
            )
            cb.pack(side="left", padx=5)
            self.sheet_vars[sheet_name] = var

            # 預覽按鈕
            btn_preview = tk.Button(
                row_frame, text="預覽",
                command=lambda s=sheet_name: self.show_preview(s)
            )
            btn_preview.pack(side="left", padx=2)

            # 全選 / 全部取消
            btn_select_all = tk.Button(
                row_frame, text="全選",
                command=lambda s=sheet_name: self.select_all_columns(s)
            )
            btn_select_all.pack(side="left", padx=2)

            btn_deselect_all = tk.Button(
                row_frame, text="全部取消",
                command=lambda s=sheet_name: self.deselect_all_columns(s)
            )
            btn_deselect_all.pack(side="left", padx=2)

            # 欄位勾選區（預設禁用、灰色）
            col_frame = tk.Frame(row_frame)
            col_frame.pack(side="left", padx=10)
            self.column_vars[sheet_name] = {"frame": col_frame, "vars": {}, "widgets": {}}

            df = self.excel_file.parse(sheet_name, header=0, nrows=1)
            for j, col in enumerate(df.columns):
                var_col = tk.BooleanVar(value=False)
                cb_col = tk.Checkbutton(
                    col_frame,
                    text=str(col),
                    variable=var_col,
                    state="disabled",              # ✅ 預設禁用
                    disabledforeground="gray"      # ✅ 顯示灰字
                )
                cb_col.pack(side="left", padx=2)
                self.column_vars[sheet_name]["vars"][j] = var_col
                self.column_vars[sheet_name]["widgets"][j] = cb_col

    def toggle_sheet_columns(self, sheet_name):
        """依照 sheet 的勾選狀態，啟用 / 停用欄位選擇"""
        enabled = self.sheet_vars[sheet_name].get()
        for cb in self.column_vars[sheet_name]["widgets"].values():
            cb.config(state="normal" if enabled else "disabled")

    def select_all_columns(self, sheet_name):
        for var in self.column_vars[sheet_name]["vars"].values():
            var.set(True)

    def deselect_all_columns(self, sheet_name):
        for var in self.column_vars[sheet_name]["vars"].values():
            var.set(False)

    def show_preview(self, sheet_name):
        """預覽整個工作表，不受欄位勾選影響"""
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0)

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        # 顯示前 20 列資料
        for _, row in df.head(100).iterrows():
            self.tree.insert("", "end", values=list(row))

    def run(self):
        if not self.file_path:
            messagebox.showwarning("警告", "請先選擇 Excel 檔案")
            return

        selected_sheets = [s for s, var in self.sheet_vars.items() if var.get()]
        if not selected_sheets:
            messagebox.showwarning("警告", "請至少勾選一個工作表")
            return

        all_texts = []
        for sheet_name in selected_sheets:
            df = pd.read_excel(self.file_path, header=0, sheet_name=sheet_name)

            selected_indexes = [
                i for i, var in self.column_vars[sheet_name]["vars"].items() if var.get()
            ]
            if not selected_indexes:
                continue

            df_filtered = df.iloc[:, selected_indexes].dropna()
            concat_series = df_filtered.fillna("").astype(str).agg("".join, axis=1)
            all_texts.extend(concat_series.tolist())

        if not all_texts:
            messagebox.showwarning("警告", "沒有可輸出的資料")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                for line in all_texts:
                    f.write(line + "\n")
            messagebox.showinfo("完成", f"已輸出到 {save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewer(root)
    root.mainloop()
