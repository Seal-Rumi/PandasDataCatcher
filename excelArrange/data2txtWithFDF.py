import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class ExcelViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sheet Viewer v.1.1.0       Author: SealRumi")
        self.root.geometry("1200x700")

        self.file_path = None
        self.excel_file = None
        self.fdf_path = None
        self.sheet_vars = {}
        self.column_vars = {}
        self.fdf_fields = []

        # === 按鈕區 ===
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", pady=5)

        self.btn_run = tk.Button(btn_frame, text="生成文字檔", command=self.run)
        self.btn_run.pack(side="left", padx=5)

        file_frame = tk.Frame(btn_frame)
        file_frame.pack(side="left", padx=5)

        self.btn_open = tk.Button(file_frame, text="選擇 Excel 檔", command=self.open_file)
        self.btn_open.pack(side="left")

        self.lbl_filename = tk.Label(file_frame, text="（未選擇檔案）", anchor="w")
        self.lbl_filename.pack(side="left", padx=10)

        fdf_frame = tk.Frame(btn_frame)
        fdf_frame.pack(side="left", padx=5)

        self.btn_open_fdf = tk.Button(fdf_frame, text="載入 FDF", command=self.open_fdf)
        self.btn_open_fdf.pack(side="left")

        self.lbl_fdfname = tk.Label(fdf_frame, text="（未載入 FDF）", anchor="w")
        self.lbl_fdfname.pack(side="left", padx=10)

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

        filename = self.file_path.split("/")[-1]
        self.lbl_filename.config(text=filename)

        self.sheet_vars.clear()
        self.column_vars.clear()
        for widget in self.sheet_frame.winfo_children():
            widget.destroy()
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

        self.excel_file = pd.ExcelFile(self.file_path)

        for i, sheet_name in enumerate(self.excel_file.sheet_names):
            row_frame = tk.Frame(self.sheet_frame)
            row_frame.pack(fill="x", pady=2, anchor="w")

            var = tk.BooleanVar()
            cb = tk.Checkbutton(
                row_frame,
                text=f"{i}: {sheet_name}",
                variable=var,
                command=lambda s=sheet_name: self.toggle_sheet_columns(s)
            )
            cb.pack(side="left", padx=5)
            self.sheet_vars[sheet_name] = var

            btn_preview = tk.Button(
                row_frame, text="預覽",
                command=lambda s=sheet_name: self.show_preview(s)
            )
            btn_preview.pack(side="left", padx=2)

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
                    state="disabled",
                    disabledforeground="gray"
                )
                cb_col.pack(side="left", padx=2)
                self.column_vars[sheet_name]["vars"][j] = var_col
                self.column_vars[sheet_name]["widgets"][j] = cb_col

    def open_fdf(self):
        self.fdf_path = filedialog.askopenfilename(filetypes=[("FDF files", "*.fdf")])
        if not self.fdf_path:
            return

        filename = self.fdf_path.split("/")[-1]
        self.lbl_fdfname.config(text=filename)

        # === Reset ===
        self.fdf_fields.clear()

        with open(self.fdf_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        field = {}
        for line in lines:
            line = line.strip()
            if line.startswith("[F"):
                if field:
                    self.fdf_fields.append(field)
                field = {}
            elif "=" in line:
                k, v = line.split("=", 1)
                if k == "Length":
                    field["Length"] = int(v)
                elif k == "Name":
                    field["Name"] = v
                elif k == "Type":
                    field["Type"] = int(v)
        if field:
            self.fdf_fields.append(field)

        self.preview_fdf()

    def preview_fdf(self):
        win = tk.Toplevel(self.root)
        win.title("FDF 欄位預覽")
        win.geometry("400x300")

        tree = ttk.Treeview(win, columns=("Name", "Length", "Type"), show="headings")
        tree.heading("Name", text="Name")
        tree.heading("Length", text="Length")
        tree.heading("Type", text="Type")

        for field in self.fdf_fields:
            tree.insert("", "end", values=(field["Name"], field["Length"], field["Type"]))

        tree.pack(fill="both", expand=True)

    def toggle_sheet_columns(self, sheet_name):
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
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0)

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        for _, row in df.head(100).iterrows():
            self.tree.insert("", "end", values=list(row))

    def run(self):
        if not self.file_path:
            messagebox.showwarning("警告", "請先選擇 Excel 檔案")
            return
        if not self.fdf_fields:
            messagebox.showwarning("警告", "請先載入 FDF")
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

            df_filtered = df.iloc[:, selected_indexes].fillna("")

            for _, row in df_filtered.iterrows():
                line = ""
                for value, field in zip(row, self.fdf_fields):
                    s = str(value)
                    length = field["Length"]
                    if field["Type"] == 1:  # 字串 → 左對齊，補空格
                        s = s.ljust(length)[:length]
                    else:  # 數字 → 右對齊，補0
                        s = s.replace(".0", "")  # 移除浮點尾巴
                        s = s.rjust(length, "0")[:length]
                    line += s
                all_texts.append(line)

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