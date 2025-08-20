import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class ExcelCustomizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 客製化工具")
        self.root.geometry("1024x768")

        # 上方 frame (選擇 Excel)
        self.top_frame = tk.Frame(root, height=50, bg="#e0e0e0")
        self.top_frame.pack(side="top", fill="x", padx=10, pady=5)

        tk.Button(self.top_frame, text="選擇 Excel 檔案", command=self.load_file).pack(side="left", padx=5, pady=5)
        self.file_label = tk.Label(self.top_frame, text="尚未選擇檔案", anchor="w")
        self.file_label.pack(side="left", fill="x", expand=True, padx=5, pady=5, anchor="w")

        # 左右兩個 Frame
        self.left_frame = tk.Frame(root, width=400, bg="#dcdcdc")
        self.left_frame.pack(side="left", fill="y")
        self.right_frame = tk.Frame(root, width=400)
        self.right_frame.pack(side="right", fill="both", expand=True)

        # 左側：工作表選擇 + 新增欄位按鈕
        top_controls_frame = tk.Frame(self.left_frame, bg="#dcdcdc")
        top_controls_frame.pack(fill="x", padx=5, pady=5)
        tk.Label(top_controls_frame, text="選擇工作表:", bg="#dcdcdc").pack(side="left")
        self.sheet_option = ttk.Combobox(top_controls_frame, state="readonly")
        self.sheet_option.pack(side="left", padx=5)
        self.sheet_option.bind("<<ComboboxSelected>>", self.update_columns)
        tk.Button(top_controls_frame, text="新增欄位", command=self.add_custom_field).pack(side="left", padx=5)

        # 標題列 frame
        self.header_frame = tk.Frame(self.left_frame, bg="#dcdcdc")
        self.header_frame.pack(fill="x", padx=5, pady=(10,5))
        tk.Label(self.header_frame, text="名稱", width=15, bg="#dcdcdc").pack(side="left")
        tk.Label(self.header_frame, text="抓取資料", width=15, bg="#dcdcdc").pack(side="left", padx=5)
        tk.Label(self.header_frame, text="自定義值", width=15, bg="#dcdcdc").pack(side="left", padx=5)
        tk.Label(self.header_frame, text="刪除", width=5, bg="#dcdcdc").pack(side="left", padx=5)

        # 自訂欄位 container
        self.custom_fields_container = tk.Frame(self.left_frame, bg="#dcdcdc")
        self.custom_fields_container.pack(fill="x")
        self.custom_fields_container.children_list = []

        # 左下方按鈕：預覽 & 輸出
        bottom_buttons_frame = tk.Frame(self.left_frame, bg="#dcdcdc")
        bottom_buttons_frame.pack(side="bottom", fill="x", padx=5, pady=10)
        tk.Button(bottom_buttons_frame, text="預覽資料", command=self.preview_data).pack(side="left", padx=5)
        tk.Button(bottom_buttons_frame, text="輸出資料", command=self.export_data).pack(side="left", padx=5)

        # Excel 物件
        self.excel_file = None
        self.df = None
        self.columns = []

        # 預覽表格
        self.preview_table = None

    def load_file(self):
        file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xls *.xlsx")])
        if not file_path:
            return

        self.file_label.config(text=file_path)

        try:
            self.excel_file = pd.ExcelFile(file_path)
            self.sheet_option['values'] = self.excel_file.sheet_names
            self.sheet_option.set('')
            self.columns = []
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取 Excel 檔案:\n{e}")

    def update_columns(self, event):
        sheet_name = self.sheet_option.get()
        if not sheet_name:
            return
        try:
            self.df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
            self.columns = self.df.columns.tolist()
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取工作表失敗:\n{e}")

    def add_custom_field(self):
        if not self.columns:
            messagebox.showwarning("提醒", "請先選擇工作表")
            return

        field_frame = tk.Frame(self.custom_fields_container, bg="#dcdcdc")
        field_frame.pack(fill="x", padx=5, pady=2)

        # 名稱欄位
        name_entry = tk.Entry(field_frame, width=15)
        name_entry.pack(side="left")

        # 選取器
        options = ["==自定義資料=="] + self.columns
        column_combobox = ttk.Combobox(field_frame, values=options, state="readonly", width=15)
        column_combobox.pack(side="left", padx=5)

        # Value 欄位
        value_entry = tk.Entry(field_frame, width=15)
        value_entry.pack(side="left")
        value_entry.config(state="normal")

        # 刪除按鈕
        delete_btn = tk.Button(field_frame, text="🗑️", command=lambda f=field_frame: self.delete_field(f))
        delete_btn.pack(side="left", padx=5)

        def on_select(event):
            selected = column_combobox.get()
            if selected == "==自定義資料==":
                value_entry.config(state="normal")
            else:
                value_entry.delete(0, tk.END)
                value_entry.config(state="disabled")
            name_entry.delete(0, tk.END)
            name_entry.insert(0, selected)

        column_combobox.bind("<<ComboboxSelected>>", on_select)
        self.custom_fields_container.children_list.append((field_frame, name_entry, column_combobox, value_entry))

    def delete_field(self, frame):
        frame.destroy()
        self.custom_fields_container.children_list = [t for t in self.custom_fields_container.children_list if t[0] != frame]

    def preview_data(self):
        if self.df is None or self.df.empty or not self.custom_fields_container.children_list:
            messagebox.showwarning("提醒", "請先選擇工作表並新增欄位")
            return

        data = {}
        for _, name_entry, column_combobox, value_entry in self.custom_fields_container.children_list:
            col_name = name_entry.get()
            selected = column_combobox.get()
            if selected == "==自定義資料==":
                data[col_name] = [value_entry.get()] * len(self.df)
            else:
                data[col_name] = self.df[selected].tolist()

        preview_df = pd.DataFrame(data)

        # 顯示在右側 frame
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        self.preview_table = ttk.Treeview(self.right_frame)
        self.preview_table.pack(fill="both", expand=True)
        self.preview_table["columns"] = list(preview_df.columns)
        self.preview_table["show"] = "headings"
        for col in preview_df.columns:
            self.preview_table.heading(col, text=col)
        for _, row in preview_df.iterrows():
            self.preview_table.insert("", "end", values=list(row))

        self.current_preview = preview_df

    def export_data(self):
        if not hasattr(self, 'current_preview'):
            messagebox.showwarning("提醒", "請先預覽資料")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            self.current_preview.to_excel(file_path, index=False)
            messagebox.showinfo("完成", f"已儲存 Excel: {file_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCustomizerApp(root)
    root.mainloop()
