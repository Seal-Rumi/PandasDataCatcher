import pandas as pd
from tkinter import Tk, filedialog

# 1. 對話方塊讀取指定 xls / xlsx 檔案
root = Tk()
root.withdraw()  # 不顯示主視窗
file_path = filedialog.askopenfilename(
    title="選擇Excel檔案",
    filetypes=[("Excel files", "*.xls *.xlsx")]
)

if not file_path:
    print("未選擇檔案")
else:
    # 所有工作表
    excel_file = pd.ExcelFile(file_path)
    print("所有工作表：")
    print(excel_file.sheet_names)

    # 指定工作表的第一行有哪些欄位
    sheet_name = input("請輸入要查看的工作表名稱: ")
    if sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"\n工作表 {sheet_name} 的欄位：")
        print(df.columns.tolist())
    else:
        print("輸入的工作表不存在")
