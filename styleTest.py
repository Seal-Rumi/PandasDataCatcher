import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

def center_window(window, width, height):
    """讓 Tkinter 視窗置中"""
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = int((screen_width / 2) - (width / 2))
    y = int((screen_height / 2) - (height / 2))
    window.geometry(f"{width}x{height}+{x}+{y}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Test")

    # 呼叫置中函式
    center_window(root, 800, 600)

    root.mainloop()
