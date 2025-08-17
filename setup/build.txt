@echo off
REM ==============================
REM 自動化打包腳本 (Windows)
REM ==============================

call venv\Scripts\activate
echo [1/4] 啟動虛擬環境...
echo [2/4] 安裝需求套件...
pip install --upgrade pip
pip install -r requirements.txt

echo [3/4] 使用 PyInstaller 打包...
pyinstaller --onefile --windowed PandasDataCatcher.py

echo [4/4] 打包完成！
echo 產生的 exe 檔在 dist\ 資料夾裡。
pauseS