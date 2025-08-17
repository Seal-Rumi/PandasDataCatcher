@echo off
echo 建立虛擬環境...
python -m venv venv

echo 啟動虛擬環境並安裝需求...
call venv\Scripts\activate
pip install -r requirements.txt

echo 完成！
pause