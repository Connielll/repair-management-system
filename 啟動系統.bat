@echo off
chcp 65001 >nul
title 維修案件管理系統
echo ========================================================
echo   維修案件管理系統 - 自動啟動器
echo ========================================================
echo.
echo [1/2] 檢查並安裝必要套件...
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 套件安裝失敗，請確認已安裝 Python 且已勾選 "Add python.exe to PATH"。
    pause
    exit /b
)

echo.
echo [2/2] 正在啟動系統畫面，即將開啟瀏覽器...
python -m streamlit run app.py
if %errorlevel% neq 0 (
    echo.
    echo [提示] 若畫面未自動開啟，請手動在瀏覽器輸入 http://localhost:8501
    pause
)
