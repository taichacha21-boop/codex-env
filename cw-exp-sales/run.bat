@echo off
chcp 65001 >nul
setlocal
set "PYTHONUTF8=1"
set "CW_EXP_SALES_PAUSE=1"

python --version >nul 2>&1
if errorlevel 1 (
    echo Pythonが見つかりません。README.md の手順でPythonをインストールしてください。
    if "%~1"=="" pause
    exit /b 1
)

python "%~dp0main.py" %*
set "CW_EXP_SALES_EXIT=%ERRORLEVEL%"
exit /b %CW_EXP_SALES_EXIT%
