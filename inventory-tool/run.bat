@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    py -3 main.py
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python が見つかりません。README.md の準備手順を確認してください。
        pause
        popd
        exit /b 1
    )
    python main.py
)
if errorlevel 1 (
    echo 起動できませんでした。Python と requirements.txt のインストールを確認してください。
    pause
    popd
    exit /b 1
)
popd
endlocal
