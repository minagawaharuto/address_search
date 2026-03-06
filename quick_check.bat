@echo off
echo ====================================
echo Excel ファイル簡易診断ツール
echo ====================================
echo.

if "%~1"=="" (
    echo 使い方: このバッチファイルにExcelファイルをドラッグ&ドロップしてください
    echo.
    pause
    exit /b
)

echo ファイル: %~1
echo.

python check_file.py "%~1"

echo.
echo ====================================
pause
