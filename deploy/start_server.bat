@echo off
chcp 65001 >nul
title 予定管理システム - 本番サーバー

echo ============================================
echo   予定管理システム 本番サーバー起動
echo ============================================
echo.

REM --- 設定 ---
set PORT=5000
set THREADS=4

REM --- アプリのルートディレクトリへ移動 ---
cd /d "%~dp0.."

REM --- Python の存在確認 ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [エラー] Python が見つかりません。Python 3.9 以上をインストールしてください。
    pause
    exit /b 1
)

REM --- waitress の存在確認 ---
python -c "import waitress" >nul 2>&1
if errorlevel 1 (
    echo [エラー] waitress がインストールされていません。
    echo   pip install -r requirements_web.txt を実行してください。
    pause
    exit /b 1
)

echo ポート    : %PORT%
echo スレッド  : %THREADS%
echo 停止      : Ctrl+C または このウインドウを閉じる
echo.
echo サーバーを起動しています...
echo.

python run_production.py --port %PORT% --threads %THREADS%

pause
