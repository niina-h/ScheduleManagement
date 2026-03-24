@echo off
chcp 65001 >nul
title 予定管理システム - タスクスケジューラ登録

echo ============================================
echo   予定管理システム 自動起動登録
echo   （タスクスケジューラ方式）
echo ============================================
echo.
echo ※ 管理者として実行してください
echo.

REM --- 管理者権限チェック ---
net session >nul 2>&1
if errorlevel 1 (
    echo [エラー] 管理者権限が必要です。
    echo   このバッチファイルを右クリック →「管理者として実行」してください。
    pause
    exit /b 1
)

REM --- 設定 ---
set APP_DIR=C:\Apps\ScheduleManagement
set TASK_NAME=ScheduleServer

REM --- Python のフルパスを取得 ---
for /f "delims=" %%i in ('where python') do set PYTHON_PATH=%%i
echo Python: %PYTHON_PATH%
echo アプリ: %APP_DIR%
echo.

REM --- 既存タスクがあれば削除 ---
schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1

REM --- タスク登録（PC起動時に自動実行） ---
schtasks /Create /TN "%TASK_NAME%" /TR "\"%PYTHON_PATH%\" \"%APP_DIR%\run_production.py\"" /SC ONSTART /RU SYSTEM /RL HIGHEST /F

if errorlevel 1 (
    echo [エラー] タスク登録に失敗しました。
    pause
    exit /b 1
)

REM --- 今すぐ起動 ---
schtasks /Run /TN "%TASK_NAME%"

echo.
echo ============================================
echo   タスクスケジューラ登録完了
echo ============================================
echo.
echo タスク名 : %TASK_NAME%
echo 起動条件 : PC起動時（SYSTEM権限で自動実行）
echo.
echo 確認   : schtasks /Query /TN "%TASK_NAME%"
echo 停止   : schtasks /End /TN "%TASK_NAME%"
echo 削除   : schtasks /Delete /TN "%TASK_NAME%" /F
echo.
echo PC再起動後も自動的にサーバーが起動します。
echo.
pause
