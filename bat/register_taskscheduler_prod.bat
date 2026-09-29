@echo off
chcp 65001 >nul
title 予定管理システム - タスクスケジューラ再登録（本番実配置版）

echo ============================================
echo   予定管理システム タスクスケジューラ再登録
echo   （本番実配置 C:\DEV\ScheduleManagement 版）
echo   ユーザーログイン不要 / SYSTEM 権限で自動起動
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

REM --- 設定（本番の実配置に合わせる） ---
set APP_DIR=C:\DEV\ScheduleManagement
set TASK_NAME=ScheduleServer

REM --- Python のフルパスを取得（システム側 Python 3.13 を優先） ---
if exist "C:\Python\Python313\python.exe" (
    set PYTHON_PATH=C:\Python\Python313\python.exe
) else (
    for /f "delims=" %%i in ('where python') do set PYTHON_PATH=%%i
)

echo Python : %PYTHON_PATH%
echo アプリ : %APP_DIR%
echo タスク : %TASK_NAME%
echo.

REM --- 実行対象ファイルが存在するか確認 ---
if not exist "%APP_DIR%\run_production.py" (
    echo [エラー] %APP_DIR%\run_production.py が見つかりません。
    pause
    exit /b 1
)

REM --- 既存タスクがあれば削除 ---
schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1

REM --- タスク登録（PC起動時 / SYSTEM / 最高権限） ---
schtasks /Create /TN "%TASK_NAME%" /TR "\"%PYTHON_PATH%\" \"%APP_DIR%\run_production.py\"" /SC ONSTART /RU SYSTEM /RL HIGHEST /F

if errorlevel 1 (
    echo [エラー] タスク登録に失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   登録完了
echo ============================================
echo.
echo タスク名   : %TASK_NAME%
echo 実行ユーザー : SYSTEM （ユーザーログイン不要）
echo 起動条件   : PC起動時（自動）
echo 実行対象   : %PYTHON_PATH% %APP_DIR%\run_production.py
echo.
echo ★ PC再起動後、niina-h がログインしていなくてもサーバーが起動します。
echo.
echo 確認コマンド : schtasks /Query /TN "%TASK_NAME%" /V /FO LIST
echo 手動起動    : schtasks /Run /TN "%TASK_NAME%"
echo 手動停止    : schtasks /End /TN "%TASK_NAME%"
echo.
pause
