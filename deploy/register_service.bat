@echo off
chcp 65001 >nul
title 予定管理システム - サービス登録

echo ============================================
echo   予定管理システム Windows サービス登録
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
set SERVICE_NAME=ScheduleService
set DISPLAY_NAME=予定管理システム

REM --- Python のフルパスを取得 ---
for /f "delims=" %%i in ('where python') do set PYTHON_PATH=%%i
echo Python: %PYTHON_PATH%
echo アプリ: %APP_DIR%
echo.

REM --- NSSM の存在確認 ---
where nssm >nul 2>&1
if errorlevel 1 (
    echo [情報] nssm が見つかりません。
    echo.
    echo NSSM（Non-Sucking Service Manager）が必要です。
    echo   1. https://nssm.cc/download からダウンロード
    echo   2. nssm.exe を C:\Windows\ にコピー
    echo   3. このバッチを再実行
    echo.
    echo --- または、タスクスケジューラで代用できます ---
    echo   register_taskscheduler.bat を実行してください。
    pause
    exit /b 1
)

REM --- 既存サービスがあれば停止・削除 ---
nssm stop %SERVICE_NAME% >nul 2>&1
nssm remove %SERVICE_NAME% confirm >nul 2>&1

REM --- サービス登録 ---
nssm install %SERVICE_NAME% "%PYTHON_PATH%" "%APP_DIR%\run_production.py"
nssm set %SERVICE_NAME% DisplayName "%DISPLAY_NAME%"
nssm set %SERVICE_NAME% Description "予定管理システム - Webサーバー（waitress）"
nssm set %SERVICE_NAME% AppDirectory "%APP_DIR%"
nssm set %SERVICE_NAME% Start SERVICE_AUTO_START
nssm set %SERVICE_NAME% AppStdout "%APP_DIR%\db\service_stdout.log"
nssm set %SERVICE_NAME% AppStderr "%APP_DIR%\db\service_stderr.log"
nssm set %SERVICE_NAME% AppRotateFiles 1
nssm set %SERVICE_NAME% AppRotateBytes 1048576

REM --- サービス起動 ---
nssm start %SERVICE_NAME%

echo.
echo ============================================
echo   サービス登録完了
echo ============================================
echo.
echo サービス名 : %SERVICE_NAME%
echo 状態確認   : nssm status %SERVICE_NAME%
echo 停止       : nssm stop %SERVICE_NAME%
echo 再起動     : nssm restart %SERVICE_NAME%
echo 削除       : nssm remove %SERVICE_NAME% confirm
echo.
echo PC再起動後も自動的にサーバーが起動します。
echo.
pause
