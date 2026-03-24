@echo off
chcp 65001 >nul
title 予定管理システム - 本番更新

echo ============================================
echo   予定管理システム 本番更新
echo ============================================
echo.

REM --- 設定（環境に合わせて変更） ---
set SOURCE=c:\DEV(ClaudCode)\ScheduleManagement
set DEST=C:\Apps\ScheduleManagement

echo 開発元 : %SOURCE%
echo 本番先 : %DEST%
echo.
set TASK_NAME=ScheduleServer

echo ※ DB・SECRET_KEY・ログは上書きしません
echo ※ サーバーを自動で停止→更新→再起動します
echo.
pause

REM --- サーバー停止 ---
echo [1/4] サーバーを停止中...
schtasks /End /TN "%TASK_NAME%" >nul 2>&1
timeout /t 3 /nobreak >nul

REM --- ソースコード同期（DB・設定・Git除外） ---
echo [2/4] ファイルを同期中...
robocopy "%SOURCE%\web_app" "%DEST%\web_app" /E /PURGE /XD __pycache__ >nul
robocopy "%SOURCE%\data" "%DEST%\data" /E >nul
robocopy "%SOURCE%\deploy" "%DEST%\deploy" /E >nul
copy /Y "%SOURCE%\run_web.py" "%DEST%\run_web.py" >nul
copy /Y "%SOURCE%\run_production.py" "%DEST%\run_production.py" >nul
copy /Y "%SOURCE%\requirements_web.txt" "%DEST%\requirements_web.txt" >nul

REM --- 依存パッケージ更新（新しいものがあれば） ---
echo [3/4] パッケージ確認中...
cd /d "%DEST%"
pip install -r requirements_web.txt --quiet

REM --- サーバー再起動 ---
echo [4/4] サーバーを再起動中...
schtasks /Run /TN "%TASK_NAME%"

echo.
echo ============================================
echo   更新完了 - サーバー再起動済み
echo ============================================
echo.
pause
