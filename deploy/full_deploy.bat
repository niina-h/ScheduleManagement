@echo off
chcp 65001 >nul
title 予定管理システム - フルデプロイ（DB含む）

echo ============================================
echo   予定管理システム フルデプロイ（DB含む）
echo ============================================
echo.

REM --- 設定（環境に合わせて変更） ---
set SOURCE=c:\DEV(ClaudCode)\ScheduleManagement
set DEST=C:\Apps\ScheduleManagement
set TASK_NAME=ScheduleServer

echo 開発元 : %SOURCE%
echo 本番先 : %DEST%
echo.
echo ★★★ 注意 ★★★
echo   本番DBを開発DBで上書きします。
echo   本番の既存データは上書きされます。
echo   （バックアップは自動で作成されます）
echo.
pause

REM --- サーバー停止 ---
echo [1/5] サーバーを停止中...
schtasks /End /TN "%TASK_NAME%" >nul 2>&1
timeout /t 3 /nobreak >nul

REM --- 本番DBバックアップ ---
echo [2/5] 本番DBをバックアップ中...
set BACKUP_SUFFIX=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%
set BACKUP_SUFFIX=%BACKUP_SUFFIX: =0%
if exist "%DEST%\db\web_app.db" (
    copy /Y "%DEST%\db\web_app.db" "%DEST%\db\web_app.db.bak_%BACKUP_SUFFIX%" >nul
    echo   → web_app.db.bak_%BACKUP_SUFFIX% に保存しました
)
if exist "%DEST%\db\schedule.db" (
    copy /Y "%DEST%\db\schedule.db" "%DEST%\db\schedule.db.bak_%BACKUP_SUFFIX%" >nul
    echo   → schedule.db.bak_%BACKUP_SUFFIX% に保存しました
)

REM --- ソースコード同期 ---
echo [3/5] ファイルを同期中...
robocopy "%SOURCE%\web_app" "%DEST%\web_app" /E /PURGE /XD __pycache__ >nul
robocopy "%SOURCE%\data" "%DEST%\data" /E >nul
robocopy "%SOURCE%\deploy" "%DEST%\deploy" /E >nul
copy /Y "%SOURCE%\run_web.py" "%DEST%\run_web.py" >nul
copy /Y "%SOURCE%\run_production.py" "%DEST%\run_production.py" >nul
copy /Y "%SOURCE%\requirements_web.txt" "%DEST%\requirements_web.txt" >nul

REM --- DB上書きコピー ---
echo [4/5] DBを上書きコピー中...
if not exist "%DEST%\db" mkdir "%DEST%\db"
copy /Y "%SOURCE%\db\web_app.db" "%DEST%\db\web_app.db" >nul
echo   → web_app.db をコピーしました
if exist "%SOURCE%\db\schedule.db" (
    copy /Y "%SOURCE%\db\schedule.db" "%DEST%\db\schedule.db" >nul
    echo   → schedule.db をコピーしました
)

REM --- 依存パッケージ更新 ---
echo [5/5] パッケージ確認中...
cd /d "%DEST%"
pip install -r requirements_web.txt --quiet

REM --- サーバー再起動 ---
echo サーバーを再起動中...
schtasks /Run /TN "%TASK_NAME%"

echo.
echo ============================================
echo   フルデプロイ完了 - DB上書き済み
echo   サーバー再起動済み
echo ============================================
echo.
pause
