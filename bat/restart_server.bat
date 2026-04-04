@echo off
chcp 65001 >nul
REM --- pythonw プロセスを強制終了 ---
taskkill /F /IM pythonw.exe >nul 2>&1
timeout /t 3 /nobreak >nul
REM --- サーバー起動 ---
cd /d C:\App\ScheduleManagement
start "" C:\Python\Python313\pythonw.exe run_production.py
