@echo off
chcp 932 >nul
title Register ScheduleBackupPull Task (Developer PC)

REM ============================================
REM Run this on the DEVELOPER PC.
REM Registers a daily task at 18:30 to pull DB from production.
REM ============================================

set TASK_NAME=ScheduleBackupPull
set SCRIPT=%~dp0pull_backup_from_prod.bat

echo Registering task: %TASK_NAME%
echo Script: %SCRIPT%

schtasks /Create /TN "%TASK_NAME%" /TR "\"%SCRIPT%\"" /SC DAILY /ST 18:30 /F
if errorlevel 1 (
    echo ERROR: Failed to register task.
    pause
    exit /b 1
)

echo.
echo OK: Task "%TASK_NAME%" runs daily at 18:30.
echo Backup target: c:\DEV(ClaudCode)\Backup\ScheduleManagement\
echo Confirm: schtasks /Query /TN "%TASK_NAME%"
echo.
pause
