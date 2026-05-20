@echo off
chcp 932 >nul
title Register ScheduleBackup Task (Production)

REM ============================================
REM Run this on the PRODUCTION server (192.168.70.141).
REM Registers a daily task at 18:00 to run backup_db.bat.
REM ============================================

set TASK_NAME=ScheduleBackup
set SCRIPT=C:\App\ScheduleManagement\bat\backup_db.bat

echo Registering task: %TASK_NAME%
schtasks /Create /TN "%TASK_NAME%" /TR "\"%SCRIPT%\"" /SC DAILY /ST 18:00 /F /RL HIGHEST
if errorlevel 1 (
    echo ERROR: Failed to register task. Run as administrator.
    pause
    exit /b 1
)

echo.
echo OK: Task "%TASK_NAME%" runs daily at 18:00.
echo Confirm: schtasks /Query /TN "%TASK_NAME%"
echo.
pause
