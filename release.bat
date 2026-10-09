@echo off
echo === START ===
copy /Y "C:\DEV\ScheduleManagement\web_app\models.py" "C:\App\ScheduleManagement\web_app\models.py"
copy /Y "C:\DEV\ScheduleManagement\web_app\routes\daily.py" "C:\App\ScheduleManagement\web_app\routes\daily.py"
copy /Y "C:\DEV\ScheduleManagement\web_app\routes\planner.py" "C:\App\ScheduleManagement\web_app\routes\planner.py"
copy /Y "C:\DEV\ScheduleManagement\web_app\templates\daily.html" "C:\App\ScheduleManagement\web_app\templates\daily.html"
copy /Y "C:\DEV\ScheduleManagement\web_app\templates\gantt_input_test.html" "C:\App\ScheduleManagement\web_app\templates\gantt_input_test.html"
echo === END ===
pause
