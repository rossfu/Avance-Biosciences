@echo off

REM Set the log file path
set LOG_FILE=C:\nifi-1.21.0\scripts\log_file.txt

REM Append the command to the log file
echo %date% %time% - python3 -W ignore C:\nifi-1.21.0\scripts\create_shortcut.py %1 %2 >> %LOG_FILE%

REM Run the Python script with the modified arguments
python3 -W ignore C:\nifi-1.21.0\scripts\create_shortcut_full.py %1 %2 %3