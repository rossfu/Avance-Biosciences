@echo off

REM Check if the third argument (%3) exists
IF "%~3" NEQ "" (
    REM If %3 exists, append "," + %3 to %2
    set "arg2=%2,%3"
) ELSE (
    REM If %3 doesn't exist, use %2 as is
    set "arg2=%2"
)

REM Run the Python script with the modified arguments
python3 -W ignore C:\nifi-1.21.0\scripts\nanodrop_1000.py ""%1"" ""%arg2%""