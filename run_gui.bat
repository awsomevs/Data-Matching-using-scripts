@echo off
title Launching Multi-File Data Comparator GUI...
echo Starting Data Comparator Desktop Application...

if exist "venv\Scripts\python.exe" (
    echo Using virtual environment...
    venv\Scripts\python.exe app.py
) else (
    echo Using system Python...
    python app.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with an error. Press any key to close window...
    pause > nul
)

