@echo off
REM =========================================================================
REM Batch File: Sync_fri.bat
REM Description : Runs the Sync_fri.py script to synchronize the FRI folders.
REM Author      : Akshay Solanki
REM Created On  : 19-Oct-2025
REM Usage       : Double-click this file to sync directories automatically.
REM Dependencies: Python must be installed and added to PATH.
REM =========================================================================

echo.
echo 📂 Running Sync_fri.py...
echo.

REM Execute the Python synchronization script
python sync_fri.py

echo.
echo ✅ Sync task completed. Press any key to exit.
pause > nul
