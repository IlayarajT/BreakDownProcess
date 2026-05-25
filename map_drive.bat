@echo off
:: =============================================================================
::  map_drive.bat  —  Map V: drive to the repository server
::
::  Run this ONCE on each machine (Watcher, Processing, ParaStyler).
::  For persistent mapping, run it with administrator rights or add it to
::  Windows startup / login script.
:: =============================================================================

setlocal

set "SERVER_PATH=\\192.168.0.102\d$\REPOSITORY"
set "DRIVE_LETTER=V:"

:: Check if already mapped
if exist "%DRIVE_LETTER%\" (
    echo [INFO] %DRIVE_LETTER% is already mapped.
    net use %DRIVE_LETTER%
    echo.
    echo Press any key to exit ...
    pause >nul
    exit /b 0
)

echo [INFO] Mapping %DRIVE_LETTER% to %SERVER_PATH% ...
net use %DRIVE_LETTER% %SERVER_PATH% /persistent:yes

if errorlevel 1 (
    echo.
    echo [ERROR] Could not map %DRIVE_LETTER%.
    echo         Check:
    echo           1. Server %SERVER_PATH% is reachable
    echo           2. You have permission to access the share
    echo           3. Run this script as Administrator if needed
    pause
    exit /b 1
)

echo.
echo [SUCCESS] %DRIVE_LETTER% mapped to %SERVER_PATH%
echo.
pause
endlocal
