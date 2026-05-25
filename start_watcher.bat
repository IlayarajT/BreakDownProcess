@echo off
:: =============================================================================
::  start_watcher.bat  —  Start BreakDown Watcher (HOTFOLDER mode)
::
::  PLACE THIS FILE at:  V:\TOOLS\BreakDown\start_watcher.bat
::
::  HOW TO RUN (on the Watcher Machine):
::    1. Double-click this file  OR  open CMD and run it.
::    2. Keep the CMD window OPEN — closing it stops the watcher.
::    3. Press Ctrl+C to stop the watcher gracefully.
::
::  WHAT IT DOES:
::    Launches watcher.exe in HOTFOLDER mode.
::    The watcher polls V:\FOR_BREAKDOWN\INPUT\SAGE\ every 10 seconds.
::    When a new package (.zip/.tar/.rar/.7z) is found it calls:
::        BreakDown.exe -p=mAnalyzer -f="<file>" -c="SAGE"
::    and waits up to 900 seconds for it to complete.
::
::  REQUIREMENTS:
::    - V: drive mapped to \\192.168.0.102\d$\REPOSITORY
::    - Java JRE 17+ in PATH  (for JAR invocations)
::    - Microsoft Word installed (for mNormalizer COM automation)
:: =============================================================================

setlocal EnableDelayedExpansion

:: ── Resolve this script's folder as the tool root ────────────────────────────
set "TOOL_DIR=%~dp0"
if "!TOOL_DIR:~-1!"=="\" set "TOOL_DIR=!TOOL_DIR:~0,-1!"

set "WATCHER_EXE=!TOOL_DIR!\watcher.exe"
set "LOG_FILE=!TOOL_DIR!\watcher_startup.log"

:: ── Banner ────────────────────────────────────────────────────────────────────
echo.
echo ============================================================
echo   BreakDown Watcher  ^|  HOTFOLDER mode
echo ============================================================
echo   Tool root : !TOOL_DIR!
echo   Exe       : !WATCHER_EXE!
echo   Location  : HOTFOLDER
echo   Input     : V:\FOR_BREAKDOWN\INPUT\SAGE
echo   Started   : %DATE% %TIME%
echo ============================================================
echo.
echo   Press Ctrl+C to stop the watcher.
echo   Do NOT close this window while processing is active.
echo.

:: ── Validate V: drive ────────────────────────────────────────────────────────
if not exist "V:\" (
    echo [ERROR] V: drive is not mapped.
    echo         Run:  net use V: \\192.168.0.102\d$\REPOSITORY
    echo         Then retry.
    pause
    exit /b 1
)

:: ── Validate watcher.exe ─────────────────────────────────────────────────────
if not exist "!WATCHER_EXE!" (
    echo [ERROR] watcher.exe not found:  !WATCHER_EXE!
    echo         Run install_breakdown.bat to install the tool first.
    pause
    exit /b 1
)

:: ── Validate Java ────────────────────────────────────────────────────────────
java -version >nul 2>&1
if errorlevel 1 (
    echo [WARN] Java not found in PATH.
    echo        JAR-based steps (pre-clean, ParaStyler) will fail.
    echo        Install Java JRE/JDK 17+ and add to PATH.
    echo.
)

:: ── Log startup ──────────────────────────────────────────────────────────────
echo [%DATE% %TIME%] Watcher started (HOTFOLDER mode) >> "!LOG_FILE!"

:: ── Change to tool directory and launch watcher ───────────────────────────────
cd /d "!TOOL_DIR!"

"!WATCHER_EXE!" -p="watcher" -l="HOTFOLDER"

:: ── If we reach here, watcher exited ─────────────────────────────────────────
echo.
echo [INFO] Watcher process has stopped.
echo [%DATE% %TIME%] Watcher stopped. >> "!LOG_FILE!"

pause
endlocal
