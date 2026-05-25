@echo off
setlocal EnableDelayedExpansion

:: =============================================================================
::  install_breakdown.bat  —  BreakDown Tool Installation Script
::
::  USAGE:
::      install_breakdown.bat                  (installs to V:\TOOLS\BreakDown)
::      install_breakdown.bat  "D:\CustomPath" (installs to custom location)
::
::  PRE-REQUISITES:
::      1. V: drive mapped to \\192.168.0.102\d$\REPOSITORY (or run map_drive.bat)
::      2. Java JRE/JDK 17+ installed and  java  in PATH
::      3. Run from the project/release root folder
::         (folder containing BreakDown.exe or dist\BreakDown_v*\)
::
::  WHAT THIS SCRIPT DOES:
::      Step 1   Validate environment (V: drive, Java)
::      Step 2   Locate the built EXE (dist folder or current folder)
::      Step 3   Create target folder structure on V:\TOOLS\BreakDown
::      Step 4   Copy BreakDown.exe + watcher.exe
::      Step 5   Copy _internal (PyInstaller runtime)
::      Step 6   Copy JAR files (DocxManipulator, ParaStyler, aspose-words)
::      Step 7   Copy Saxon licence files
::      Step 8   Copy config YAML / JSON files
::      Step 9   Copy SupportingFiles (templates, checkDocRunning, etc.)
::      Step 10  Copy XSL stylesheets
::      Step 11  Write / update startupConfig.yaml
::      Step 12  Create working folders under V:\FOR_BREAKDOWN
::      Step 13  Create start_watcher.bat + start_watcher_s3.bat in install dir
::      Step 14  Print summary
:: =============================================================================

echo.
echo =============================================================
echo   BreakDown  ^|  Installation Script
echo =============================================================
echo.

:: ------------------------------------------------------------
::  Resolve install root  (default: V:\TOOLS\BreakDown)
:: ------------------------------------------------------------
set "INSTALL_ROOT=V:\TOOLS\BreakDown"
if not "%~1"=="" set "INSTALL_ROOT=%~1"

echo [INFO] Install target : %INSTALL_ROOT%
echo.

:: ============================================================
::  STEP 1  —  Validate environment
:: ============================================================
echo [step 1/14]  Validating environment ...

:: --- Check V: drive is mapped
if not exist "V:\" (
    echo [ERROR] V: drive is not mapped.
    echo         Run:  net use V: \\192.168.0.102\d$\REPOSITORY
    echo         Or double-click map_drive.bat and retry.
    exit /b 1
)
echo             V: drive  OK

:: --- Check Java
java -version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Java not found in PATH.
    echo         Install Java JRE/JDK 17+ and add it to the PATH.
    exit /b 1
)
for /f "tokens=3" %%v in ('java -version 2^>^&1 ^| findstr /i "version"') do (
    echo             Java version : %%v
)

echo.

:: ============================================================
::  STEP 2  —  Locate the built EXE / release folder
:: ============================================================
echo [step 2/14]  Locating built EXE ...

set "SRC_ROOT=%~dp0"
:: Remove trailing backslash
if "!SRC_ROOT:~-1!"=="\" set "SRC_ROOT=!SRC_ROOT:~0,-1!"

:: Check if BreakDown.exe is already here (running from release folder)
set "BREAKDOWN_EXE=!SRC_ROOT!\BreakDown.exe"
set "WATCHER_EXE=!SRC_ROOT!\watcher.exe"
set "INTERNAL_DIR=!SRC_ROOT!\_internal"

if not exist "!BREAKDOWN_EXE!" (
    :: Try dist\BreakDown_v*\ folders
    for /d %%d in ("!SRC_ROOT!\dist\BreakDown_v*") do (
        if exist "%%d\BreakDown.exe" (
            set "SRC_ROOT=%%d"
            set "BREAKDOWN_EXE=%%d\BreakDown.exe"
            set "WATCHER_EXE=%%d\watcher.exe"
            set "INTERNAL_DIR=%%d\_internal"
        )
    )
)

if not exist "!BREAKDOWN_EXE!" (
    echo [ERROR] BreakDown.exe not found.
    echo         Run  build.bat  first, or run this script from the release folder.
    exit /b 1
)

echo             BreakDown.exe : !BREAKDOWN_EXE!
echo             watcher.exe   : !WATCHER_EXE!
echo             _internal     : !INTERNAL_DIR!
echo.

:: ============================================================
::  STEP 3  —  Create folder structure
:: ============================================================
echo [step 3/14]  Creating folder structure ...

:: Main installation folder
if not exist "!INSTALL_ROOT!" mkdir "!INSTALL_ROOT!"

:: Sub-folders inside install root
for %%f in (
    "config"
    "SupportingFiles"
    "DocxManipulator"
    "DocxManipulator\jar"
    "ParaStyler"
    "aspose-words"
    "aspose-words\jar"
    "xsl"
) do (
    if not exist "!INSTALL_ROOT!\%%~f" mkdir "!INSTALL_ROOT!\%%~f"
)

:: Working data folders on V: drive
for %%f in (
    "FOR_BREAKDOWN"
    "FOR_BREAKDOWN\INPUT"
    "FOR_BREAKDOWN\INPUT\SAGE"
    "FOR_BREAKDOWN\INPUT\XYZ"
    "FOR_BREAKDOWN\PROCESS"
    "FOR_BREAKDOWN\ERROR"
    "FOR_BREAKDOWN\LOG"
    "FOR_BREAKDOWN\MERGER_INPUT"
    "FOR_BREAKDOWN\MERGER_INPUT\SAGE"
    "FOR_BREAKDOWN\MERGER_INPUT\XYZ"
    "FOR_BREAKDOWN\MERGER_ERROR"
    "FOR_BREAKDOWN\MERGER_ERROR\SAGE"
    "FOR_BREAKDOWN\MERGER_ERROR\XYZ"
    "FOR_BREAKDOWN\ParaStyler_INPUT"
    "FOR_BREAKDOWN\ParaStyler_INPUT\SAGE"
    "FOR_BREAKDOWN\ParaStyler_INPUT\XYZ"
    "FOR_BREAKDOWN\ParaStyler_ERROR"
    "FOR_BREAKDOWN\ParaStyler_ERROR\SAGE"
    "FOR_BREAKDOWN\ParaStyler_ERROR\XYZ"
    "FOR_BREAKDOWN\BreakDown_INPUT"
    "FOR_BREAKDOWN\BreakDown_INPUT\SAGE"
    "FOR_BREAKDOWN\BreakDown_INPUT\XYZ"
    "FOR_BREAKDOWN\BreakDown_DONE"
    "FOR_BREAKDOWN\BreakDown_DONE\SAGE"
    "FOR_BREAKDOWN\BreakDown_DONE\XYZ"
    "FOR_BREAKDOWN\BreakDown_ERROR"
    "FOR_BREAKDOWN\BreakDown_ERROR\SAGE"
    "FOR_BREAKDOWN\BreakDown_ERROR\XYZ"
    "FOR_CONVERSION"
    "FOR_CONVERSION\SAGE"
    "FOR_CONVERSION\XYZ"
) do (
    if not exist "V:\%%~f" mkdir "V:\%%~f"
)

echo             Folders created.
echo.

:: ============================================================
::  STEP 4  —  Copy BreakDown.exe + watcher.exe
:: ============================================================
echo [step 4/14]  Copying executables ...

copy /y "!BREAKDOWN_EXE!" "!INSTALL_ROOT!\BreakDown.exe" >nul
if errorlevel 1 ( echo [WARN] Could not copy BreakDown.exe )

if exist "!WATCHER_EXE!" (
    copy /y "!WATCHER_EXE!" "!INSTALL_ROOT!\watcher.exe" >nul
    if errorlevel 1 ( echo [WARN] Could not copy watcher.exe )
) else (
    echo [WARN] watcher.exe not found at !WATCHER_EXE! — skipping.
)

echo             Executables copied.
echo.

:: ============================================================
::  STEP 5  —  Copy _internal (PyInstaller runtime)
:: ============================================================
echo [step 5/14]  Copying _internal runtime ...

if exist "!INTERNAL_DIR!" (
    xcopy /e /i /y /q "!INTERNAL_DIR!" "!INSTALL_ROOT!\_internal" >nul
    echo             _internal copied.
) else (
    echo [INFO] No _internal folder found (single-file EXE build or source run).
)
echo.

:: ============================================================
::  STEP 6  —  Copy JAR files
:: ============================================================
echo [step 6/14]  Copying JAR files ...

:: DocxManipulator JARs (from source tree)
if exist "!SRC_ROOT!\DocxManipulator\docx-manipulator.jar" (
    copy /y "!SRC_ROOT!\DocxManipulator\docx-manipulator.jar" "!INSTALL_ROOT!\DocxManipulator\docx-manipulator.jar" >nul
    copy /y "!SRC_ROOT!\DocxManipulator\docx-manipulator.jar" "!INSTALL_ROOT!\DocxManipulator\sage-auto-styler.jar" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\docx-manipulator.jar" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\docx-manipulator.jar" "!INSTALL_ROOT!\DocxManipulator\jar\docx-manipulator.jar" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\docx-bookmark.jar" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\docx-bookmark.jar" "!INSTALL_ROOT!\DocxManipulator\jar\docx-bookmark.jar" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\docx-manipulator_cpo_as_client.jar" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\docx-manipulator_cpo_as_client.jar" "!INSTALL_ROOT!\DocxManipulator\jar\" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\newASjid.yml" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\newASjid.yml" "!INSTALL_ROOT!\DocxManipulator\jar\newASjid.yml" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\newASjid.config" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\newASjid.config" "!INSTALL_ROOT!\DocxManipulator\jar\newASjid.config" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\config.properties" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\config.properties" "!INSTALL_ROOT!\DocxManipulator\jar\config.properties" >nul
)
if exist "!SRC_ROOT!\DocxManipulator\jar\Config.yml" (
    copy /y "!SRC_ROOT!\DocxManipulator\jar\Config.yml" "!INSTALL_ROOT!\DocxManipulator\jar\Config.yml" >nul
)

:: Aspose-Words dependency JARs
if exist "!SRC_ROOT!\aspose-words\jar" (
    xcopy /e /i /y /q "!SRC_ROOT!\aspose-words\jar" "!INSTALL_ROOT!\aspose-words\jar" >nul
)
if exist "!SRC_ROOT!\aspose-words\aspose-words.jar" (
    copy /y "!SRC_ROOT!\aspose-words\aspose-words.jar" "!INSTALL_ROOT!\aspose-words\aspose-words.jar" >nul
)

:: ParaStyler JARs + model
if exist "!SRC_ROOT!\ParaStyler\saxon9pe.jar" (
    copy /y "!SRC_ROOT!\ParaStyler\saxon9pe.jar" "!INSTALL_ROOT!\ParaStyler\saxon9pe.jar" >nul
)
if exist "!SRC_ROOT!\ParaStyler\weka-stable-3.6.6.jar" (
    copy /y "!SRC_ROOT!\ParaStyler\weka-stable-3.6.6.jar" "!INSTALL_ROOT!\ParaStyler\weka-stable-3.6.6.jar" >nul
)
if exist "!SRC_ROOT!\ParaStyler\asprop30x.arff.randomCommitee_50.model" (
    copy /y "!SRC_ROOT!\ParaStyler\asprop30x.arff.randomCommitee_50.model" "!INSTALL_ROOT!\ParaStyler\" >nul
)
if exist "!SRC_ROOT!\ParaStyler\commons-cli-1.2.jar" (
    copy /y "!SRC_ROOT!\ParaStyler\commons-cli-1.2.jar"   "!INSTALL_ROOT!\ParaStyler\" >nul
    copy /y "!SRC_ROOT!\ParaStyler\commons-io-2.4.jar"    "!INSTALL_ROOT!\ParaStyler\" >nul
    copy /y "!SRC_ROOT!\ParaStyler\guava-10.0.jar"        "!INSTALL_ROOT!\ParaStyler\" >nul
    copy /y "!SRC_ROOT!\ParaStyler\jsr305-1.3.9.jar"      "!INSTALL_ROOT!\ParaStyler\" >nul
    copy /y "!SRC_ROOT!\ParaStyler\log4j-api-2.0-beta8.jar"  "!INSTALL_ROOT!\ParaStyler\" >nul
    copy /y "!SRC_ROOT!\ParaStyler\log4j-core-2.0-beta8.jar" "!INSTALL_ROOT!\ParaStyler\" >nul
)
if exist "!SRC_ROOT!\ParaStyler\run.bat" (
    copy /y "!SRC_ROOT!\ParaStyler\run.bat" "!INSTALL_ROOT!\ParaStyler\run.bat" >nul
)

echo             JAR files copied.
echo.

:: ============================================================
::  STEP 7  —  Copy licence files
:: ============================================================
echo [step 7/14]  Copying licence files ...

if exist "!SRC_ROOT!\ParaStyler\saxon-license.lic" (
    copy /y "!SRC_ROOT!\ParaStyler\saxon-license.lic" "!INSTALL_ROOT!\ParaStyler\saxon-license.lic" >nul
    copy /y "!SRC_ROOT!\ParaStyler\saxon-license.lic" "!INSTALL_ROOT!\config\saxon-license.lic" >nul
)
if exist "!SRC_ROOT!\config\saxon-license.lic" (
    copy /y "!SRC_ROOT!\config\saxon-license.lic" "!INSTALL_ROOT!\config\saxon-license.lic" >nul
)

echo             Licence files copied.
echo.

:: ============================================================
::  STEP 8  —  Copy config files
:: ============================================================
echo [step 8/14]  Copying configuration files ...

for %%f in (
    "breakDown.yaml"
    "watcher.yaml"
    "mAnalyser.yaml"
    "mNormalizer.yaml"
    "mMerger.yaml"
    "dbConfig.yaml"
    "dialogConfig.yaml"
    "paraStyles.yaml"
    "breakdownSequence.json"
    "backMatterTitles.json"
    "log_config.cfg"
    "GetSageAid.yaml"
    "createArticleInfo.yaml"
    "spacy_config.cfg"
) do (
    if exist "!SRC_ROOT!\config\%%~f" (
        copy /y "!SRC_ROOT!\config\%%~f" "!INSTALL_ROOT!\config\%%~f" >nul
    )
)

:: Copy CandM icons (used by GUI)
if exist "!SRC_ROOT!\config\CandM.ico" (
    copy /y "!SRC_ROOT!\config\CandM.ico" "!INSTALL_ROOT!\config\CandM.ico" >nul
)
if exist "!SRC_ROOT!\config\CandM.png" (
    copy /y "!SRC_ROOT!\config\CandM.png" "!INSTALL_ROOT!\config\CandM.png" >nul
)

echo             Config files copied.
echo.

:: ============================================================
::  STEP 9  —  Copy SupportingFiles
:: ============================================================
echo [step 9/14]  Copying SupportingFiles ...

for %%f in (
    "SAGE_styles.docx"
    "SAGE_styles.dotx"
    "SAGESTYLES.dotx"
    "CMSTYLES.dotx"
    "SAGE_styles.dot"
    "SAGE_styles.doc"
    "checkDocRunning.exe"
    "checkDocRunning.bat"
    "checkDocRunning.yaml"
    "sageJournalInfo.json"
    "defaultValue.json"
    "BreakDownLogo.png"
    "BreakDown.json"
    "hash_value.txt"
) do (
    if exist "!SRC_ROOT!\SupportingFiles\%%~f" (
        copy /y "!SRC_ROOT!\SupportingFiles\%%~f" "!INSTALL_ROOT!\SupportingFiles\%%~f" >nul
    )
)

:: Copy chromedriver (used by Selenium for article info lookup)
if exist "!SRC_ROOT!\SupportingFiles\chromedriver.exe" (
    copy /y "!SRC_ROOT!\SupportingFiles\chromedriver.exe" "!INSTALL_ROOT!\SupportingFiles\chromedriver.exe" >nul
)

:: Copy handle utilities (used for file-lock detection)
if exist "!SRC_ROOT!\handle.exe"   copy /y "!SRC_ROOT!\handle.exe"   "!INSTALL_ROOT!\handle.exe"   >nul
if exist "!SRC_ROOT!\handle64.exe" copy /y "!SRC_ROOT!\handle64.exe" "!INSTALL_ROOT!\handle64.exe" >nul

echo             SupportingFiles copied.
echo.

:: ============================================================
::  STEP 10  —  Copy XSL stylesheets
:: ============================================================
echo [step 10/14]  Copying XSL stylesheets ...

if exist "!SRC_ROOT!\xsl" (
    xcopy /e /i /y /q "!SRC_ROOT!\xsl" "!INSTALL_ROOT!\xsl" /exclude:!SRC_ROOT!\xsl_exclude.txt >nul 2>&1
    xcopy /e /i /y /q "!SRC_ROOT!\xsl" "!INSTALL_ROOT!\xsl" >nul 2>&1
)

:: Copy ParaStyler XSL files (para_info, author_label, tableFormat, etc.)
for %%f in (
    "para_info.xsl"
    "author_label.xsl"
    "tableFormat.xsl"
    "data.json"
) do (
    if exist "!SRC_ROOT!\xsl\%%~f" (
        copy /y "!SRC_ROOT!\xsl\%%~f" "!INSTALL_ROOT!\xsl\%%~f" >nul
    )
)

:: Copy ParaStyler XSL files from ParaStyler folder
for %%f in (
    "normalize.xsl"
    "normalize_author_names.xsl"
    "paraStyle.xsl"
    "paraAppearance.xsl"
    "main.xsl"
    "fuzzyMatcher.xsl"
    "keywords.xsl"
    "listInfo.xsl"
    "textProps.xsl"
    "xmltoarff.xsl"
) do (
    if exist "!SRC_ROOT!\ParaStyler\%%~f" (
        copy /y "!SRC_ROOT!\ParaStyler\%%~f" "!INSTALL_ROOT!\ParaStyler\%%~f" >nul
    )
)

echo             XSL files copied.
echo.

:: ============================================================
::  STEP 11  —  Write startupConfig.yaml
:: ============================================================
echo [step 11/14]  Writing startupConfig.yaml ...

(
echo MAPPING:
echo     DRIVE: V
echo     PATH: \\192.168.0.102\d$\REPOSITORY
echo CONFIG:
echo     BreakDown: !INSTALL_ROOT!
) > "!INSTALL_ROOT!\startupConfig.yaml"

echo             startupConfig.yaml written.
echo.

:: ============================================================
::  STEP 12  —  Working folders already created in Step 3
::              Verify key ones exist
:: ============================================================
echo [step 12/14]  Verifying working folders ...

set MISSING=0
for %%f in (
    "V:\FOR_BREAKDOWN\INPUT\SAGE"
    "V:\FOR_BREAKDOWN\PROCESS"
    "V:\FOR_BREAKDOWN\ERROR"
    "V:\FOR_BREAKDOWN\LOG"
    "V:\FOR_BREAKDOWN\MERGER_INPUT\SAGE"
    "V:\FOR_BREAKDOWN\ParaStyler_INPUT\SAGE"
) do (
    if not exist %%f (
        echo [WARN] Folder not found: %%f
        set MISSING=1
    )
)
if "!MISSING!"=="0" echo             All working folders present.
echo.

:: ============================================================
::  STEP 13  —  Create start_watcher.bat
:: ============================================================
echo [step 13/14]  Creating watcher startup batch files ...

:: --- HOTFOLDER watcher ---
(
echo @echo off
echo :: =======================================================
echo ::  start_watcher.bat  —  Start BreakDown Watcher
echo ::  Location watched:  HOTFOLDER  (V:\FOR_BREAKDOWN\INPUT\SAGE)
echo ::
echo ::  HOW TO RUN:
echo ::    1. Double-click this file  OR  run from CMD.
echo ::    2. Keep the CMD window OPEN — closing it stops the watcher.
echo ::    3. Press Ctrl+C to stop the watcher gracefully.
echo :: =======================================================
echo.
echo setlocal
echo.
echo set "TOOL_DIR=!INSTALL_ROOT!"
echo set "WATCHER_EXE=!INSTALL_ROOT!\watcher.exe"
echo.
echo if not exist "%%WATCHER_EXE%%" (
echo     echo [ERROR] watcher.exe not found at %%WATCHER_EXE%%
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo ==========================================
echo echo   BreakDown Watcher  -  HOTFOLDER mode
echo echo ==========================================
echo echo   Tool     : %%WATCHER_EXE%%
echo echo   Location : HOTFOLDER
echo echo   Input    : V:\FOR_BREAKDOWN\INPUT\SAGE
echo echo   Started  : %%DATE%% %%TIME%%
echo echo ==========================================
echo echo.
echo echo Press Ctrl+C to stop.
echo echo.
echo.
echo cd /d "!INSTALL_ROOT!"
echo "%%WATCHER_EXE%%" -p="watcher" -l="HOTFOLDER"
echo.
echo echo.
echo echo [INFO] Watcher stopped.
echo pause
echo endlocal
) > "!INSTALL_ROOT!\start_watcher.bat"

:: --- S3 watcher ---
(
echo @echo off
echo :: =======================================================
echo ::  start_watcher_s3.bat  —  Start BreakDown Watcher (S3)
echo ::  Location watched:  S3 bucket
echo ::
echo ::  HOW TO RUN:
echo ::    1. Double-click this file  OR  run from CMD.
echo ::    2. Keep the CMD window OPEN.
echo ::    3. Press Ctrl+C to stop the watcher gracefully.
echo :: =======================================================
echo.
echo setlocal
echo.
echo set "TOOL_DIR=!INSTALL_ROOT!"
echo set "WATCHER_EXE=!INSTALL_ROOT!\watcher.exe"
echo.
echo if not exist "%%WATCHER_EXE%%" (
echo     echo [ERROR] watcher.exe not found at %%WATCHER_EXE%%
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo ==========================================
echo echo   BreakDown Watcher  -  S3 mode
echo echo ==========================================
echo echo   Tool     : %%WATCHER_EXE%%
echo echo   Location : S3
echo echo   Started  : %%DATE%% %%TIME%%
echo echo ==========================================
echo echo.
echo echo Press Ctrl+C to stop.
echo echo.
echo.
echo cd /d "!INSTALL_ROOT!"
echo "%%WATCHER_EXE%%" -p="watcher" -l="S3"
echo.
echo echo.
echo echo [INFO] S3 Watcher stopped.
echo pause
echo endlocal
) > "!INSTALL_ROOT!\start_watcher_s3.bat"

echo             start_watcher.bat created
echo             start_watcher_s3.bat created
echo.

:: ============================================================
::  STEP 14  —  Summary
:: ============================================================
echo.
echo =============================================================
echo   INSTALLATION COMPLETE
echo =============================================================
echo.
echo   Install root     : !INSTALL_ROOT!
echo   startupConfig    : !INSTALL_ROOT!\startupConfig.yaml
echo   Watcher script   : !INSTALL_ROOT!\start_watcher.bat
echo.
echo   WORKING FOLDERS (V:\FOR_BREAKDOWN\):
echo     INPUT\SAGE  INPUT\XYZ  PROCESS  ERROR  LOG
echo     MERGER_INPUT\SAGE   ParaStyler_INPUT\SAGE
echo     BreakDown_INPUT  BreakDown_DONE  BreakDown_ERROR
echo     FOR_CONVERSION\SAGE
echo.
echo   NEXT STEPS:
echo     1.  Review and customise V:\TOOLS\BreakDown\config\watcher.yaml
echo         (set BREAKDOWN_EXE path if needed)
echo     2.  On the WATCHER machine, double-click:
echo         !INSTALL_ROOT!\start_watcher.bat
echo     3.  Keep the CMD window open — it monitors every 10 seconds.
echo     4.  Drop a .zip package in V:\FOR_BREAKDOWN\INPUT\SAGE\
echo         and watch the pipeline run.
echo.
echo =============================================================
echo.

endlocal
