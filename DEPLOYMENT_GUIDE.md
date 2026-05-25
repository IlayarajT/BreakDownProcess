# BreakDown — Deployment & Installation Guide

**Version:** 1.2.5 (build 20260513)  
**Platform:** Windows 10/11 (64-bit)  
**Last updated:** May 2026

---

## Table of Contents

1. [System Overview](#1-system-overview)
2. [Pre-requisites](#2-pre-requisites)
3. [Build the Tool (Developer)](#3-build-the-tool-developer)
4. [Install the Tool](#4-install-the-tool)
5. [Folder Structure Created](#5-folder-structure-created)
6. [Configuration Files Reference](#6-configuration-files-reference)
7. [Running the Watcher](#7-running-the-watcher)
8. [Batch Files Reference](#8-batch-files-reference)
9. [Manual Processing Commands](#9-manual-processing-commands)
10. [Troubleshooting](#10-troubleshooting)
11. [Uninstalling / Upgrading](#11-uninstalling--upgrading)

---

## 1. System Overview

BreakDown is a Windows-based automated manuscript processing pipeline for SAGE Publications. It transforms raw author submission packages (zip/tar/rar/7z) into styled Word documents ready for typesetting.

**Pipeline stages:**

```
Upload ZIP to V:\FOR_BREAKDOWN\INPUT\SAGE\
    ↓  Watcher detects file (every 10 sec)
mAnalyzer   — extract, read metadata, identify JID + AID
    ↓
mNormalizer — open in Word COM, run macros, save, pre-clean via JAR
    ↓
mMerger     — merge multi-part docs into single JID_AID_CLN.docx
    ↓
ParaStyler  — apply SAGE paragraph styles (external Java process)
    ↓
V:\FOR_CONVERSION\SAGE\[JID]\[AID]\*_AS.docx  (output)
```

---

## 2. Pre-requisites

Install all of the following on **each machine** that will run BreakDown.

### 2.1 Required on ALL machines

| Software | Version | Notes |
|---|---|---|
| Windows | 10/11 64-bit | |
| Java JRE/JDK | 17 or later | Must be in `PATH`. Download from [Adoptium](https://adoptium.net). |
| Network drive `V:` | mapped | `\\192.168.0.102\d$\REPOSITORY` |

**Verify Java:**
```cmd
java -version
```
Expected: `openjdk version "17.x.x"` (or later).

**Map the V: drive:**
```cmd
net use V: \\192.168.0.102\d$\REPOSITORY /persistent:yes
```
Or double-click `map_drive.bat`.

### 2.2 Required on Processing Machines (mNormalizer)

| Software | Version | Notes |
|---|---|---|
| Microsoft Word | 2016 / 2019 / 365 | COM automation required |
| Python (optional) | 3.12 | Only if running from source |

> Word must be installed and licensed. The mNormalizer stage opens documents via the Win32 COM API and runs macros.

### 2.3 Required on Dev Machine (build only)

| Software | Version |
|---|---|
| Python | 3.12 |
| PyInstaller | latest |
| All packages in `requirements.txt` |

---

## 3. Build the Tool (Developer)

> Skip this section if you are installing from a pre-built release ZIP.

### 3.1 Build BreakDown.exe + watcher.exe

From the project root (`D:\mProjects\BreakDown\`):

```cmd
:: Patch version bump + PyInstaller build
build.bat

:: Minor version bump
build.bat minor

:: Force a specific version
build.bat set 1.3.0

:: Build without version bump (date only)
build.bat build-only
```

Output: `dist\BreakDown_v<version>\`

### 3.2 Build watcher.exe separately (if needed)

```cmd
python -m PyInstaller watcher.spec --noconfirm
```

Output: `dist\watcher\watcher.exe`

Copy it into `dist\BreakDown_v<version>\watcher.exe`.

### 3.3 What the build produces

```
dist\BreakDown_v1.2.5\
├── BreakDown.exe          ← main processing tool
├── watcher.exe            ← standalone watcher
├── _internal\             ← PyInstaller runtime (DLLs, packages)
└── release_info.txt       ← build metadata
```

---

## 4. Install the Tool

### 4.1 Quick install (recommended)

1. Copy the release folder (or full project folder) to any local machine.
2. Ensure V: drive is mapped (`map_drive.bat`).
3. Open **Command Prompt** (Run as Administrator recommended).
4. Navigate to the folder and run:

```cmd
install_breakdown.bat
```

Or to install to a custom path:

```cmd
install_breakdown.bat "D:\MyTools\BreakDown"
```

The script will:
- Validate Java and V: drive
- Create `V:\TOOLS\BreakDown\` (the install root)
- Copy all EXEs, JARs, config files, templates
- Create all working folders under `V:\FOR_BREAKDOWN\`
- Generate `startupConfig.yaml`
- Generate `start_watcher.bat` and `start_watcher_s3.bat`

### 4.2 Manual install steps

If the install script cannot be run:

#### Step A — Create folder structure
```cmd
mkdir V:\TOOLS\BreakDown
mkdir V:\TOOLS\BreakDown\config
mkdir V:\TOOLS\BreakDown\SupportingFiles
mkdir V:\TOOLS\BreakDown\DocxManipulator
mkdir V:\TOOLS\BreakDown\DocxManipulator\jar
mkdir V:\TOOLS\BreakDown\ParaStyler
mkdir V:\TOOLS\BreakDown\xsl
```

#### Step B — Copy executables
```cmd
copy BreakDown.exe       V:\TOOLS\BreakDown\
copy watcher.exe         V:\TOOLS\BreakDown\
xcopy /e _internal       V:\TOOLS\BreakDown\_internal\
```

#### Step C — Copy JAR files

| Source | Destination |
|---|---|
| `DocxManipulator\docx-manipulator.jar` | `V:\TOOLS\BreakDown\DocxManipulator\sage-auto-styler.jar` |
| `DocxManipulator\jar\*` | `V:\TOOLS\BreakDown\DocxManipulator\jar\` |
| `ParaStyler\saxon9pe.jar` | `V:\TOOLS\BreakDown\ParaStyler\` |
| `ParaStyler\weka-stable-3.6.6.jar` | `V:\TOOLS\BreakDown\ParaStyler\` |
| `ParaStyler\commons-*.jar` | `V:\TOOLS\BreakDown\ParaStyler\` |
| `ParaStyler\guava-10.0.jar` | `V:\TOOLS\BreakDown\ParaStyler\` |
| `ParaStyler\log4j-*.jar` | `V:\TOOLS\BreakDown\ParaStyler\` |
| `ParaStyler\jsr305-*.jar` | `V:\TOOLS\BreakDown\ParaStyler\` |
| `aspose-words\jar\*` | `V:\TOOLS\BreakDown\aspose-words\jar\` |

#### Step D — Copy licence files

| Source | Destination |
|---|---|
| `ParaStyler\saxon-license.lic` | `V:\TOOLS\BreakDown\ParaStyler\saxon-license.lic` |
| `ParaStyler\saxon-license.lic` | `V:\TOOLS\BreakDown\config\saxon-license.lic` |

> ⚠️ **Saxon licence is mandatory.** Without it, XSLT transforms will fail.

#### Step E — Copy config files
```cmd
xcopy /y config\*.yaml       V:\TOOLS\BreakDown\config\
xcopy /y config\*.json       V:\TOOLS\BreakDown\config\
xcopy /y config\*.cfg        V:\TOOLS\BreakDown\config\
```

#### Step F — Copy SupportingFiles
```cmd
xcopy /y SupportingFiles\*   V:\TOOLS\BreakDown\SupportingFiles\
```

Key files to verify:

| File | Purpose |
|---|---|
| `SAGE_styles.docx` | Word template with all SAGE paragraph styles |
| `checkDocRunning.exe` | COM-safe document status checker utility |
| `checkDocRunning.yaml` | `info_path` for STARTED/COMPLETED/ERROR status files |
| `sageJournalInfo.json` | Journal ID → TLA mapping (used by mAnalyzer) |
| `defaultValue.json` | Default field values |

#### Step G — Copy XSL stylesheets
```cmd
xcopy /y xsl\*               V:\TOOLS\BreakDown\xsl\
xcopy /y ParaStyler\*.xsl    V:\TOOLS\BreakDown\ParaStyler\
```

#### Step H — Write startupConfig.yaml
Create `V:\TOOLS\BreakDown\startupConfig.yaml`:

```yaml
MAPPING:
    DRIVE: V
    PATH: \\192.168.0.102\d$\REPOSITORY
CONFIG:
    BreakDown: V:\TOOLS\BreakDown
```

#### Step I — Create working folders
```cmd
mkdir V:\FOR_BREAKDOWN\INPUT\SAGE
mkdir V:\FOR_BREAKDOWN\INPUT\XYZ
mkdir V:\FOR_BREAKDOWN\PROCESS
mkdir V:\FOR_BREAKDOWN\ERROR
mkdir V:\FOR_BREAKDOWN\LOG
mkdir V:\FOR_BREAKDOWN\MERGER_INPUT\SAGE
mkdir V:\FOR_BREAKDOWN\MERGER_INPUT\XYZ
mkdir V:\FOR_BREAKDOWN\ParaStyler_INPUT\SAGE
mkdir V:\FOR_BREAKDOWN\ParaStyler_INPUT\XYZ
mkdir V:\FOR_BREAKDOWN\BreakDown_DONE\SAGE
mkdir V:\FOR_BREAKDOWN\BreakDown_ERROR\SAGE
mkdir V:\FOR_CONVERSION\SAGE
mkdir V:\FOR_CONVERSION\XYZ
```

---

## 5. Folder Structure Created

After installation, `V:\TOOLS\BreakDown\` contains:

```
V:\TOOLS\BreakDown\
│
├── BreakDown.exe                  ← main CLI tool
├── watcher.exe                    ← watcher daemon
├── start_watcher.bat    ◄── Open on Watcher Machine (HOTFOLDER)
├── start_watcher_s3.bat ◄── Open on Watcher Machine (S3)
├── startupConfig.yaml             ← must point to this folder
│
├── _internal\                     ← PyInstaller runtime
│   ├── DocxManipulator\
│   │   ├── sage-auto-styler.jar   ← pre-clean JAR
│   │   └── jar\                   ← Aspose + dependencies
│   └── ParaStyler\
│       ├── saxon9pe.jar            ← XSLT processor
│       ├── saxon-license.lic       ← Saxon PE licence  ⚠️ REQUIRED
│       └── *.xsl
│
├── config\
│   ├── breakDown.yaml             ← folder paths, logger config
│   ├── watcher.yaml               ← customers, hotfolder paths
│   ├── mAnalyser.yaml             ← DocTypes, MergeOrder, folders
│   ├── mNormalizer.yaml           ← macros, kill-process list
│   ├── mMerger.yaml               ← merge filename, source option
│   ├── dbConfig.yaml              ← MySQL config (db_system: false)
│   ├── paraStyles.yaml            ← 60+ SAGE paragraph styles
│   ├── breakdownSequence.json     ← style name → tag mapping
│   ├── backMatterTitles.json      ← back-matter section patterns
│   ├── dialogConfig.yaml          ← GUI dialog configuration
│   ├── log_config.cfg             ← Python logging config
│   └── saxon-license.lic          ← Saxon licence copy
│
├── SupportingFiles\
│   ├── SAGE_styles.docx           ← Word template  ⚠️ REQUIRED
│   ├── SAGESTYLES.dotx
│   ├── CMSTYLES.dotx
│   ├── checkDocRunning.exe        ← file-lock status helper
│   ├── checkDocRunning.yaml
│   ├── sageJournalInfo.json       ← JID→TLA lookup  ⚠️ REQUIRED
│   ├── defaultValue.json
│   └── BreakDownLogo.png
│
├── ParaStyler\
│   ├── saxon9pe.jar
│   ├── saxon-license.lic
│   ├── weka-stable-3.6.6.jar
│   ├── asprop30x.arff.randomCommitee_50.model
│   ├── run.bat
│   └── *.xsl  (normalize, paraStyle, author_label, etc.)
│
├── DocxManipulator\
│   ├── sage-auto-styler.jar       (= docx-manipulator.jar)
│   └── jar\
│       ├── aspose-words-22.10-jdk17.jar
│       ├── jackson-databind-2.9.8.jar
│       ├── jackson-core-2.9.8.jar
│       └── (other dependency JARs)
│
└── xsl\
    ├── para_info.xsl
    ├── author_label.xsl
    ├── tableFormat.xsl
    └── data.json
```

---

## 6. Configuration Files Reference

### 6.1 `startupConfig.yaml` — Bootstrap config

```yaml
MAPPING:
    DRIVE: V
    PATH: \\192.168.0.102\d$\REPOSITORY
CONFIG:
    BreakDown: V:\TOOLS\BreakDown   # must match install root
```

This is the **only** file BreakDown.exe reads from its own folder.  
Everything else is read from the path in `CONFIG.BreakDown`.

---

### 6.2 `config\breakDown.yaml` — Main folder & logger config

```yaml
LOGGER:
    ROOT:      V:\FOR_BREAKDOWN\LOG\break_down.log
    BREAK_DOWN: V:\FOR_BREAKDOWN\LOG\[CUSTOMER]\[JID]\[AID]\[JID]_[AID]_LOG.log

FOLDERS:
    INPUT:          V:\FOR_BREAKDOWN\INPUT
    PROCESS:        V:\FOR_BREAKDOWN\PROCESS
    ERROR:          V:\FOR_BREAKDOWN\ERROR
    OUTPUT:         V:\FOR_CONVERSION\[CUSTOMER]\[JID]\[AID]
    LOG:            V:\FOR_BREAKDOWN\LOG
    MERGER_INPUT:   V:\FOR_BREAKDOWN\MERGER_INPUT\[CUSTOMER]
    MERGER_ERROR:   V:\FOR_BREAKDOWN\MERGER_ERROR\[CUSTOMER]
    ParaStyler_INPUT:  V:\FOR_BREAKDOWN\ParaStyler_INPUT\[CUSTOMER]
    ParaStyler_ERROR:  V:\FOR_BREAKDOWN\ParaStyler_ERROR\[CUSTOMER]
    BreakDown_INPUT:   V:\FOR_BREAKDOWN\BreakDown_INPUT\[CUSTOMER]
    BreakDown_ERROR:   V:\FOR_BREAKDOWN\BreakDown_ERROR\[CUSTOMER]
    BreakDown_DONE:    V:\FOR_BREAKDOWN\BreakDown_DONE\[CUSTOMER]

TimeOut: 300
```

Placeholders `[CUSTOMER]`, `[JID]`, `[AID]` are substituted at runtime.

---

### 6.3 `config\watcher.yaml` — Watcher configuration

```yaml
BREAKDOWN_EXE: V:\TOOLS\BreakDown\BreakDown.exe

HOTFOLDER:
  CUSTOMERS:
    - SAGE           # ← add/remove customer names here

  SAGE:
    FOLDERS:
      INPUT: V:\FOR_BREAKDOWN\INPUT\SAGE
      ERROR: V:\FOR_BREAKDOWN\ERROR\SAGE

S3:
  CUSTOMERS:
    - SAGE
  SAGE:
    REPOSITORY:      # S3 bucket name
    ACCESSID:        # AWS access key ID
    ACCESSKEY:       # AWS secret key
```

**To add a new customer:**
1. Add the customer name to `CUSTOMERS` list.
2. Add the customer block with `INPUT` and `ERROR` folders.
3. Create the folders on V: drive.

---

### 6.4 `config\mAnalyser.yaml` — Analyser configuration

```yaml
DocTypes:      [title, main, ack, author_note, bio, figure, table, other]
MergeOrder:    {1: title, 2: main, 3: ack, 4: author_note, 5: bio, 6: figure, 7: table}

FOLDERS:
    MERGER: V:\FOR_BREAKDOWN\MERGER_INPUT\[CUSTOMER]\[JID]_[AID]
    PROCESS: V:\FOR_BREAKDOWN\PROCESS
    ERROR:   V:\FOR_BREAKDOWN\ERROR
```

---

### 6.5 `config\mNormalizer.yaml` — Normaliser & Macros

```yaml
KillProcess:
    - WINWORD.exe
    - EXCEL.exe
    - POWERPNT.exe

SAGE:
    RunMacros:         # macros run in order on each document
        - AcceptTrackChange
        - TotalCleanUP
        - UnlinkFieldcodesExceptMath
        - ConvertEndnoteToFootnote
        - EnqTableToText
        - RemoveLineNumbers
        - FlattenTextBoxes
        - RemoveAllFramesInDoc
        - RemoveUnwantedSpaces
        - EliminateMultipleSpaces
        - RemoveDocVar
        - CitationsToStaticText
        - RemoveFirstLineIndent
        - RemoveAllHyperlinks
        - CleanTableCells
    PreClean:
        enabled: false           # set true to invoke sage-auto-styler.jar
        replace_macros: false    # set true to skip Word macros
        jar_name: sage-auto-styler.jar
        jar_args: ["-pre"]
        timeout: 300
```

---

### 6.6 `config\mMerger.yaml`

```yaml
FILENAME:
    SAGE: FOLDER_CLN.docx   # output filename template

SOURCE: Retain              # Retain | Remove  (source .docx files after merge)
```

---

### 6.7 `config\dbConfig.yaml`

```yaml
db_system: false   # set true to enable MySQL logging
```

When `db_system: true`, also configure `config\database.ini` with MySQL credentials.

---

## 7. Running the Watcher

### 7.1 On the Watcher Machine

1. Ensure V: drive is mapped.
2. Ensure Java is in PATH.
3. Navigate to `V:\TOOLS\BreakDown\`.
4. **Double-click `start_watcher.bat`**.
5. A CMD window opens and shows:

```
============================================================
  BreakDown Watcher  |  HOTFOLDER mode
============================================================
  Tool root : V:\TOOLS\BreakDown
  Exe       : V:\TOOLS\BreakDown\watcher.exe
  Location  : HOTFOLDER
  Input     : V:\FOR_BREAKDOWN\INPUT\SAGE
  Started   : 25/05/2026 09:00:00
============================================================

  Press Ctrl+C to stop the watcher.

Progress: [==================================================] Watching...
```

6. **Keep this window open.** Closing it stops the watcher.

### 7.2 What the watcher does

- Polls `V:\FOR_BREAKDOWN\INPUT\SAGE\` every **10 seconds**.
- For each file found:
  - Calls `BreakDown.exe -p=mAnalyzer -f="<file>" -c="SAGE"`
  - Waits up to **900 seconds** (15 min) for completion.
  - On timeout/error: moves file to `V:\FOR_BREAKDOWN\ERROR\SAGE\`, logs to `error_log.html`.
- Every **300 seconds**: logs CPU and RAM usage.

### 7.3 Stopping the watcher

Press `Ctrl+C` in the CMD window. The watcher catches SIGINT and exits cleanly.

---

## 8. Batch Files Reference

| File | Machine | Purpose |
|---|---|---|
| `start_watcher.bat` | **Watcher Machine** | Start the watcher daemon (HOTFOLDER) |
| `start_watcher_s3.bat` | Watcher Machine | Start watcher for S3 input |
| `map_drive.bat` | Any | Map V: drive to the repository server |
| `install_breakdown.bat` | Any (first setup) | Full installation |
| `local_styler.bat` | Any | Run sage-auto-styler.jar on one .docx manually |
| `local_styler_dev.bat` | Dev Machine | Dev/source version of local styler |
| `build.bat` | Dev Machine | Build BreakDown.exe + watcher.exe via PyInstaller |
| `ParaStyler\run.bat` | ParaStyler Machine | Run ParaStyler on a package |

### `start_watcher.bat` — must remain open

This batch file launches `watcher.exe` as a **foreground process** in the same CMD window. The window **must stay open** for the watcher to continue running.

To run as a background service (optional), use Windows Task Scheduler or NSSM (Non-Sucking Service Manager):

```cmd
nssm install BreakDownWatcher "V:\TOOLS\BreakDown\watcher.exe" -p="watcher" -l="HOTFOLDER"
nssm start BreakDownWatcher
```

---

## 9. Manual Processing Commands

All manual commands are run via `BreakDown.exe` (or `python main.py` from source).

### 9.1 Run mAnalyzer on a specific file

```cmd
V:\TOOLS\BreakDown\BreakDown.exe -p="mAnalyzer" -f="V:\FOR_BREAKDOWN\INPUT\SAGE\Article_Attachments-2026-05-25.zip" -c="SAGE"
```

### 9.2 Run mNormalizer manually

```cmd
V:\TOOLS\BreakDown\BreakDown.exe -p="mNormalizer" -c="SAGE" -jf="V:\FOR_BREAKDOWN\MERGER_INPUT\SAGE\SGO_123456\SGO_123456.json"
```

### 9.3 Run mMerger manually

```cmd
V:\TOOLS\BreakDown\BreakDown.exe -p="mMerger" -c="SAGE" -jf="V:\FOR_BREAKDOWN\MERGER_INPUT\SAGE\SGO_123456\SGO_123456.json"
```

### 9.4 Open mSelect GUI (manual doc selection)

```cmd
V:\TOOLS\BreakDown\BreakDown.exe -p="mSelect"
```

### 9.5 Update SAGE Journal Info

```cmd
V:\TOOLS\BreakDown\BreakDown.exe -p="createSageJournalInfo"
```

### 9.6 Run local ParaStyler (pre-clean only) on a .docx

```cmd
V:\TOOLS\BreakDown\local_styler.bat "V:\FOR_BREAKDOWN\ParaStyler_INPUT\SAGE\SGO_123456\SGO_123456_CLN.docx"
```

---

## 10. Troubleshooting

### 10.1 Watcher not picking up files

| Symptom | Check |
|---|---|
| Files sit in INPUT\ unprocessed | Is `start_watcher.bat` window still open? |
| Error: `watcher.exe not found` | Run `install_breakdown.bat` first |
| Error: `V: drive not mapped` | Run `map_drive.bat` |
| Timeout errors in log | Check Word is not hung; kill WINWORD.EXE |

### 10.2 mNormalizer / Word COM errors

| Symptom | Solution |
|---|---|
| `CLSIDToClassMap` error | BreakDown auto-clears gen_py cache on startup |
| `COMException` on open | Word may be hung — kill WINWORD.EXE via Task Manager |
| Document read-only error | Close any open Word windows before running |
| Macro not found | Ensure CMSTYLES.dotx / SAGESTYLES.dotx are loaded in Word |

### 10.3 JAR errors

| Symptom | Solution |
|---|---|
| `JAR file not found: sage-auto-styler.jar` | Check `_internal\DocxManipulator\` exists in install root |
| `Java not found` | Add Java to PATH: `setx PATH "%PATH%;C:\Program Files\Java\jdk-17\bin"` |
| Saxon licence error | Verify `ParaStyler\saxon-license.lic` exists and is valid |
| `weka` model error | Check `asprop30x.arff.randomCommitee_50.model` is present |

### 10.4 Metadata not found error

| Symptom | Solution |
|---|---|
| Package moved to ERROR — "Metadata file not found" | Check zip contains `*-metadata.xml` or `SAGE-metadata-*.xml` |
| `GetArticleId` / smart_login fails | Update `sageJournalInfo.json`; check network/VPN for SAGE login |

### 10.5 Log locations

| Log | Path |
|---|---|
| Main BreakDown log | `V:\FOR_BREAKDOWN\LOG\break_down.log` |
| Per-article log | `V:\FOR_BREAKDOWN\LOG\[CUSTOMER]\[JID]\[AID]\[JID]_[AID]_LOG.log` |
| Watcher error log | `V:\FOR_BREAKDOWN\LOG\error_log.html` |
| Watcher startup log | `V:\TOOLS\BreakDown\watcher_startup.log` |

---

## 11. Uninstalling / Upgrading

### Upgrade

1. Build the new version: `build.bat minor` (or `patch` / `major`).
2. Stop the watcher (`Ctrl+C` in the watcher CMD window).
3. Run `install_breakdown.bat` — it will overwrite existing files.
4. Restart the watcher.

> Config files are overwritten during upgrade. Back up any customised YAML files first.

### Uninstall

```cmd
:: Stop watcher first (Ctrl+C in watcher window)

:: Remove install folder
rmdir /s /q V:\TOOLS\BreakDown

:: Remove working folders (CAUTION: this deletes all processed files)
:: rmdir /s /q V:\FOR_BREAKDOWN
:: rmdir /s /q V:\FOR_CONVERSION

:: Unmap drive
net use V: /delete
```

---

## Appendix A — Full File Inventory

### JARs required in `V:\TOOLS\BreakDown\`

| JAR | Folder | Purpose |
|---|---|---|
| `sage-auto-styler.jar` | `DocxManipulator\` | Pre-clean + style application |
| `docx-manipulator.jar` | `DocxManipulator\` | Same as above (alias) |
| `aspose-words-22.10-jdk17.jar` | `DocxManipulator\jar\` | Word document manipulation |
| `jackson-databind-2.9.8.jar` | `DocxManipulator\jar\` | JSON serialisation |
| `jackson-core-2.9.8.jar` | `DocxManipulator\jar\` | JSON core |
| `jackson-annotations-2.9.0.jar` | `DocxManipulator\jar\` | JSON annotations |
| `jsoup-1.8.3.jar` | `DocxManipulator\jar\` | HTML parsing |
| `commons-io-2.4.jar` | `DocxManipulator\jar\` | File utilities |
| `commons-logging-1.1.1.jar` | `DocxManipulator\jar\` | Logging |
| `log4j-api-2.16.0.jar` | `DocxManipulator\jar\` | Logging API |
| `log4j-core-2.16.0.jar` | `DocxManipulator\jar\` | Logging core |
| `slf4j-api-1.5.6.jar` | `DocxManipulator\jar\` | SLF4J |
| `filters-2.0.235.jar` | `DocxManipulator\jar\` | Filters |
| `gluegen-rt-main-2.3.2.jar` | `DocxManipulator\jar\` | OpenGL (Aspose) |
| `jogl-all-main-2.3.2.jar` | `DocxManipulator\jar\` | OpenGL (Aspose) |
| `jai-imageio-core-1.3.0.jar` | `DocxManipulator\jar\` | Image I/O |
| `mime-util-2.1.1.jar` | `DocxManipulator\jar\` | MIME detection |
| `saxon9pe.jar` | `ParaStyler\` | Saxon XSLT PE processor |
| `weka-stable-3.6.6.jar` | `ParaStyler\` | ML classifier |
| `commons-cli-1.2.jar` | `ParaStyler\` | CLI parsing |
| `commons-io-2.4.jar` | `ParaStyler\` | File I/O |
| `guava-10.0.jar` | `ParaStyler\` | Google Guava |
| `log4j-api-2.0-beta8.jar` | `ParaStyler\` | Logging |
| `log4j-core-2.0-beta8.jar` | `ParaStyler\` | Logging |
| `jsr305-1.3.9.jar` | `ParaStyler\` | JSR-305 annotations |

### Licence files required

| File | Folder | Note |
|---|---|---|
| `saxon-license.lic` | `ParaStyler\` | **Mandatory** for Saxon PE |
| `saxon-license.lic` | `config\` | Backup copy |

### Critical supporting files

| File | Folder | Note |
|---|---|---|
| `SAGE_styles.docx` | `SupportingFiles\` | **Mandatory** – Word style template |
| `sageJournalInfo.json` | `SupportingFiles\` | **Mandatory** – JID/TLA lookup |
| `checkDocRunning.exe` | `SupportingFiles\` | COM status helper |
| `checkDocRunning.yaml` | `SupportingFiles\` | info_path config |

---

*End of Deployment Guide*
