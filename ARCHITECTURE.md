# BreakDown — Architecture & Deployment Documentation

---

## 1. Component Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                         BREAKDOWN PROCESSING SYSTEM                             │
│                              v1.2.5  (build 20260513)                           │
└─────────────────────────────────────────────────────────────────────────────────┘

 ┌────────────────────────────────────────────────────────────────────────────────┐
 │  LAYER 1 — ENTRY POINT / ORCHESTRATION                                         │
 │                                                                                │
 │  ┌─────────────────────────────────────────────────────────────────────────┐   │
 │  │  watcher.exe  (watcher.py)                                               │   │
 │  │  ─────────────────────────────────────────────────────────────────────  │   │
 │  │  • Runs as a continuous loop (schedule.every 10 seconds)                 │   │
 │  │  • Reads watcher.yaml → resolves CUSTOMERS & INPUT folders               │   │
 │  │  • Polls V:\FOR_BREAKDOWN\INPUT\SAGE\ for new packages                   │   │
 │  │  • Invokes  BreakDown.exe -p=mAnalyzer -f=<file> -c=<customer>          │   │
 │  │  • Resource monitor fires every 300 s (log CPU/RAM)                      │   │
 │  │  • Moves failed files → ERROR folder, logs to error_log.html             │   │
 │  │  • Timeout per file: 900 seconds                                         │   │
 │  └─────────────────────────────────────────────────────────────────────────┘   │
 │            │  subprocess call                                                   │
 │            ▼                                                                    │
 │  ┌─────────────────────────────────────────────────────────────────────────┐   │
 │  │  BreakDown.exe  (main.py)  — CLI Dispatcher                              │   │
 │  │  ─────────────────────────────────────────────────────────────────────  │   │
 │  │  • Reads startupConfig.yaml → resolves config folder on V:\TOOLS\       │   │
 │  │  • Reads breakDown.yaml → folder map, logger paths                       │   │
 │  │  • Cleans win32com gen_py cache on startup (COM stability)               │   │
 │  │  • Validates CLI args  (-p, -f, -c, -j, -jf, -jid, -aid, -l)           │   │
 │  │  • Routes to sub-process based on -p value:                              │   │
 │  │       mAnalyzer  ──▶  mAnalyser module                                  │   │
 │  │       mSelect    ──▶  Qt6 GUI dialog (manual selection)                  │   │
 │  │       mNormalizer──▶  ProcessDoc module                                  │   │
 │  │       mMerger    ──▶  DocxMerger module                                  │   │
 │  │       createSageJournalInfo ──▶ GetJournalInfo module                    │   │
 │  └─────────────────────────────────────────────────────────────────────────┘   │
 └────────────────────────────────────────────────────────────────────────────────┘

 ┌────────────────────────────────────────────────────────────────────────────────┐
 │  LAYER 2 — PIPELINE STAGES                                                      │
 │                                                                                │
 │  STAGE 1: mAnalyzer  (mAnalyser.py)                                            │
 │  ┌─────────────────────────────────────────────────────────────────────────┐   │
 │  │  Input: zip / tar / rar / 7z / directory  from hotfolder                 │   │
 │  │                                                                          │   │
 │  │  ┌─────────────────┐    ┌──────────────────┐    ┌───────────────────┐   │   │
 │  │  │ 1. Extract       │───▶│ 2. Classify Files│───▶│ 3. Read Metadata  │   │   │
 │  │  │  (zip/tar/rar/7z)│    │  short_metadata  │    │  XML → JID + AID  │   │   │
 │  │  │  Recursive       │    │  long_metadata   │    │  + ms_no          │   │   │
 │  │  │  archive support │    │  doc / graphics  │    │                   │   │   │
 │  │  └─────────────────┘    └──────────────────┘    └────────┬──────────┘   │   │
 │  │                                                           │              │   │
 │  │  ┌─────────────────┐    ┌──────────────────┐             │              │   │
 │  │  │ 5. Stage Files   │◀───│ 4. GetArticleId  │◀───────────┘              │   │
 │  │  │  MERGER_INPUT or │    │  smart_login()   │                           │   │
 │  │  │  ParaStyler_INPUT│    │  (Selenium/       │                           │   │
 │  │  │  (skip_merger    │    │   sageJournal     │                           │   │
 │  │  │   single docx)   │    │   JSON lookup)    │                           │   │
 │  │  └─────────────────┘    └──────────────────┘                           │   │
 │  │                                                                          │   │
 │  │  Metadata types:                                                         │   │
 │  │    Short metadata: SAGE-metadata-*.xml | dd-nnn.xml pattern             │   │
 │  │    Long metadata:  *-metadata.xml (article_set root)                    │   │
 │  │    Auto-merge:  long-meta with merge=true | single doc                  │   │
 │  │    Manual-merge: long-meta with merge=false | multi-doc unknown order   │   │
 │  └─────────────────────────────────────────────────────────────────────────┘   │
 │            │  auto-merge path                  │  skip_merger (single docx)    │
 │            ▼                                   ▼                               │
 │  STAGE 2: mNormalizer  (mNormalizer.py)   STAGE 2b: Direct to ParaStyler       │
 │  ┌─────────────────────────────────────┐   ┌──────────────────────────────┐   │
 │  │  Input: selected .doc/.docx files    │   │  JID_AID_CLN.docx renamed    │   │
 │  │                                      │   │  → V:\...\ParaStyler_INPUT\  │   │
 │  │  1. Kill running WINWORD/EXCEL       │   └──────────────────────────────┘   │
 │  │  2. DocxPreClean (XML-level clean)   │                                      │
 │  │  3. DocxImageCleaner (large images)  │                                      │
 │  │  4. .doc → .docx conversion          │                                      │
 │  │     (DocToDocx via COM)              │                                      │
 │  │  5. Open in Word (COM automation)    │                                      │
 │  │     WordSessionController            │                                      │
 │  │     (restart every 3 files,          │                                      │
 │  │      max 10 restarts)                │                                      │
 │  │  6. Run Word Macros:                 │                                      │
 │  │     AcceptTrackChange               │                                      │
 │  │     TotalCleanUP                    │                                      │
 │  │     UnlinkFieldcodesExceptMath      │                                      │
 │  │     ConvertEndnoteToFootnote        │                                      │
 │  │     RemoveLineNumbers + 12 more     │                                      │
 │  │  7. Save document                   │                                      │
 │  │  8. PreClean via sage-auto-styler   │                                      │
 │  │     .jar (if PreClean.enabled)       │                                      │
 │  │  9. Status files: STARTED/COMPLETED │                                      │
 │  │     /ERROR in info_path             │                                      │
 │  └─────────────────────────────────────┘                                      │
 │            │                                                                   │
 │            ▼  (if auto-merge = True)                                           │
 │  STAGE 3: mMerger  (mMerger.py)                                                │
 │  ┌─────────────────────────────────────────────────────────────────────────┐   │
 │  │  Input: normalised .docx files in MERGER_INPUT\[CUSTOMER]\[JID_AID]\   │   │
 │  │                                                                          │   │
 │  │  merge_docx_robust():  docxcompose Composer                              │   │
 │  │    • Append files in MergeOrder (title→main→ack→bio→fig→table)         │   │
 │  │    • Output: FOLDER_CLN.docx                                             │   │
 │  │    • Source files: Retain → moved to docs\ subfolder                    │   │
 │  │                    Remove → deleted                                      │   │
 │  │  move_files_to_docs():                                                   │   │
 │  │    → Moves merged output to ParaStyler_INPUT\[CUSTOMER]\[JID_AID]\     │   │
 │  └─────────────────────────────────────────────────────────────────────────┘   │
 │            │                                                                   │
 │            ▼                                                                   │
 │  STAGE 4: ParaStyler  (external Java / XSL process)                            │
 │  ┌─────────────────────────────────────────────────────────────────────────┐   │
 │  │  Input: *_CLN.docx in V:\FOR_BREAKDOWN\ParaStyler_INPUT\               │   │
 │  │                                                                          │   │
 │  │  Two sub-engines:                                                        │   │
 │  │    A. Saxon XSLT (ParaStyler\saxon9pe.jar + *.xsl)                      │   │
 │  │       para_info.xsl / author_label.xsl / tableFormat.xsl               │   │
 │  │       Requires: saxon-license.lic                                        │   │
 │  │                                                                          │   │
 │  │    B. DocxManipulator JAR (sage-auto-styler.jar)                        │   │
 │  │       java -jar sage-auto-styler.jar -dx <file> -ipas                   │   │
 │  │       Produces: *_AS.docx                                               │   │
 │  │       ApplyStyles.py: maps XML para names → SAGE style names            │   │
 │  │       Uses SAGE_styles.docx template for style definitions              │   │
 │  │                                                                          │   │
 │  │  Output: *_AS.docx → V:\FOR_CONVERSION\[CUSTOMER]\[JID]\[AID]\        │   │
 │  └─────────────────────────────────────────────────────────────────────────┘   │
 └────────────────────────────────────────────────────────────────────────────────┘

 ┌────────────────────────────────────────────────────────────────────────────────┐
 │  LAYER 3 — SUPPORTING MODULES                                                   │
 │                                                                                │
 │  ┌──────────────────┐  ┌────────────────────┐  ┌──────────────────────────┐   │
 │  │  DataBase         │  │  TransformXml       │  │  CreateArticleInfo       │   │
 │  │  (dbprocess.py)   │  │  (TransformXml.py)  │  │  (CreateArticleInfo.py) │   │
 │  │  ─────────────    │  │  ─────────────────  │  │  ────────────────────   │   │
 │  │  MySQL tracking   │  │  Saxon XSLT engine  │  │  GetArticleId.smart_    │   │
 │  │  package_id       │  │  (lazy init JVM)    │  │  login() : Selenium     │   │
 │  │  unique_id        │  │  JET crash-safe      │  │  ChromeDriver           │   │
 │  │  process_status   │  │  jet_dump cleanup   │  │  Reads sageJournalInfo  │   │
 │  │  db_system:false  │  │  One JVM/process    │  │  .json for TLA mapping  │   │
 │  │  by default       │  └────────────────────┘  └──────────────────────────┘   │
 │  └──────────────────┘                                                          │
 │  ┌──────────────────┐  ┌────────────────────┐  ┌──────────────────────────┐   │
 │  │  utils/           │  │  com_manager.py    │  │  DocxPreClean.py         │   │
 │  │  file_utils.py    │  │  ─────────────────  │  │  ────────────────────   │   │
 │  │  process_runner   │  │  COMManager         │  │  XML-level DOCX clean    │   │
 │  │  resource_monitor │  │  WordApplicationMgr │  │  before Word opens it    │   │
 │  │  error_logger     │  │  COM context mgr    │  │                          │   │
 │  │  progress.py      │  │  gen_py cache repair│  │  DocxImageCleaner.py    │   │
 │  │  retry.py         │  │  + com_utils.py     │  │  Strips large images    │   │
 │  └──────────────────┘  └────────────────────┘  └──────────────────────────┘   │
 │  ┌──────────────────┐  ┌────────────────────┐  ┌──────────────────────────┐   │
 │  │  loadconfig.py    │  │  getAppPath.py      │  │  version.py              │   │
 │  │  ─────────────    │  │  ─────────────────  │  │  ─────────────────────  │   │
 │  │  Reads            │  │  Frozen EXE path    │  │  __version__ = 1.2.5    │   │
 │  │  startupConfig    │  │  vs script path     │  │  get_version_string()   │   │
 │  │  → configFolder   │  └────────────────────┘  └──────────────────────────┘   │
 │  └──────────────────┘                                                          │
 └────────────────────────────────────────────────────────────────────────────────┘

 ┌────────────────────────────────────────────────────────────────────────────────┐
 │  LAYER 4 — CONFIGURATION FILES                                                  │
 │                                                                                │
 │  startupConfig.yaml      → CONFIG.BreakDown = V:\TOOLS\BreakDown               │
 │  config\breakDown.yaml   → Folder paths, logger, timeout                       │
 │  config\watcher.yaml     → BREAKDOWN_EXE path, HOTFOLDER customer list        │
 │  config\mAnalyser.yaml   → DocTypes, MergeOrder, folder overrides             │
 │  config\mNormalizer.yaml → KillProcess list, RunMacros, PreClean config       │
 │  config\mMerger.yaml     → FILENAME pattern, SOURCE option                    │
 │  config\dbConfig.yaml    → db_system: false (MySQL optional)                  │
 │  config\paraStyles.yaml  → 60+ SAGE para style definitions + XML→style map   │
 │  config\breakdownSequence.json → breakdownStyles + breakdownMappingTags       │
 │  config\backMatterTitles.json  → back-matter section title patterns           │
 │  config\backMatterTitles.json  → back-matter section title patterns           │
 │  SupportingFiles\sageJournalInfo.json → JID → TLA + metadata lookup          │
 │  SupportingFiles\SAGE_styles.docx     → Word template with SAGE styles        │
 │  SupportingFiles\checkDocRunning.yaml → info_path for status files            │
 └────────────────────────────────────────────────────────────────────────────────┘

 ┌────────────────────────────────────────────────────────────────────────────────┐
 │  LAYER 5 — JAVA COMPONENTS                                                      │
 │                                                                                │
 │  DocxManipulator\                                                              │
 │  ├── docx-manipulator.jar         ← main JAR (sage-auto-styler alias)          │
 │  └── jar\                                                                      │
 │      ├── newASjid.yml             ← JAXB config                                │
 │      ├── aspose-words-22.10-jdk17.jar                                          │
 │      ├── jackson-databind-2.9.8.jar                                            │
 │      ├── jsoup-1.8.3.jar                                                       │
 │      └── log4j-core-2.16.0.jar + (10 more dependency JARs)                    │
 │                                                                                │
 │  ParaStyler\                                                                   │
 │  ├── saxon9pe.jar                 ← Saxon XSLT processor                       │
 │  ├── saxon-license.lic            ← Saxon PE licence                           │
 │  ├── weka-stable-3.6.6.jar        ← ML model runner                            │
 │  ├── asprop30x.arff.randomCommitee_50.model ← trained classifier              │
 │  └── (commons-*, guava, log4j JARs)                                            │
 │                                                                                │
 │  aspose-words\jar\                                                             │
 │  └── aspose-words-22.10-jdk17.jar + gluegen/jogl/jai JARs                     │
 └────────────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Deployment Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                         DEPLOYMENT TOPOLOGY                                     │
└─────────────────────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────────────────┐
│  NETWORK SHARE  (File Server)                                                │
│  \\192.168.0.102\d$\REPOSITORY   ←── mapped as  V:\  on all machines       │
│                                                                              │
│  V:\                                                                         │
│  ├── TOOLS\                                                                  │
│  │   └── BreakDown\          ← INSTALLATION FOLDER (see Section 3)          │
│  │       ├── BreakDown.exe                                                   │
│  │       ├── watcher.exe                                                     │
│  │       ├── start_watcher.bat   ← opened on WATCHER MACHINE               │
│  │       ├── startupConfig.yaml                                              │
│  │       ├── _internal\          (PyInstaller runtime, only if --onedir)    │
│  │       │   ├── DocxManipulator\                                            │
│  │       │   │   ├── sage-auto-styler.jar                                    │
│  │       │   │   └── jar\(deps)                                              │
│  │       │   └── ParaStyler\                                                 │
│  │       │       ├── saxon9pe.jar                                            │
│  │       │       └── saxon-license.lic                                       │
│  │       ├── config\                                                         │
│  │       │   ├── breakDown.yaml                                              │
│  │       │   ├── watcher.yaml                                                │
│  │       │   ├── mAnalyser.yaml                                              │
│  │       │   ├── mNormalizer.yaml                                            │
│  │       │   ├── mMerger.yaml                                                │
│  │       │   ├── dbConfig.yaml                                               │
│  │       │   ├── paraStyles.yaml                                             │
│  │       │   ├── breakdownSequence.json                                      │
│  │       │   ├── backMatterTitles.json                                       │
│  │       │   ├── dialogConfig.yaml                                           │
│  │       │   ├── log_config.cfg                                              │
│  │       │   └── saxon-license.lic                                           │
│  │       └── SupportingFiles\                                                │
│  │           ├── SAGE_styles.docx                                            │
│  │           ├── SAGESTYLES.dotx                                             │
│  │           ├── CMSTYLES.dotx                                               │
│  │           ├── checkDocRunning.exe                                         │
│  │           ├── checkDocRunning.yaml                                        │
│  │           ├── sageJournalInfo.json                                        │
│  │           ├── defaultValue.json                                           │
│  │           └── BreakDownLogo.png                                           │
│  │                                                                           │
│  └── FOR_BREAKDOWN\          ← WORKING FOLDERS                              │
│      ├── INPUT\                                                              │
│      │   └── SAGE\           ← Upstream deposits packages HERE             │
│      │       └── Article_Attachments-YYYY-MM-DD-HH-MM-SS.zip               │
│      ├── PROCESS\            ← Extracted in-flight packages                 │
│      ├── ERROR\              ← Failed packages (with sub-folders)           │
│      ├── LOG\                ← break_down.log + per-article logs            │
│      │   └── [CUSTOMER]\[JID]\[AID]\                                        │
│      ├── MERGER_INPUT\                                                       │
│      │   └── SAGE\[JID_AID]\ ← Pre-merge staged .docx files                │
│      ├── MERGER_ERROR\                                                       │
│      ├── ParaStyler_INPUT\                                                   │
│      │   └── SAGE\[JID_AID]\ ← *_CLN.docx awaiting ParaStyler              │
│      ├── ParaStyler_ERROR\                                                   │
│      ├── BreakDown_INPUT\                                                    │
│      ├── BreakDown_DONE\                                                     │
│      └── BreakDown_ERROR\                                                    │
│                                                                              │
│  V:\FOR_CONVERSION\                                                          │
│      └── SAGE\[JID]\[AID]\   ← Final *_AS.docx output                      │
└──────────────────────────────────────────────────────────────────────────────┘
         ▲                    ▲                     ▲
         │  V: mapped          │  V: mapped           │  V: mapped
         │                    │                     │
┌────────┴───────┐  ┌─────────┴──────┐  ┌──────────┴──────────┐
│  WATCHER       │  │  PROCESSING    │  │  PARASTYLER          │
│  MACHINE       │  │  MACHINE(S)    │  │  MACHINE             │
│  ────────      │  │  ────────────  │  │  ─────────────────   │
│  Runs:         │  │  BreakDown.exe │  │  External Java       │
│  watcher.exe   │  │  mAnalyzer     │  │  ParaStyler process  │
│  via           │  │  mNormalizer   │  │  (manual or auto     │
│  start_watcher │  │  mMerger       │  │   triggered)         │
│  .bat          │  │  MS Word       │  │  Monitors:           │
│                │  │  (COM auto)    │  │  ParaStyler_INPUT\   │
│  Polls every   │  │                │  │  Produces: *_AS.docx │
│  10 seconds    │  │  Java (JRE 17) │  │  → FOR_CONVERSION\  │
│                │  │  required for  │  │                      │
│  CMD window    │  │  JAR invocation│  │                      │
│  stays open    │  │                │  │                      │
└────────────────┘  └────────────────┘  └──────────────────────┘

═══════════════════════════════════════════════════════════════════
                    DATA FLOW SUMMARY
═══════════════════════════════════════════════════════════════════

  Upstream          Watcher            BreakDown Pipeline
  ─────────         ─────────          ──────────────────────────────────────────────
  Drop ZIP   ──▶   Detect new   ──▶   mAnalyzer    ──▶   mNormalizer   ──▶
  to INPUT\         file every          Extract            Word COM
  SAGE\             10 sec              Classify           Run Macros
                                        Metadata            Save .docx
                                        GetArticleId        PreClean JAR
                                        ↙          ↘
                                  Single        Multi-doc
                                  docx           or long-meta
                                  ↓              ↓
                             skip_merger      mMerger
                                  ↓         Merge→CLN.docx
                                  └─────────────┘
                                        ↓
                                 ParaStyler_INPUT\
                                        ↓
                                  ParaStyler
                                  Saxon + JAR
                                        ↓
                                 FOR_CONVERSION\
                                   *_AS.docx  ✓
```

---

## 3. Installation Folder Structure (Target)

```
V:\TOOLS\BreakDown\                         ← INSTALL_ROOT
│
├── BreakDown.exe                            main processing tool
├── watcher.exe                              standalone watcher process
├── start_watcher.bat          ◄─── OPEN THIS on Watcher Machine
├── start_watcher_s3.bat       ◄─── (optional) for S3 location
├── startupConfig.yaml                       points to config folder
│
├── _internal\                               (PyInstaller --onedir runtime)
│   ├── DocxManipulator\
│   │   ├── sage-auto-styler.jar
│   │   └── jar\
│   │       ├── aspose-words-22.10-jdk17.jar
│   │       ├── jackson-databind-2.9.8.jar
│   │       ├── jackson-core-2.9.8.jar
│   │       ├── jackson-annotations-2.9.0.jar
│   │       ├── jsoup-1.8.3.jar
│   │       ├── commons-io-2.4.jar
│   │       ├── commons-logging-1.1.1.jar
│   │       ├── log4j-api-2.16.0.jar
│   │       ├── log4j-core-2.16.0.jar
│   │       ├── slf4j-api-1.5.6.jar
│   │       ├── filters-2.0.235.jar
│   │       └── (gluegen / jogl / jai jars)
│   └── ParaStyler\
│       ├── saxon9pe.jar
│       ├── saxon-license.lic
│       ├── weka-stable-3.6.6.jar
│       ├── asprop30x.arff.randomCommitee_50.model
│       ├── commons-cli-1.2.jar
│       ├── commons-io-2.4.jar
│       ├── guava-10.0.jar
│       ├── log4j-api-2.0-beta8.jar
│       ├── log4j-core-2.0-beta8.jar
│       ├── jsr305-1.3.9.jar
│       └── *.xsl  (XSLT stylesheets)
│
├── config\
│   ├── breakDown.yaml
│   ├── watcher.yaml
│   ├── mAnalyser.yaml
│   ├── mNormalizer.yaml
│   ├── mMerger.yaml
│   ├── dbConfig.yaml
│   ├── dialogConfig.yaml
│   ├── paraStyles.yaml
│   ├── breakdownSequence.json
│   ├── backMatterTitles.json
│   ├── log_config.cfg
│   └── saxon-license.lic
│
└── SupportingFiles\
    ├── SAGE_styles.docx         ← Word template with all SAGE styles
    ├── SAGESTYLES.dotx
    ├── CMSTYLES.dotx
    ├── SAGE_styles.dot
    ├── checkDocRunning.exe      ← COM-safe document status checker
    ├── checkDocRunning.yaml     ← info_path config for status files
    ├── sageJournalInfo.json     ← JID → TLA mapping database
    ├── defaultValue.json
    └── BreakDownLogo.png
```

---

## 4. Batch Files — Purpose & Where to Run

| Batch File | Run On | Purpose |
|---|---|---|
| `start_watcher.bat` | **Watcher Machine** | Starts the watcher loop (monitors hotfolder) |
| `start_watcher_s3.bat` | Watcher Machine (S3) | Starts watcher pointing to S3 location |
| `local_styler.bat` | Any Machine | Run sage-auto-styler.jar on a single .docx manually |
| `local_styler_dev.bat` | Dev Machine | Dev version of local styler (uses source path) |
| `build.bat` | Dev Machine | PyInstaller build (creates BreakDown.exe + watcher.exe) |

**`start_watcher.bat` must remain open** — closing the CMD window stops the watcher.
