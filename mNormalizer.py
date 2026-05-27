import os
import re
import subprocess
import time
import traceback
import shutil
import importlib
from pathlib import Path
from suppNormalizer import SupplementDocxNormalizer
import win32com.client as win32
import yaml
from WordMacros import WordMacros
from dbprocess import DataBase
from doc_to_docx import DocToDocx
from docxPreClean import DocxPreClean
from loadconfig import getconfig
from openDocFile import OpenDocFile
from word_manager import WordManager
from word_controller import WordSessionController
from word_open_new import open_word_document
from word_failures import WordFailureType
from com_utils import repair_com_cache
from largeImageCleaner import DocxImageCleaner



class ProcessDoc:
    # ==============================================================
    # INIT
    # ==============================================================

    def __init__(self):
        self.db_process = DataBase()
        self.configFolder, self.breakDownConfig = getconfig()

        normalizer_yaml = os.path.join(
            self.configFolder, "config", "mNormalizer.yaml"
        )
        with open(normalizer_yaml, "r") as stream:
            self.normalizer_config = yaml.safe_load(stream)

        checkdoc_yaml = os.path.join(
            self.configFolder, "SupportingFiles", "checkDocRunning.yaml"
        )
        with open(checkdoc_yaml, "r") as stream:
            self.checkdoc_config = yaml.safe_load(stream)

        self.info_path = self.checkdoc_config["info_path"]
        Path(self.info_path).mkdir(parents=True, exist_ok=True)

        self.check_doc = os.path.join(
            self.configFolder, "SupportingFiles", "checkDocRunning.exe"
        )

    # ==============================================================
    # UTILITY METHODS
    # ==============================================================

    def _ensure_long_path(self, file_path):
        """Convert to long path format if needed and path exists."""
        if not file_path:
            return file_path

        # Remove existing long path prefix if present
        if file_path.startswith('\\\\?\\'):
            file_path = file_path[4:]

        # Check if path exists and its length
        try:
            if os.path.exists(file_path) and len(file_path) > 240:
                return f'\\\\?\\{file_path}'
        except Exception:
            pass

        return file_path

    def _release_file_locks(self, file_path):
        if not file_path or not os.path.exists(file_path.replace('\\\\?\\', '')):
            return False
        clean_path = file_path.replace('\\\\?\\', '')
        print(f"[INFO] Attempting to release locks on: {clean_path}")
        try:
            temp_path = f"{clean_path}.lockrelease"
            os.rename(clean_path, temp_path)
            time.sleep(0.5)
            os.rename(temp_path, clean_path)
            print("[INFO] File rename trick successful")
            return True
        except Exception:
            pass
        try:
            word_processes = ["WINWORD.EXE", "word.exe"]
            for proc in word_processes:
                subprocess.run(
                    f"taskkill /f /im {proc}",
                    shell=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=5
                )
            time.sleep(2)  # Give time for processes to die
            print("[INFO] Killed Word processes")
        except Exception:
            pass
        try:
            for i in range(5):
                try:
                    with open(clean_path, 'rb') as f:
                        f.read(1)  # Just test if we can read
                        print(f"[INFO] File is readable after {i + 1} attempts")
                        return True
                except Exception:
                    time.sleep(1)
        except Exception:
            pass
        print(f"[WARN] Could not fully release locks on {clean_path}")
        return False

    def _safe_delete(self, file_path):
        """Safely delete a file with retries."""
        if not file_path or not os.path.exists(file_path):
            return True

        for _ in range(5):
            try:
                os.remove(file_path)
                return True
            except PermissionError:
                time.sleep(0.5)
                self._release_file_locks(file_path)
            except Exception:
                break
        return False

    def _wait_for_file(self, file_path, timeout=30, min_size=1024):
        """Wait for file to exist and have minimum size."""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if os.path.exists(file_path) and os.path.getsize(file_path) >= min_size:
                    return True
            except Exception:
                pass
            time.sleep(1)
        return False

    # ==============================================================
    # DOC CONVERSION METHODS
    # ==============================================================

    def _convert_doc_to_docx(self, doc_path, word_app):
        """
        Safely convert .doc to .docx with proper error handling.
        Returns: (success, new_path, temp_path)
        """
        original_path = doc_path
        doc_path = self._ensure_long_path(doc_path)

        if not os.path.exists(doc_path.replace('\\\\?\\', '')):
            return False, None, None

        # Create temp directory for conversion
        temp_dir = os.path.join(os.path.dirname(doc_path), f"temp_conv_{int(time.time())}")
        os.makedirs(temp_dir, exist_ok=True)

        temp_docx = os.path.join(temp_dir, f"{Path(doc_path).stem}.docx")

        max_attempts = 2
        for attempt in range(max_attempts):
            try:
                doctodocx = DocToDocx()
                convert_result, docx_name, _ = doctodocx.SaveAsDocx(
                    doc_path.replace('\\\\?\\', ''),  # Remove prefix for COM
                    word_app
                )

                if convert_result and docx_name and os.path.exists(docx_name):
                    # Move to temp location to avoid conflicts
                    if os.path.exists(temp_docx):
                        self._safe_delete(temp_docx)
                    shutil.move(docx_name, temp_docx)

                    # Wait for file to be ready
                    if self._wait_for_file(temp_docx, timeout=15):
                        return True, temp_docx, temp_dir

            except (AttributeError, ModuleNotFoundError, ImportError) as e:
                err_msg = str(e)
                if attempt < max_attempts - 1 and (
                    'MinorVersion' in err_msg
                    or 'CLSIDToClassMap' in err_msg
                    or 'win32com.gen_py' in err_msg
                    or 'gen_py' in err_msg
                ):
                    print(f"[WARN] COM cache corruption detected (attempt {attempt + 1}): {e}")
                    print("[INFO] Repairing COM cache and retrying...")
                    try:
                        repair_com_cache()
                    except Exception:
                        # Fallback: manually clear the gen_py cache directory
                        try:
                            import win32com
                            gen_py_candidates = []
                            try:
                                gen_py_candidates.append(win32com.__gen_path__)
                            except Exception:
                                pass
                            local_app = os.environ.get('LOCALAPPDATA', '')
                            if local_app:
                                gen_py_candidates.append(os.path.join(local_app, 'Temp', 'gen_py'))
                            gen_py_candidates.append(os.path.join(os.path.dirname(win32com.__file__), 'gen_py'))

                            for candidate in gen_py_candidates:
                                if os.path.isdir(candidate):
                                    print(f"[INFO] Clearing COM cache at: {candidate}")
                                    shutil.rmtree(candidate, ignore_errors=True)
                                    os.makedirs(candidate, exist_ok=True)
                                    # CRITICAL: Recreate __init__.py so gen_py is a valid package
                                    init_file = os.path.join(candidate, "__init__.py")
                                    with open(init_file, "w") as f:
                                        f.write("")
                        except Exception as cleanup_err:
                            print(f"[WARN] Manual cache cleanup failed: {cleanup_err}")

                    # Purge stale sub-modules but keep 'win32com.gen_py' itself
                    try:
                        import sys as _sys
                        stale = [k for k in _sys.modules
                                 if k.startswith('win32com.gen_py.') and k != 'win32com.gen_py']
                        for k in stale:
                            del _sys.modules[k]
                    except Exception:
                        pass

                    time.sleep(2)
                    continue  # Retry the conversion
                else:
                    print(f"[ERROR] Conversion failed for {original_path}: {e}")
                    traceback.print_exc()
                    break

            except Exception as e:
                print(f"[ERROR] Conversion failed for {original_path}: {e}")
                traceback.print_exc()
                break

        # Cleanup on failure
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass

        return False, None, None

    # ==============================================================
    # MAIN PROCESSING METHOD
    # ==============================================================

    def process_word_doc(self, input_details, unique_id):
        """Process Word documents with comprehensive error handling."""

        if input_details is None:
            print("[ERROR] process_word_doc called with None input_details")
            return False, input_details

        selected_files = input_details["selected"]
        customer = input_details["customer"]
        run_macros = self.normalizer_config.get(customer, {}).get("RunMacros", [])

        # ── Pre-clean configuration ────────────────────────────────────────
        # If PreClean.enabled is True, the local sage-auto-styler JAR is
        # invoked AFTER the document is saved to disk. When
        # PreClean.replace_macros is True, the Word macro block is skipped
        # entirely (the JAR does the equivalent work).
        pre_clean_cfg = self.normalizer_config.get(customer, {}).get(
            "PreClean", {}
        ) or {}
        pre_clean_enabled = bool(pre_clean_cfg.get("enabled", False))
        pre_clean_replaces_macros = bool(
            pre_clean_cfg.get("replace_macros", False)
        )
        pre_clean_jar_args = pre_clean_cfg.get("jar_args", ["-pre"]) or ["-pre"]
        pre_clean_timeout = int(pre_clean_cfg.get("timeout", 300))
        pre_clean_jar_name = pre_clean_cfg.get("jar_name", "sage-auto-styler.jar")

        # If pre-clean replaces macros, blank the macro list so the macro
        # block becomes a no-op without further branching.
        if pre_clean_enabled and pre_clean_replaces_macros:
            run_macros = []
            print(f"[INFO] PreClean enabled for {customer}; "
                  f"Word macros will be skipped — JAR will pre-clean each file.")

        normalizer_result = True
        failed_macros = []
        remove_docs = {}
        converted_docx = {}
        temp_dirs = []

        # Track original file mapping for cleanup
        original_files = selected_files.copy()

        # Preserve file order
        sel_dic = {i: {k: v} for i, (k, v) in enumerate(selected_files.items())}

        # Initialize controller with proper settings
        controller = WordSessionController(
            restart_after=3,  # Restart more frequently for stability
            visible=True,
            max_restarts=10
        )
        word = None

        try:
            # ========== INITIALIZE WORD ==========
            print("[INFO] Initializing Word application...")
            for attempt in range(3):
                try:
                    word = controller.start()
                    self.close_all_documents(word)
                    print(f"[SUCCESS] Word initialized (attempt {attempt + 1})")
                    break
                except Exception as e:
                    print(f"[WARN] Word initialization attempt {attempt + 1} failed: {e}")
                    if attempt == 2:
                        raise Exception(f"Failed to initialize Word after 3 attempts: {e}")
                    time.sleep(3)
                    controller.close()

        except Exception as e:
            error_msg = f"Failed to initialize Word: {str(e)}"
            print(f"[FATAL] {error_msg}")
            self.db_process.update_db(
                unique_id, "mNormalizer", "ERROR", "", error_msg
            )
            return False, input_details

        # Process each file
        for key_seq in list(sel_dic.keys()):
            file_dic = sel_dic[key_seq]
            file = list(file_dic.keys())[0]
            original_file_name = file_dic[file]
            fl_name = Path(original_file_name).stem

            print(f"\n{'=' * 60}")
            print(f"Processing: {original_file_name}")
            print(f"{'=' * 60}")

            # Initialize per-file variables
            doc = None
            temp_dir = None
            is_converted_doc = False
            current_file_name = original_file_name
            file_success = True
            macros_result = True

            # ========== STATUS FILE MANAGEMENT ==========
            info_start = os.path.join(self.info_path, f"STARTED_{fl_name}.txt")
            info_error = os.path.join(self.info_path, f"ERROR_{fl_name}.txt")
            info_complete = os.path.join(self.info_path, f"COMPLETED_{fl_name}.txt")

            # Clean up old status files
            for status_file in [info_start, info_error, info_complete]:
                if os.path.exists(status_file):
                    self._safe_delete(status_file)

            # Mark start
            with open(info_start, "w") as fh:
                fh.write("1")

            try:
                # ========== .DOC FILE CONVERSION ==========
                if current_file_name.lower().endswith(".doc"):
                    print(f"[INFO] Converting .doc to .docx: {current_file_name}")

                    # Release locks and wait
                    lock_released = self._release_file_locks(current_file_name)
                    if not lock_released:
                        print(f"[WARNING] File may still be locked: {current_file_name}")
                        # Try to kill Word processes more aggressively
                        self._kill_word_processes()
                        time.sleep(3)

                    # Convert with current Word instance
                    convert_success, converted_path, temp_dir = self._convert_doc_to_docx(
                        current_file_name,
                        word
                    )

                    if convert_success and converted_path and os.path.exists(converted_path):
                        print(f"[SUCCESS] Converted to: {converted_path}")

                        # Store conversion info
                        remove_docs[key_seq] = {file: original_file_name}
                        converted_docx[key_seq] = {f"{file}x": converted_path}

                        # Update current file to work with
                        current_file_name = converted_path
                        is_converted_doc = True
                        temp_dirs.append(temp_dir)

                        # Delete original .doc file if conversion successful
                        try:
                            if os.path.exists(original_file_name):
                                os.remove(original_file_name)
                                print(f"[INFO] Deleted original .doc file: {original_file_name}")
                        except Exception as e:
                            print(f"[WARN] Could not delete original .doc file: {e}")
                    else:
                        raise Exception("DOC to DOCX conversion failed")

                # ========== PRE-PROCESSING ==========
                # Check if file exists before preprocessing
                if not os.path.exists(current_file_name.replace('\\\\?\\', '')):
                    raise Exception(f"File not found: {current_file_name}")

                # Supplement file processing
                if re.search(r"(supplementary|supplement|supp)", current_file_name, re.IGNORECASE):
                    print("[INFO] Processing supplement file...")
                    try:
                        large_table_found, page_count = SupplementDocxNormalizer(current_file_name).process()
                        print(f"[INFO] Supplement processed - Large table: {large_table_found}, Pages: {page_count}")
                    except Exception as e:
                        print(f"[WARN] Supplement processing failed (continuing): {e}")

                # Pre-clean document
                print("[INFO] Applying pre-cleaning...")
                try:
                    DocxPreClean().preCleanDocx(current_file_name)
                    cleaner = DocxImageCleaner(current_file_name)
                    result = cleaner.process()
                    time.sleep(1)
                except Exception as e:
                    print(f"[WARN] Pre-cleaning failed (continuing): {e}")

                # ========== OPEN DOCUMENT ==========
                open_path = self._ensure_long_path(current_file_name)
                print(f"[INFO] Opening document: {open_path}")

                doc = None
                failure = None

                for open_attempt in range(3):
                    try:
                        current_word = controller.get_word()
                        delay_range = 2
                        timeout_sec = 60
                        try:
                            file_size_mb = os.path.getsize(current_file_name) / (1024 * 1024)
                            if file_size_mb > 50:
                                print(f"[INFO] Large file detected: {file_size_mb:.1f}MB")
                                delay_range = 10
                                timeout_sec = 120
                                time.sleep(5)
                                try:
                                    _ = word.Version
                                except Exception:
                                    print("[WARN] Word not responsive, restarting...")
                                    controller.restart()
                                    word = controller.get_word()
                        except Exception:
                            pass
                        doc, failure, word = open_word_document(
                            current_word,
                            open_path,
                            retries=3,
                            delay_between_retries=delay_range,  # Increased delay
                            timeout=timeout_sec  # 2 minute timeout for large files
                        )
                        # doc, failure, word = open_word_document(
                        #     current_word,
                        #     open_path,
                        #     retries=1,  # open_word_document has its own retries
                        #     delay_between_retries=2
                        # )

                        if doc:
                            print(f"[SUCCESS] Document opened on attempt {open_attempt + 1}")
                            break
                        elif failure == WordFailureType.FILE_LOCKED:
                            print(f"[WARN] File locked on attempt {open_attempt + 1}")
                            self._release_file_locks(current_file_name)
                            self._kill_word_processes()
                            time.sleep(3)

                            # Restart Word if locked file persists
                            if open_attempt >= 1:
                                controller.restart()
                                word = controller.get_word()
                    except Exception as e:
                        print(f"[WARN] Open attempt {open_attempt + 1} failed: {e}")
                        if open_attempt == 2:
                            raise
                        time.sleep(2)

                # Check if document opened successfully
                if not doc:
                    error_msg = f"Failed to open document: {failure or 'Unknown error'}"
                    print(f"[ERROR] {error_msg}")

                    # Clean up status files
                    if os.path.exists(info_start):
                        os.remove(info_start)
                    with open(info_error, "w") as fh:
                        fh.write("1")

                    self.db_process.update_db(
                        unique_id, "mNormalizer", "ERROR", "",
                        f"{original_file_name}: {error_msg}"
                    )
                    normalizer_result = False
                    file_success = False
                    continue

                # ========== RUN MACROS ==========
                if run_macros:
                    print("Running Word macros, please wait...")
                    macros_result = True
                    failed_macros_current = []

                    for macro in run_macros:
                        try:
                            print(f"  - Executing: {macro}")
                            word.Application.Run(macro)
                            time.sleep(0.5)  # Reduced delay
                        except Exception as e:
                            error_msg = f"Macro '{macro}' failed: {str(e)}"
                            print(f"[ERROR] {error_msg}")
                            failed_macros_current.append(macro)
                            macros_result = False

                    # If macros failed, restart Word before continuing
                    if not macros_result:
                        failed_macros.extend(failed_macros_current)
                        print("[WARN] Macros failed, restarting Word for next operation...")
                        controller.restart()
                        word = controller.get_word()

                # ========== SAVE DOCUMENT ==========
                save_ok = False

                if doc and macros_result:
                    try:
                        print("[INFO] Saving document...")

                        word.DisplayAlerts = 0
                        doc.Activate()

                        # Activate Word window (non-fatal if it fails — common in automated/headless runs)
                        try:
                            word.Activate()
                        except Exception as activate_err:
                            print(f"[WARN] word.Activate() failed (non-fatal, continuing save): {activate_err}")

                        if doc.ReadOnly:
                            raise Exception("Document is read-only after processing")

                        # F4: Determine the save target path.
                        # doc.Name may be "Document1" when the file was opened via
                        # OpenAndRepair on a network drive — in that case doc.FullName
                        # points to a temp location and doc.Save() would trigger a
                        # Save As dialog.  Always resolve the canonical path from
                        # current_file_name and use SaveAs2 so the save is always
                        # explicit and silent.
                        doc_name = doc.Name  # filename only (no path)
                        is_unnamed = (
                            doc_name.lower().startswith("document")
                            and doc_name.replace("Document", "").replace("document", "").strip().isdigit()
                        )

                        if is_unnamed:
                            print(f"[WARN] doc.Name is '{doc_name}' (unnamed document detected). "
                                  f"Skipping doc.Saved=False and using SaveAs2 with explicit path.")
                            target_path = current_file_name
                        else:
                            # F4: Only force-dirty when the document has a real name;
                            # setting Saved=False on "Document1" guarantees a Save As dialog.
                            doc.Saved = False
                            target_path = doc.FullName  # full path + filename

                        # Always use SaveAs2 with an explicit path — avoids every
                        # variant of the "Save As dialog surprise".
                        word.DisplayAlerts = 0
                        try:
                            doc.SaveAs2(
                                target_path,
                                FileFormat=16  # wdFormatXMLDocument (.docx)
                            )
                            save_ok = True
                            print("[SUCCESS] Document saved via SaveAs2")
                        except Exception as save_err:
                            print(f"[WARN] SaveAs2 failed ({save_err}), falling back to doc.Save()...")
                            doc.Save()
                            save_ok = True
                            print("[SUCCESS] Document saved via doc.Save()")

                    except Exception as e:
                        print(f"[ERROR] Save failed: {e}")

                        # Emergency SaveAs (last resort)
                        try:
                            emergency_path = os.path.join(
                                os.path.dirname(current_file_name),
                                f"{Path(current_file_name).stem}_emergency_{int(time.time())}.docx"
                            )
                            doc.SaveAs(emergency_path)
                            print(f"[WARN] Emergency saved to {emergency_path}")

                            # Replace original file with emergency file to avoid duplicates
                            try:
                                doc.Close(SaveChanges=False)
                                doc = None  # Mark as closed so the close block below skips
                                time.sleep(1)

                                if os.path.exists(current_file_name):
                                    self._safe_delete(current_file_name)
                                    print(f"[INFO] Deleted original file: {current_file_name}")

                                shutil.move(emergency_path, current_file_name)
                                print(f"[SUCCESS] Replaced original with emergency file: {current_file_name}")
                            except Exception as replace_err:
                                print(f"[WARN] Could not replace original with emergency file: {replace_err}")
                                # Update current_file_name so downstream knows the actual path
                                current_file_name = emergency_path

                            save_ok = True
                        except Exception as ee:
                            print(f"[FATAL] Emergency save failed: {ee}")
                            save_ok = False

                # ========== CLOSE DOCUMENT ==========
                if doc:
                    try:
                        # Close without saving if save failed
                        if not save_ok:
                            doc.Close(SaveChanges=False)
                        else:
                            doc.Close()  # Save changes already done
                        print("[INFO] Document closed")
                        doc = None
                    except Exception as e:
                        print(f"[WARN] Failed to close document: {e}")

                # ========== RESULT HANDLING ==========
                if not file_success or not macros_result or not save_ok:
                    # Determine failure reason
                    if not macros_result:
                        failed = f"MACRO_FAILED: {', '.join(failed_macros[-5:])}"  # Last 5 failed macros
                    elif not save_ok:
                        failed = "SAVE_FAILED"
                    else:
                        failed = "PROCESS_FAILED"

                    print(f"[ERROR] Processing failed: {failed}")

                    # Update error status
                    if os.path.exists(info_start):
                        os.remove(info_start)
                    with open(info_error, "w") as fh:
                        fh.write("1")

                    self.db_process.update_db(
                        unique_id, "mNormalizer", "ERROR", "",
                        f"{original_file_name}: {failed}"
                    )
                    normalizer_result = False
                else:
                    print("[SUCCESS] Processing completed successfully")

                    # Move converted .docx from temp dir to original location
                    if is_converted_doc and current_file_name and os.path.exists(current_file_name):
                        # Build the final path: original .doc location with .docx extension
                        final_docx_path = os.path.join(
                            os.path.dirname(original_file_name),
                            f"{Path(original_file_name).stem}.docx"
                        )
                        try:
                            if os.path.exists(final_docx_path):
                                self._safe_delete(final_docx_path)
                            shutil.move(current_file_name, final_docx_path)
                            print(f"[INFO] Moved converted file to: {final_docx_path}")

                            # Update the converted_docx mapping so merger uses the correct path
                            converted_docx[key_seq] = {f"{file}x": final_docx_path}
                            current_file_name = final_docx_path
                        except Exception as move_err:
                            print(f"[WARN] Could not move converted file to root folder: {move_err}")
                            # File stays in temp dir; converted_docx already has the temp path

                    # ========== PRE-CLEAN VIA LOCAL JAR ==========
                    if pre_clean_enabled and current_file_name and os.path.exists(current_file_name):
                        try:
                            from docxManipulator import DocxManipulator
                            print(f"[INFO] Running pre-clean on: {current_file_name}")
                            manipulator = DocxManipulator(jar_name=pre_clean_jar_name)
                            pc_ok, pc_out = manipulator.docx_preclean(
                                current_file_name,
                                jar_args=pre_clean_jar_args,
                                timeout=pre_clean_timeout,
                            )
                            if not pc_ok:
                                print(f"[ERROR] Pre-clean failed for: {current_file_name}")
                                if os.path.exists(info_start):
                                    os.remove(info_start)
                                with open(info_error, "w") as fh:
                                    fh.write("1")
                                self.db_process.update_db(
                                    unique_id, "mNormalizer", "ERROR", "",
                                    f"{original_file_name}: PRECLEAN_FAILED"
                                )
                                normalizer_result = False
                                continue
                            else:
                                # docx_preclean already moved the original to org/ and
                                # renamed _PRE.docx back to the original filename, so
                                # current_file_name is correct and ready for downstream.
                                print(f"[INFO] Pre-clean completed: {pc_out}")
                        except Exception as pc_err:
                            print(f"[WARN] Pre-clean step raised: {pc_err}")
                            traceback.print_exc()

                    # Update success status
                    if os.path.exists(info_start):
                        os.remove(info_start)
                    with open(info_complete, "w") as fh:
                        fh.write("1")

                    self.db_process.update_db(
                        unique_id, "mNormalizer", "COMPLETED", "", "NO ERROR"
                    )

            except Exception as e:
                print(f"[FATAL] Unexpected error processing {original_file_name}: {e}")
                traceback.print_exc()

                # Emergency cleanup
                try:
                    if doc:
                        doc.Close(SaveChanges=False)
                except Exception:
                    pass

                # Mark error
                if os.path.exists(info_start):
                    os.remove(info_start)
                with open(info_error, "w") as fh:
                    fh.write("1")

                self.db_process.update_db(
                    unique_id, "mNormalizer", "ERROR", "",
                    f"{original_file_name}: UNEXPECTED_ERROR - {str(e)[:200]}"
                )
                normalizer_result = False

                # Restart Word for next file
                try:
                    controller.restart()
                    word = controller.get_word()
                except Exception:
                    print("[WARN] Failed to restart Word controller")

            finally:
                # Clean up temp files for converted .doc files
                if is_converted_doc and temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir, ignore_errors=True)
                        print(f"[INFO] Cleaned up temp directory: {temp_dir}")
                    except Exception as e:
                        print(f"[WARN] Failed to clean temp directory: {e}")

                # Restart Word periodically
                controller.maybe_restart()

        # ========== FINAL CLEANUP ==========
        print("\n[INFO] Performing final cleanup...")

        try:
            controller.close()
        except Exception as e:
            print(f"[WARN] Error closing Word controller: {e}")

        # Clean up all temp directories
        for temp_dir in temp_dirs:
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception:
                    pass

        # Update selection dictionary with processed files
        sel_dic = self._rebuild_selection(sel_dic, remove_docs, converted_docx)
        input_details["selected"] = sel_dic["selected"]
        input_details["remove_docs"] = sel_dic["removed"]

        print("[INFO] Cleanup completed")
        return normalizer_result, input_details


    def _kill_word_processes(self):
        """Force kill all Word processes."""
        word_processes = ["WINWORD.EXE", "word.exe"]
        for proc in word_processes:
            try:
                subprocess.run(
                    f"taskkill /f /im {proc}",
                    shell=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=5
                )
            except Exception:
                pass
        time.sleep(2)  # Give processes time to terminate
    # ==============================================================
    # HELPER METHODS
    # ==============================================================

    def close_all_documents(self, word):
        """Close all open documents in Word."""
        try:
            for doc in list(word.Documents):
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            # Give Word time to close documents
            time.sleep(1)
        except Exception as e:
            print(f"[WARN] Error closing documents: {e}")

    @staticmethod
    def _rebuild_selection(sel_dic, remove_docs, converted_docx):
        """Rebuild the selection dictionary after processing."""
        removed = {}
        selected = {}

        # Track removed .doc files
        for k in remove_docs:
            rem = remove_docs[k]
            fname = list(rem.keys())[0]
            removed[fname] = rem[fname]
            sel_dic.pop(k, None)

        # Add converted .docx files
        for k in converted_docx:
            sel_dic[k] = converted_docx[k]

        # Rebuild selected dictionary with proper ordering
        for idx in sorted(sel_dic.keys()):
            f = sel_dic[idx]
            fname = list(f.keys())[0]
            selected[fname] = f[fname]

        return {"selected": selected, "removed": removed}

    # ==============================================================
    # PROCESS CHECK
    # ==============================================================

    @staticmethod
    def process_exists(process_name):
        """Check if a process is running."""
        try:
            output = subprocess.check_output(
                ["TASKLIST", "/FI", f"imagename eq {process_name}"],
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return output.strip().split("\r\n")[-1].lower().startswith(
                process_name.lower()
            )
        except Exception:
            return False

    def kill_running_apps(self):
        """Kill specified processes."""
        kill_processes = self.normalizer_config.get("KillProcess", [])
        for process in kill_processes:
            if self.process_exists(process):
                try:
                    subprocess.call(
                        f"TASKKILL /F /IM {process}",
                        shell=True,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                    print(f"[INFO] Killed process: {process}")
                except Exception as e:
                    print(f"[WARN] Failed to kill {process}: {e}")

# processDoc = ProcessDoc()
# input_det = {'selected': {'main document': r'V:\FOR_BREAKDOWN\BreakDown_DONE\SAGE\TAB_1412846\TAB_1412846_CLN_AS.docx'}, 'customer': 'SAGE', 'folder': 'SHM_1381259', 'process': 'mMerge'}
# test, tst_1 = processDoc.process_word_doc(input_det, '77de416c0369475ea9c70f683136c41c')