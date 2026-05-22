import os
import re
import time
import threading
from pathlib import Path

import win32com.client as win32
from win32com.client import constants

from WordMacros import WordMacros
from docxPreClean import DocxPreClean

# ----------------------------------------------------------------------
# Global lock for Word COM (Word is NOT thread-safe)
# ----------------------------------------------------------------------
_WORD_LOCK = threading.RLock()


class DocToDocx:
    def __init__(self):
        self.wd = self._init_word()
        self.wd.Visible = True
        time.sleep(2)

    # ------------------------------------------------------------------
    # Word initialization with gen_py auto-recovery
    # ------------------------------------------------------------------
    def _init_word(self):
        try:
            return win32.gencache.EnsureDispatch("Word.Application")
        except (AttributeError, ModuleNotFoundError, ImportError):
            print("[WARN] COM cache corrupted, attempting repair...")
            self._cleanup_gen_py()
            self._invalidate_gen_py_modules()
            try:
                return win32.gencache.EnsureDispatch("Word.Application")
            except (AttributeError, ModuleNotFoundError, ImportError):
                # If EnsureDispatch still fails after cache cleanup,
                # fall back to late-bound Dispatch (no type library needed)
                print("[WARN] EnsureDispatch failed after cache cleanup, falling back to Dispatch")
                return win32.Dispatch("Word.Application")

    def _invalidate_gen_py_modules(self):
        """
        Remove all cached gen_py *sub*-modules from sys.modules so Python
        doesn't reuse the stale/corrupted in-memory copies.
        Keep 'win32com.gen_py' itself so it remains importable as a package.
        """
        import sys
        # Only remove the typelib-specific sub-modules (e.g. win32com.gen_py.00020905-...)
        # Do NOT remove 'win32com.gen_py' itself — that breaks future imports.
        stale_keys = [
            k for k in sys.modules
            if k.startswith('win32com.gen_py.') and k != 'win32com.gen_py'
        ]
        for k in stale_keys:
            del sys.modules[k]

        # Also reset gencache's internal state so it re-scans from disk
        try:
            win32.gencache.__init__()
        except Exception:
            pass

    def _cleanup_gen_py(self):
        """
        Safely remove win32com gen_py cache to fix COM corruption.
        Deletes all contents but preserves the directory with __init__.py
        so Python can still import win32com.gen_py as a package.
        """
        import shutil

        paths_to_try = []

        # Primary: win32com's own reported gen_py path
        try:
            paths_to_try.append(Path(win32.__gen_path__))
        except Exception:
            pass

        # Fallback: standard gen_py location next to the win32com package
        try:
            paths_to_try.append(Path(win32.__file__).parent / "gen_py")
        except Exception:
            pass

        # Fallback: LOCALAPPDATA\Temp\gen_py (common on some setups)
        try:
            local_app = os.environ.get('LOCALAPPDATA', '')
            if local_app:
                paths_to_try.append(Path(local_app) / "Temp" / "gen_py")
        except Exception:
            pass

        for gen_py_path in paths_to_try:
            if not gen_py_path.exists():
                continue
            try:
                shutil.rmtree(gen_py_path, ignore_errors=True)
                gen_py_path.mkdir(parents=True, exist_ok=True)
                # CRITICAL: Recreate __init__.py so gen_py is a valid Python package
                init_file = gen_py_path / "__init__.py"
                init_file.write_text("")
                print(f"[INFO] Cleared COM gen_py cache at: {gen_py_path}")
            except Exception as e:
                print(f"[WARN] Could not clear gen_py cache at {gen_py_path}: {e}")

    # ------------------------------------------------------------------
    # Convert .doc -> .docx (thread-safe, namespace-safe)
    # ------------------------------------------------------------------
    def doc_to_docx(self, docfile):
        with _WORD_LOCK:
            try:
                docx_file = f"{docfile}x"

                # Step 1: Convert DOC → DOCX
                word_doc = self.wd.Documents.Open(docfile)
                word_doc.SaveAs2(docx_file, FileFormat=16)
                word_doc.Close()

                time.sleep(1)

                # Step 2: FORCE Word to normalize XML namespaces
                # This guarantees proper xmlns:w in document.xml
                try:
                    repaired = self.wd.Documents.Open(docx_file)
                    repaired.Save()
                    repaired.Close()
                except Exception:
                    pass  # best-effort repair

                time.sleep(1)

                # Step 3: XML-safe pre-clean
                docx_clean = DocxPreClean()
                docx_clean.preCleanDocx(docx_file)

                return True, docx_file

            except Exception as exc:
                print(f"Failed to Convert: {docfile}")
                print(exc)
                return True, docfile

    # ------------------------------------------------------------------
    # SaveAs DOCX with macro-based lock check (thread-safe)
    # ------------------------------------------------------------------
    def SaveAsDocx(self, path, word):
        with _WORD_LOCK:
            wd_macros = WordMacros()
            doc = word.Documents.Open(path)
            time.sleep(2)

            is_locked = wd_macros.CheckDocLocked(doc)

            if is_locked is False:
                new_file_abs = re.sub(
                    r"doc$", "docx", os.path.abspath(path), flags=re.I
                )

                try:
                    print("[INFO]: SaveAs Docx...")
                    word.ActiveDocument.SaveAs(
                        new_file_abs,
                        FileFormat=constants.wdFormatXMLDocument,
                    )
                    word.ActiveDocument.Close()

                    time.sleep(1)

                    # Force namespace normalization
                    try:
                        repaired = word.Documents.Open(new_file_abs)
                        repaired.Save()
                        repaired.Close()
                    except Exception:
                        pass

                    time.sleep(1)

                    docx_clean = DocxPreClean()
                    docx_clean.preCleanDocx(new_file_abs)

                    return True, new_file_abs, doc

                except Exception as exc:
                    print(f"Failed to Convert: {path}")
                    print(exc)
                    return False, path, doc

            print(f"[ERROR]: Locked file {path}")
            return False, path, doc

# docto = DocToDocx()
# docto.doc_to_docx("V:\\FOR_BREAKDOWN\\PROCESS\\hpi_1416040\\Manuscript Capsular Changes_ HI_minor revisions_1.doc")