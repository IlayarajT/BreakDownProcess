import os
import re
import shutil
import sys
import time
import threading
import logging
import pywintypes
import win32com.client as client


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


class OpenDocFile:
    RETRIES = 5
    RETRY_WAIT = 2
    WORD_PROG_ID = "Word.Application"
    CALLEE_REJECTED = -2147418111

    _dispatch_lock = threading.Lock()

    def __init__(self):
        self.user_name = os.getlogin()

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _cleanup_gen_py_cache():
        logging.warning("Cleaning corrupted win32com gen_py cache")

        for module in list(sys.modules.keys()):
            if re.match(r"win32com\.gen_py\..+", module):
                del sys.modules[module]

        gen_py_path = os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Temp",
            "gen_py",
        )

        if os.path.isdir(gen_py_path):
            shutil.rmtree(gen_py_path, ignore_errors=True)

    @staticmethod
    def _is_gen_py_broken():
        try:
            from win32com.gen_py import CLSIDToClassMap  # noqa
            return False
        except Exception:
            return True

    # def ensure_word_dispatch(self):
    #     """
    #     Thread-safe Word dispatch with:
    #     - gen_py corruption detection
    #     - forced regeneration
    #     - fallback to active Word instance
    #     """
    #     with self._dispatch_lock:
    #         # 1️⃣ Try existing Word instance first
    #         try:
    #             logging.info("Trying to attach to existing Word instance")
    #             word = client.DispatchEx(self.WORD_PROG_ID)
    #             return word
    #         except Exception:
    #             logging.info("No active Word instance found")
    #
    #         # 2️⃣ Normal EnsureDispatch
    #         try:
    #             word = client.gencache.EnsureDispatch(self.WORD_PROG_ID)
    #
    #             if self._is_gen_py_broken():
    #                 raise AttributeError("gen_py cache corrupted")
    #
    #             return word
    #
    #         except Exception as exc:
    #             logging.error(
    #                 "Word COM cache issue detected: %s", exc
    #             )
    #
    #             # 3️⃣ Hard cache cleanup
    #             self._cleanup_gen_py_cache()
    #
    #             # 4️⃣ Force MakePy regeneration
    #             try:
    #                 logging.info("Forcing Word typelib regeneration")
    #                 client.gencache.GetModuleForProgID(
    #                     self.WORD_PROG_ID
    #                 )
    #             except Exception as regen_exc:
    #                 logging.warning(
    #                     "Typelib regeneration warning: %s", regen_exc
    #                 )
    #
    #             # 5️⃣ Final fallback (bypasses gen_py)
    #             logging.info("Falling back to raw Dispatch")
    #             return client.Dispatch(self.WORD_PROG_ID)

    def ensure_word_dispatch(self):
        with self._dispatch_lock:
            try:
                word = client.DispatchEx(self.WORD_PROG_ID)

                # Stability settings (VERY important)
                word.Visible = False
                word.DisplayAlerts = 0
                word.ScreenUpdating = False

                return word

            except Exception as exc:
                logging.error("Failed to start Word COM: %s", exc)
                raise

    def _open_with_retry(self, word, docname):
        for attempt in range(OpenDocFile.RETRIES):
            if not self._wait_for_file_ready(docname):
                raise RuntimeError(f"File locked: {docname}")
            try:
                doc = word.Documents.Open(
                    FileName=os.path.abspath(docname),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Revert=True,
                    Visible=True,
                    OpenAndRepair=True,
                    NoEncodingDialog=True,
                )

                time.sleep(OpenDocFile.RETRY_WAIT)
                return doc

            except pywintypes.com_error as exc:
                if exc.args and exc.args[0] == OpenDocFile.CALLEE_REJECTED:
                    logging.warning(
                        "Attempt %s failed (callee rejected), retrying",
                        attempt + 1,
                    )
                    time.sleep(OpenDocFile.RETRY_WAIT)
                else:
                    raise
        return None

    def _wait_for_file_ready(self, path, timeout=30):
        start = time.time()
        while time.time() - start < timeout:
            try:
                with open(path, "r+b"):
                    return True
            except (PermissionError, OSError):
                time.sleep(0.5)
        return False

    # ------------------------------------------------------------------
    # Public API (unchanged)
    # ------------------------------------------------------------------
    def processDocFile(self, docname, word_visible, run_macros, auto_close, list_macros):
        have_error = 0
        error_report = ""
        word = self.ensure_word_dispatch()
        word.Visible = word_visible
        try:
            word.Application.DisplayAlerts = False
        except Exception as exc:
            logging.warning(
                "Unable to disable alerts for %s: %s",
                docname,
                exc,
            )

        try:
            doc = self._open_with_retry(word, docname)
            if not doc:
                raise RuntimeError("Document open retries exhausted")

            if run_macros:
                logging.info("Processing file: %s", docname)
                for macro in list_macros or []:
                    try:
                        # word.Application.Run(macro)
                        word.Application.Run(f"'{doc.Name}'!{macro}")
                    except Exception as exc:
                        logging.error(
                            "Macro failed (%s): %s",
                            macro,
                            exc,
                        )
                        error_report += (
                            f"Error running macro {macro}: {exc}\n"
                        )
                        have_error = 1

            try:
                doc.Save()
            except Exception as exc:
                logging.error("Save failed: %s", exc)
                error_report += f"Error saving document: {exc}\n"
                have_error = 1

            if auto_close:
                try:
                    doc.Close(SaveChanges=True)
                except Exception as exc:
                    logging.error("Close failed: %s", exc)
                    error_report += f"Error closing document: {exc}\n"
                    have_error = 1

        except Exception as exc:
            logging.exception(
                "Fatal error processing %s", docname
            )
            error_report += f"Error processing document {docname}: {exc}\n"
            have_error = 1

        finally:
            try:
                word.DisplayAlerts = -1
            except:
                pass
            try:
                if word.Documents.Count == 0:
                    word.Quit()
            except:
                pass
        return have_error, error_report, word

    def openWordDocument(self, docname, word_visible=True):
        word = self.ensure_word_dispatch()
        word.Visible = word_visible

        try:
            doc = self._open_with_retry(word, docname)
            return word if doc else None
        except Exception as exc:
            logging.error(
                "Error opening document %s: %s",
                docname,
                exc,
            )
            word.Quit()
            return None



# openDoc = OpenDocFile()
# openDoc.processDocFile('V:\\FOR_BREAKDOWN\\ParaStyler_INPUT\\SAGE\\HPI_1416040\\HPI_1416040_CLN_AS.docx', True, True, False, None)
