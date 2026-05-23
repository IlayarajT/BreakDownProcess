import os
import re
import shutil
import sys
import time
import threading
import logging
import ctypes
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
    # ctypes constants for SetForegroundWindow / ShowWindow
    SW_SHOW = 5
    SW_RESTORE = 9

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
        """Spawn Word visibly so UserForm macros (e.g. Apply_Label) can SetFocus.

        DispatchEx with Visible=False (the old default) created a Word window
        the OS never painted; fm20 UserForms then rejected SetFocus with
        "control is invisible / not enabled / does not accept the focus".
        Launching Visible=True from the start avoids that.

        EnsureDispatch is tried first so the typelib cache is warm; if gen_py
        is corrupt we fall back to DispatchEx.
        """
        with self._dispatch_lock:
            word = None
            try:
                word = client.gencache.EnsureDispatch(self.WORD_PROG_ID)
            except Exception as exc:
                logging.warning(
                    "EnsureDispatch failed (%s); falling back to DispatchEx",
                    exc,
                )
                try:
                    word = client.DispatchEx(self.WORD_PROG_ID)
                except Exception as inner:
                    logging.error("Failed to start Word COM: %s", inner)
                    raise

            try:
                # UserForm macros need a real, painted window. Show it before
                # any document is opened so Word has a foreground window to
                # parent the form against.
                word.Visible = True
                word.DisplayAlerts = 0
                # Do NOT disable ScreenUpdating here — when it's False the
                # UserForm controls inside Apply_Label never paint and the
                # subsequent .SetFocus call inside the macro raises
                # fm20: "Can't move focus to the control...".
                word.ScreenUpdating = True

                # Push the new Word window to the foreground so macros that
                # rely on ActiveWindow / focus see a normal interactive state.
                try:
                    hwnd = int(word.Hwnd) if hasattr(word, "Hwnd") else 0
                    if hwnd:
                        user32 = ctypes.windll.user32
                        user32.ShowWindow(hwnd, self.SW_SHOW)
                        user32.SetForegroundWindow(hwnd)
                except Exception as fg_exc:
                    logging.debug(
                        "Foreground-window promotion skipped: %s", fg_exc
                    )

                # Tiny pause so the OS actually paints the window before
                # the document open / macro Run happens.
                time.sleep(0.3)
            except Exception as exc:
                logging.warning("Word stability setup failed: %s", exc)

            return word

    def _open_with_retry(self, word, docname):
        for attempt in range(OpenDocFile.RETRIES):
            if not self._wait_for_file_ready(docname):
                raise RuntimeError(f"File locked: {docname}")
            try:
                # OpenAndRepair=False — repair-mode can put the document in
                # a state where ActiveX / form controls aren't fully wired,
                # which surfaces as fm20 SetFocus errors inside Apply_Label.
                doc = word.Documents.Open(
                    FileName=os.path.abspath(docname),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Revert=True,
                    Visible=True,
                    OpenAndRepair=False,
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

    def _force_foreground(self, hwnd):
        """Steal foreground to `hwnd` even when our process isn't allowed to.

        Windows blocks SetForegroundWindow from non-foreground processes;
        the standard workaround is to attach our thread's input queue to
        the target window's thread, which inherits its foreground rights.
        """
        if not hwnd:
            return
        try:
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            # Synthesize an Alt keystroke — this gives our process the
            # "user initiated" flag that SetForegroundWindow checks.
            user32.keybd_event(0x12, 0, 0, 0)        # VK_MENU down
            user32.keybd_event(0x12, 0, 0x0002, 0)   # VK_MENU up
            target_thread = user32.GetWindowThreadProcessId(hwnd, None)
            my_thread = kernel32.GetCurrentThreadId()
            attached = False
            if target_thread and target_thread != my_thread:
                attached = bool(user32.AttachThreadInput(
                    my_thread, target_thread, True
                ))
            user32.ShowWindow(hwnd, self.SW_RESTORE)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            if attached:
                user32.AttachThreadInput(my_thread, target_thread, False)
        except Exception as exc:
            logging.debug("Foreground promotion skipped: %s", exc)

    def _activate_for_macro(self, word, doc):
        """Bring document/window to a state where UserForm macros can run.

        UserForm.SetFocus needs (a) the parent Word window foreground, and
        (b) the document Active. Without these, Apply_Label dies with
        fm20: 'control is invisible, not enabled, or of a type that does
        not accept the focus'.
        """
        try:
            doc.Activate()
        except Exception as exc:
            logging.debug("doc.Activate skipped: %s", exc)
        try:
            if word.ActiveWindow is not None:
                word.ActiveWindow.Visible = True
                try:
                    word.ActiveWindow.WindowState = 0  # wdWindowStateNormal
                except Exception:
                    pass
        except Exception as exc:
            logging.debug("ActiveWindow setup skipped: %s", exc)
        try:
            word.Visible = True
            word.Activate()
        except Exception as exc:
            logging.debug("word.Activate skipped: %s", exc)
        try:
            hwnd = int(word.Hwnd) if hasattr(word, "Hwnd") else 0
            self._force_foreground(hwnd)
        except Exception as exc:
            logging.debug("hwnd lookup skipped: %s", exc)
        # Give the OS one paint cycle so SetFocus inside the macro succeeds.
        time.sleep(0.5)

    def _run_macro_with_retry(self, word, doc, macro):
        """Run a macro, retrying with multiple resolution forms.

        Word's macro resolver is finicky: a qualified call like
        ``'TAE_1431025_CLN.docx'!Apply_Label`` looks up the macro in the
        document's own VBA project. When that project doesn't exist Word
        sometimes falls into a recovery path that surfaces an fm20
        UserForm error even when the actual macro lives in a startup
        add-in. So we try unqualified first (which lets Word resolve via
        all loaded templates) and fall back to the qualified form.

        Returns (ok, last_exc). On failure the caller is expected to fall
        back to the Python equivalent (e.g. apply_label_styles).
        """
        last_exc = None
        attempts = (
            macro,                   # unqualified — resolves via loaded templates
            f"'{doc.Name}'!{macro}", # document-qualified — legacy form
        )
        for attempt, call_form in enumerate(attempts, start=1):
            try:
                word.Application.Run(call_form)
                return True, None
            except Exception as exc:
                last_exc = exc
                logging.warning(
                    "Macro %s attempt %s (%s) failed: %s",
                    macro,
                    attempt,
                    call_form,
                    exc,
                )
                # Re-activate before retrying — covers the case where
                # the previous failed call left Word out of foreground.
                self._activate_for_macro(word, doc)
        return False, last_exc

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
                # Re-enable screen updating before macros — UserForm controls
                # (Apply_Label etc.) fail SetFocus when ScreenUpdating=False
                # because their controls are never rendered/painted.
                try:
                    word.ScreenUpdating = True
                except Exception:
                    pass
                # Activate doc + foreground the Word window so UserForm
                # SetFocus inside the macro can succeed.
                self._activate_for_macro(word, doc)
                logging.info("Processing file: %s", docname)
                for macro in list_macros or []:
                    ok, exc = self._run_macro_with_retry(word, doc, macro)
                    if not ok:
                        # Caller (breakDownProcess) runs the Python equivalent
                        # apply_label_styles() right after, so this stays a
                        # non-fatal error — same as the previous behavior.
                        logging.error(
                            "Macro failed (%s): %s — Python fallback will run",
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
            except Exception:
                pass
            # Always quit the Word instance — even if a document is still open
            # (e.g. macro left a modal form in a broken state after an error).
            # doc.Save() was already called above so SaveChanges=False is safe.
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                try:
                    import subprocess as _sp
                    _sp.call(
                        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                        stdout=_sp.DEVNULL, stderr=_sp.DEVNULL,
                    )
                except Exception:
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
