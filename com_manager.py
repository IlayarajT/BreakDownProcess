import contextlib
import pythoncom
import win32com.client as win32


class WordApplicationManager:
    """Context manager that creates and cleans up a Word COM application."""

    def __init__(self, visible=True):
        self.visible = visible
        self._word = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self._word = win32.gencache.EnsureDispatch("Word.Application")
        self._word.Visible = self.visible
        return self._word

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self._word is not None:
                self._word.Quit()
        except Exception:
            pass
        finally:
            self._word = None
            pythoncom.CoUninitialize()
        return False


class COMManager:
    @staticmethod
    @contextlib.contextmanager
    def com_context():
        """Initialize COM for the current thread (required in worker threads)."""
        pythoncom.CoInitialize()
        try:
            yield
        finally:
            pythoncom.CoUninitialize()
