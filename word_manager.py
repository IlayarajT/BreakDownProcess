import pythoncom
import win32com.client


class WordManager:
    """Thin Word COM application manager (legacy import shim)."""

    def __init__(self, visible=True):
        self.visible = visible
        self._word = None

    def open(self):
        pythoncom.CoInitialize()
        self._word = win32com.client.DispatchEx("Word.Application")
        self._word.Visible = self.visible
        self._word.DisplayAlerts = 0
        return self._word

    def close(self):
        try:
            if self._word is not None:
                self._word.Quit()
        except Exception:
            pass
        finally:
            self._word = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
