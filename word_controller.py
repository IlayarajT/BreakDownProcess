import subprocess
import time

import pythoncom
import win32com.client


def _kill_winword():
    try:
        subprocess.run(
            "taskkill /f /im WINWORD.EXE",
            shell=True, stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL, timeout=5
        )
    except Exception:
        pass


class WordSessionController:
    """Manages a Word COM session with automatic restart on failure."""

    def __init__(self, restart_after=5, visible=True, max_restarts=10):
        self.restart_after = restart_after
        self.visible = visible
        self.max_restarts = max_restarts
        self._word = None
        self._file_count = 0
        self._restart_count = 0

    def start(self):
        pythoncom.CoInitialize()
        self._word = win32com.client.DispatchEx("Word.Application")
        self._word.Visible = self.visible
        self._word.DisplayAlerts = 0
        self._word.AutomationSecurity = 3
        self._file_count = 0
        return self._word

    def get_word(self):
        if self._word is None:
            return self.start()
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

    def restart(self):
        if self._restart_count >= self.max_restarts:
            raise RuntimeError(
                f"Word restarted {self._restart_count} times — giving up"
            )
        self.close()
        _kill_winword()
        time.sleep(2)
        self._restart_count += 1
        return self.start()

    def maybe_restart(self):
        self._file_count += 1
        if self._file_count >= self.restart_after:
            self.restart()
