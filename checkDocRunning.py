import logging
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
import psutil
import yaml
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

from loadconfig import getconfig


STARTED_PREFIX = "STARTED_"
COMPLETED_PREFIX = "COMPLETED_"
ERROR_PREFIX = "ERROR_"
WORD_PROCESS = "WINWORD.EXE"


class MyHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        # Intentionally kept for interface completeness
        pass

    def on_created(self, event):
        file_stem = Path(event.src_path).stem

        if not re.search(rf"^{STARTED_PREFIX}", file_stem, flags=re.I):
            return

        completed_file = re.sub(
            rf"^{STARTED_PREFIX}", COMPLETED_PREFIX, file_stem, flags=re.I
        )
        error_file = re.sub(
            rf"^{STARTED_PREFIX}", ERROR_PREFIX, file_stem, flags=re.I
        )

        docAbr = DocAbserver()
        info_path = docAbr.info_path

        start_info = f"{info_path}{file_stem}.txt"
        completed_info = f"{info_path}{completed_file}.txt"
        error_info = f"{info_path}{error_file}.txt"

        docAbr.startTimer(start_info, completed_info, error_info)


class DocAbserver:
    def __init__(self):
        # app_path retained for backward compatibility (even if unused)
        self.getappath()

        self.configFolder, self.breakDownConfig = getconfig()
        self.analyser_config = self._load_config()

        self.start_range = self.analyser_config["start_range"]
        self.end_range = self.analyser_config["end_range"]
        self.sleep_in_secs = self.analyser_config["sleep_in_secs"]
        self.info_path = self.analyser_config["info_path"]

        Path(self.info_path).mkdir(parents=True, exist_ok=True)

    def getappath(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        if __file__:
            return os.path.dirname(__file__)
        return None

    def _load_config(self):
        config_path = os.path.join(
            self.configFolder, "checkDocRunning.yaml"
        )
        with open(config_path, "r") as stream:
            return yaml.safe_load(stream)

    def startTimer(self, start_info, end_info, error_info):
        if not os.path.exists(start_info):
            return

        file_completed = self._wait_for_completion(start_info, end_info)

        if not file_completed:
            self._handle_timeout(start_info, error_info)

    def _wait_for_completion(self, start_info, end_info):
        for _ in range(self.start_range, self.end_range):
            time.sleep(self.sleep_in_secs)

            # current_time intentionally preserved (even if unused)
            datetime.now().strftime("%H:%M:%S")

            if os.path.exists(end_info):
                self._safe_remove(start_info)
                self._safe_remove(end_info)
                return True

        return False

    def _handle_timeout(self, start_info, error_info):
        if self._is_word_running():
            self._kill_word()
            self._write_error_file(error_info)

        self._safe_remove(start_info)

    def _is_word_running(self):
        return WORD_PROCESS in (proc.name() for proc in psutil.process_iter())

    def _kill_word(self):
        try:
            subprocess.call(
                f"TASKKILL /F /IM {WORD_PROCESS}", shell=True
            )
        except Exception as exc:
            logging.error(exc)

    def _write_error_file(self, error_info):
        try:
            with open(error_info, "w") as file:
                file.write("0")
        except Exception as exc:
            logging.error(exc)

    @staticmethod
    def _safe_remove(path):
        try:
            os.remove(path)
        except Exception:
            time.sleep(2)
            os.remove(path)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    docAbr = DocAbserver()
    info_path = docAbr.info_path

    Path(info_path).mkdir(parents=True, exist_ok=True)

    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path=info_path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        observer.join()
