import argparse
import json
import logging
import os
import re
import shutil
import sys
import tarfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from sys import exit

from version import get_version_string, __version__, __build__

import py7zr
from PyQt6 import QtWidgets
from rarfile import is_rarfile

from collectJournalInfo import GetJournalInfo
from createBreakDownJson import CreateBreakDownJson
from dbprocess import DataBase
from loadconfig import getconfig
from mAnalyser import mAnalyzer
from mMerger import DocxMerger
from mNormalizer import ProcessDoc
from mSelect import Ui_Dialog
import pythoncom
from com_manager import COMManager


# ----------------------------------------------------------------------
# gen_py cleanup (COM cache corruption fix)
# ----------------------------------------------------------------------
def cleanup_gen_py():
    """
    Fix win32com.gen_py corruption:
    CLSIDToClassMap / EnsureDispatch errors.
    Best-effort cleanup, behavior-safe.
    """
    try:
        import win32com.client as win32
        gen_py = Path(win32.__gen_path__)
    except Exception:
        return

    if not gen_py.exists():
        return

    for item in gen_py.glob("*"):
        try:
            if item.name == "__init__.py":
                continue  # Keep __init__.py intact
            if item.is_file():
                item.unlink()
            elif item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
        except Exception:
            pass

    # Ensure __init__.py exists so gen_py remains a valid Python package
    init_file = gen_py / "__init__.py"
    if not init_file.exists():
        try:
            init_file.write_text("")
        except Exception:
            pass


# ----------------------------------------------------------------------
# Main controller class
# ----------------------------------------------------------------------
class BreakDown:
    def __init__(self):
        # gen_py must be cleaned before any Word/COM usage
        cleanup_gen_py()

        self.configFolder, self.breakDownConfig = getconfig()
        self.db_process = DataBase()

        self.common_log = self.breakDownConfig["LOGGER"]["ROOT"]
        self.processFolder = self.breakDownConfig["FOLDERS"]["PROCESS"]
        self.errorFolder = self.breakDownConfig["FOLDERS"]["ERROR"]
        self.logFolder = self.breakDownConfig["FOLDERS"]["LOG"]
        self.mergerInput = self.breakDownConfig["FOLDERS"]["MERGER_INPUT"]
        self.mergerError = self.breakDownConfig["FOLDERS"]["MERGER_ERROR"]
        self.ParaStylerInput = self.breakDownConfig["FOLDERS"]["ParaStyler_INPUT"]
        self.ParaStylerError = self.breakDownConfig["FOLDERS"]["ParaStyler_ERROR"]

        for folder in (
            self.processFolder,
            self.errorFolder,
            self.logFolder,
        ):
            Path(folder).mkdir(parents=True, exist_ok=True)

        # ---------------- logging ----------------
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.DEBUG)

        if not self.logger.handlers:
            fh = logging.FileHandler(self.common_log)
            fh.setLevel(logging.ERROR)

            ch = logging.StreamHandler()
            ch.setLevel(logging.INFO)

            fh.setFormatter(
                logging.Formatter(
                    "%(asctime)s: %(name)s - [%(levelname)s] - %(message)s"
                )
            )
            ch.setFormatter(logging.Formatter("[%(levelname)s]: %(message)s"))

            self.logger.addHandler(ch)
            self.logger.addHandler(fh)

        self.logger.info(get_version_string())

        # ---------------- argparse ----------------
        self.parser = argparse.ArgumentParser()
        self.parser.add_argument(
            "-p",
            "--process",
            help="Process Names (watcher, mAnalyser, mSelect, mNormalizer, mMerger)",
            type=str,
        )
        self.parser.add_argument("-f", "--filepath", type=str)
        self.parser.add_argument("-j", "--json", type=str)
        self.parser.add_argument("-jf", "--json_file", type=str)
        self.parser.add_argument("-c", "--customer", type=str)
        self.parser.add_argument("-jid", "--journal_id", type=str)
        self.parser.add_argument("-aid", "--article_id", type=str)
        self.parser.add_argument("-l", "--location", type=str)

        self.args = self.parser.parse_args()

        if len(sys.argv) == 1:
            print(self.parser.format_help())
            exit(0)

        self._validate_args()

        # housekeeping (preserved behavior)
        ext_path = Path("D:/mProjects/mNormalizer/ext")
        ext_path.mkdir(parents=True, exist_ok=True)
        sample = ext_path / "sample.txt"
        if not sample.exists():
            sample.write_text("0")

        CreateBreakDownJson()

    def _validate_args(self):
        if (
            not re.search("(mSelect|watcher|createSageJournalInfo)", self.args.process)
            and self.args.customer is None
        ):
            self.logger.error("Customer name is mandatory to proceed further")
            exit(0)

        if not re.search(
            "(watcher|mAnalyzer|mSelect|mNormalizer|mMerger|createSageJournalInfo)",
            self.args.process,
        ):
            self.logger.error(
                "Process Names should be mAnalyzer|mSelect|mNormalizer|mMerger|createSageJournalInfo"
            )
            exit(0)

        if re.match("mAnalyzer", self.args.process) and self.args.filepath is None:
            self.logger.error("File Path mandatory for mAnalyzer")
            exit(0)

        if re.match("mNormalizer|mMerger", self.args.process) and (
            self.args.json is None and self.args.json_file is None
        ):
            self.logger.error("Json or Json file mandatory for mNormalizer or mMerger")
            exit(0)

        if re.search("watcher", self.args.process) and self.args.location is None:
            self.logger.error('Input Location need to provide ex., -l="HOTFOLDER or S3"')
            exit(0)

        if (
            not re.search("(mSelect|watcher|createSageJournalInfo)", self.args.process)
            and self.args.filepath
            and not os.path.exists(self.args.filepath)
        ):
            self.logger.error(f"{self.args.filepath} is not available")
            exit(0)


# ----------------------------------------------------------------------
# Entry point
# ----------------------------------------------------------------------
if __name__ == "__main__":
    breakdown = BreakDown()
    db_process = DataBase()

    # ---------------- mAnalyzer ----------------
    if re.match("mAnalyzer", breakdown.args.process):
        unique_id = uuid.uuid4().hex
        filepath = breakdown.args.filepath

        is_file = os.path.isfile(filepath)
        is_dir = os.path.isdir(filepath)

        archive_file = False
        if is_file:
            archive_file = (
                zipfile.is_zipfile(filepath)
                or tarfile.is_tarfile(filepath)
                or mAnalyzer.is_sevenZfile(None, filepath)
                or mAnalyzer.my_rarfile(None, filepath)
            )

        date = datetime.today().strftime("%Y-%m-%d")
        time_now = datetime.today().strftime("%H:%M:%S")
        package_id = Path(filepath).stem

        db_process.add_db(package_id, unique_id, date, time_now)

        if not is_dir and not archive_file:
            error_path = os.path.join(
                breakdown.errorFolder, os.path.basename(filepath)
            )
            shutil.move(filepath, error_path)
            breakdown.logger.error(
                "Input should be directory or valid zip/tar file"
            )
            db_process.update_remark(
                unique_id, "Input should be directory or valid zip/tar file"
            )
            exit(0)

        analyser = mAnalyzer()
        details_found, isAuto, _, normalizer_input, word_doc_found = analyser.run_analyser(
            filepath,
            breakdown.args.customer,
            breakdown.errorFolder,
            unique_id,
        )

        if details_found is True:
            if word_doc_found is True:
                if normalizer_input is None:
                    print("[ERROR] Analyser could not prepare normalizer input. "
                          "Skipping normalizer.")
                    db_process.update_db(
                        unique_id, "mAnalyzer", "ERROR", "",
                        "create_info_file returned no normalizer input"
                    )
                    exit(1)

                # ── Single-docx short-metadata package: already in ParaStyler input
                if normalizer_input.get('skip_merger'):
                    print("[INFO] Single docx already placed in ParaStyler input. "
                          "Skipping normalizer and merger.")
                    db_process = DataBase()
                    db_process.update_db(unique_id, "mAnalyzer", "COMPLETED", "", "No error")
                    db_process.update_db(unique_id, "mNormalizer", "SKIPPED", "", "Single docx, no merge needed")
                    db_process.update_db(unique_id, "mMerger", "SKIPPED", "", "Single docx, no merge needed")

                else:
                    normaliser = ProcessDoc()
                    normaliser.kill_running_apps()
                    normalizer_result, normalizer_input = normaliser.process_word_doc(normalizer_input, unique_id)
                    if normalizer_result is True:
                        # ---------- Remove docs if flagged
                        if "remove_docs" in normalizer_input:
                            remove_docs = normalizer_input["remove_docs"]
                            for rem_doc in remove_docs:
                                doc_rem = remove_docs[rem_doc]
                                if os.path.exists(doc_rem):
                                    os.remove(doc_rem)
                                else:
                                    print(f"[INFO] Already removed: {doc_rem}")
                        normaliser.kill_running_apps()
                        if isAuto is True:
                            merger = DocxMerger()
                            merger_result = merger.merge_docx_robust(normalizer_input, unique_id)
                            if merger_result:
                                merger.move_files_to_docs(normalizer_input, unique_id)
                            else:
                                print("[ERROR]: Merging failed. Please check logs and try manual merge.")
                            # merger_result = merger.merge_docx(normalizer_input, unique_id)
                            # if merger_result is False:
                            #     merger_result = merger.merge_in_doc(normalizer_input, unique_id)
                            # if merger_result is True:
                            #     merger.move_files_to_docs(normalizer_input, unique_id)
                            normaliser.kill_running_apps()
                            if merger_result is True:
                                db_process = DataBase()
                                db_process.update_db(unique_id, "mMerger", "COMPLETED", "", "")
                                merger_folder = normalizer_input["folder"]
                                customer = breakdown.args.customer
                                source_folder = re.sub(r"\[CUSTOMER\]", customer, breakdown.mergerInput)
                                dest_folder = re.sub(r"\[CUSTOMER\]", customer, breakdown.ParaStylerInput)
                                source_folder = os.path.join(source_folder, merger_folder)
                                dest_folder = os.path.join(dest_folder, merger_folder)
                                if os.path.exists(dest_folder):
                                    shutil.rmtree(dest_folder, ignore_errors=True)
                                try:
                                    shutil.move(source_folder, dest_folder)
                                except Exception as e:
                                    print(e)
            else:
                print("Doc file(s) not find in package, Please check and proceed manuallay")
                db_process.update_db(unique_id, "mMerger", "ERROR", "", "")
        else:
            pass

    # ---------------- mSelect ----------------
    # elif re.match("mSelect", breakdown.args.process):
    #     COMManager.initialize()
    #     try:
    #         app = QtWidgets.QApplication(sys.argv)
    #         dialog = QtWidgets.QDialog()
    #         ui = Ui_Dialog()
    #         ui.setupUi(dialog)
    #         dialog.show()
    #         pythoncom.CoUninitialize()
    #         sys.exit(app.exec())
    #     finally:
    #         COMManager.uninitialize()
    elif re.match("mSelect", breakdown.args.process):
        from com_manager import COMManager

        with COMManager.com_context():
            app = QtWidgets.QApplication(sys.argv)
            dialog = QtWidgets.QMainWindow()
            ui = Ui_Dialog()
            ui.setupUi(dialog)
            dialog.show()
            sys.exit(app.exec())
    # ---------------- mNormalizer | mMerger ----------------
    elif re.match("mNormalizer|mMerger", breakdown.args.process):
        json_input = breakdown.args.json
        json_file = breakdown.args.json_file

        if isinstance(json_input, str) and json_input is not None:
            json_input = re.sub("'", "\"", json_input)
            try:
                json_object = json.loads(json_input)
                breakdown.logger.error("mNormalizer|mMerger")
            except Exception as e:
                breakdown.logger.error(f"Json error: {e}")
                exit(0)

        elif json_file is not None and os.path.isfile(json_file):
            try:
                with open(json_file, "r", encoding="utf-8") as f:
                    json_data = json.load(f)
                    breakdown.logger.info(json_data)
            except Exception:
                breakdown.logger.error("Json file is not valid")
                exit(0)

        else:
            breakdown.logger.error(
                "Check the parameters for mNormalizer and mMerger"
            )
            exit(0)

    # ---------------- createSageJournalInfo ----------------
    elif re.match("createSageJournalInfo", breakdown.args.process):
        getJrnlInfo = GetJournalInfo()
        getJrnlInfo.create_journal_json()

    else:
        breakdown.logger.info("Process not Mapped")

    # ---------------- dead-code preserved ----------------
    def is_sevenZfile(file_path):
        try:
            with py7zr.SevenZipFile(file_path, "r"):
                return True
        except py7zr.Bad7zFile:
            return False

    def is_rarfile(file_path):
        return is_rarfile(file_path)
