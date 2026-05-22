import os
import time
import logging
from file_handler import FileHandler
from doc_handler import DocHandler
from sage_smart import SageSmartConnector
from short_metadata_processor import ShortMetadataProcessor
from long_metadata_processor import LongMetadataProcessor
from mergeinfo_processor import MergeInfoProcessor
from dbprocess import DataBase
from loadconfig import getconfig
import shutil
import yaml


class MainAnalyzer:
    def __init__(self):
        self.configFolder, self.breakDownConfig = getconfig()
        self._load_config()
        self.db_process = DataBase()
        self.nlogger = logging.getLogger("mAnalyser")
        self.file_handler = FileHandler()
        self.doc_handler = DocHandler()
        self.sage_connector = SageSmartConnector(self.configFolder)
        self._setup_paths()
        self._init_processors()

    def _load_config(self):
        analyzer_yaml = os.path.join(self.configFolder, "config", "mAnalyser.yaml")
        with open(analyzer_yaml, "r") as stream:
            self.analyser_config = yaml.safe_load(stream)

    def _setup_paths(self):
        self.processFolder = self.file_handler.normalize_path(
            self.analyser_config['FOLDERS']['PROCESS']
        )
        self.errorFolder = self.file_handler.normalize_path(
            self.analyser_config['FOLDERS']['ERROR']
        )
        self.outFolder = self.file_handler.normalize_path(
            self.analyser_config['FOLDERS']['OUTPUT']
        )
        self.logFolder = self.file_handler.normalize_path(
            self.analyser_config['LOGGER']['BREAK_DOWN']
        )

    def _init_processors(self):
        self.short_processor = ShortMetadataProcessor(self.sage_connector)
        self.long_processor = LongMetadataProcessor(self.sage_connector)
        self.mergeinfo_processor = MergeInfoProcessor(self.analyser_config)

    def get_file_details(self, customer, file_list, error_folder, unique_id):
        process_folder = file_list.get('process_folder')
        if not process_folder:
            return False, None, None, None

        folder_name = os.path.basename(process_folder).split('.')[0]
        target_error_folder = os.path.join(error_folder, folder_name)

        # Process metadata files
        details_found, journal_id, article_id, ms_no = self._process_metadata_files(
            customer, file_list, process_folder
        )

        if not details_found:
            self._handle_metadata_error(target_error_folder, process_folder, unique_id)
            return False, None, None, None

        # Post-process IDs
        journal_id, article_id = self._post_process_ids(
            customer, journal_id, article_id, ms_no, process_folder
        )

        # Update database
        self.db_process.update_data(customer, ms_no, journal_id, article_id, unique_id)
        return True, journal_id, article_id, ms_no

    def _process_metadata_files(self, customer, file_list, process_folder):
        # Try short metadata first
        if file_list.get("short_metadata"):
            for short_meta in file_list["short_metadata"]:
                details_found, journal_id, article_id, ms_no = self.short_processor.process(
                    short_meta, process_folder, customer
                )
                if details_found:
                    return details_found, journal_id, article_id, ms_no

        # Try long metadata if needed
        if file_list.get("long_metadata"):
            for long_meta in file_list["long_metadata"]:
                details_found, journal_id, article_id, ms_no = self.long_processor.process(
                    long_meta, process_folder, customer
                )
                if details_found:
                    return details_found, journal_id, article_id, ms_no

        return False, None, None, None

    def _handle_metadata_error(self, target_folder, process_folder, unique_id):
        if os.path.exists(target_folder):
            shutil.rmtree(target_folder)
        shutil.move(process_folder, target_folder)
        error_msg = "Metadata file not found. Package moved to error folder."
        self.nlogger.error(error_msg)
        self.db_process.update_remark(unique_id, error_msg)

    def _post_process_ids(self, customer, journal_id, article_id, ms_no, process_folder):
        global jrn_tla
        if not article_id and ms_no and customer == "SAGE":
            _, article_id, jrn_tla = self.sage_connector.get_article_details(
                None, ms_no, None, process_folder
            )
            article_id = article_id or ms_no

        if not jrn_tla and journal_id:
            jrn_tla = self.sage_connector.get_journal_tla(journal_id)

        return jrn_tla or journal_id, article_id

    def run_analyser(self, filepath, customer, errorFolder, unique_id):
        isAuto = False
        nInput = None
        infFile = None
        word_file_found = False
        details_found = False
        process_path = None

        try:
            # Move to process folder and extract if needed
            process_success, process_path = self.file_handler.move_to_process(
                filepath, unique_id, self.processFolder, self.errorFolder
            )
            if not process_success:
                return details_found, isAuto, infFile, nInput, word_file_found

            time.sleep(5)
            file_list = self.doc_handler.get_file_list(process_path)
            details_found, jid, aid, ms_no = self.get_file_details(
                customer, file_list, errorFolder, unique_id
            )

            if not details_found:
                return details_found, isAuto, infFile, nInput, word_file_found

            word_file_found, do_next = self.doc_handler.check_manuscript_format(
                customer, file_list
            )

            if word_file_found and do_next:
                isAuto, infFile, nInput = self.mergeinfo_processor.create_info_file(
                    customer, file_list, jid, aid, ms_no, unique_id, process_path
                )
                self.db_process.update_db(unique_id, "mAnalyzer", "COMPLETED", "", "No error")
        finally:
            if process_path and os.path.exists(process_path):
                try:
                    shutil.rmtree(process_path)
                except Exception:
                    pass

        return details_found, isAuto, infFile, nInput, word_file_found


analyser = MainAnalyzer()
analyser.run_analyser('V:\\FOR_BREAKDOWN\\INPUT\\SAGE\\aut-22-0684-20240604100830.zip', "SAGE", "V:\\FOR_BREAKDOWN\\ERROR", "sfjskfusikk999kfsdf")
