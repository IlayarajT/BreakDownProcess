import argparse
import gc
import logging
import os
import signal
import subprocess
import sys
import time
import schedule
import yaml

from getAppPath import getapppath
from loadconfig import getconfig

from utils.resource_monitor import ResourceMonitor
from utils.progress import ProgressBar
from utils.error_logger import log_error
from utils.file_utils import (
    get_input_files,
    is_processable,
    move_to_error,
    terminate_word_excel,
    is_sevenZfile,
    terminate_process
)
from utils.process_runner import run_process_with_timeout


class Watcher:
    def __init__(self):
        self.app_path = getapppath()
        self.configFolder, self.breakDownConfig = getconfig()
        watcher_yaml = os.path.join(self.configFolder, 'config/watcher.yaml')
        with open(watcher_yaml, "r") as stream:
            self.watcher_config = yaml.safe_load(stream)

        self.parser = argparse.ArgumentParser()
        self.parser.add_argument('-p', '--process',
                                 help="Process Names (watcher, mAnalyser, mSelect, mNormalizer, mMerger)", type=str)
        self.parser.add_argument('-f', '--filepath', help="File path to process", type=str)
        self.parser.add_argument('-j', '--json', help="JSON value", type=str)
        self.parser.add_argument('-jf', '--json_file', help="JSON file", type=str)
        self.parser.add_argument('-c', '--customer', help="Customer Name", type=str)
        self.parser.add_argument('-jid', '--journal_id', help="Journal ID", type=str)
        self.parser.add_argument('-aid', '--article_id', help="Article ID", type=str)
        self.parser.add_argument('-l', '--location', help="Input Location", type=str)
        self.args = self.parser.parse_args()
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
        self.error_log_file = 'V:/FOR_BREAKDOWN/LOG/error_log.html'

    def signal_handler(self, sig, frame):
        self.logger.info('Watcher Terminated by User!')
        time.sleep(2)
        sys.exit(0)

    def run_watcher(self, configDetails):
        customers = configDetails[self.args.location]['CUSTOMERS']
        break_down_exe = configDetails['BREAKDOWN_EXE']
        timeout_seconds = 900
        for customer in customers:
            input_path = configDetails[self.args.location][customer]['FOLDERS']['INPUT']
            error_path = configDetails[self.args.location][customer]['FOLDERS']['ERROR']
            input_files = get_input_files(input_path)
            input_files = list(set(input_files))

            if input_files:
                for input_file in input_files:
                    start_time = time.time()
                    if is_processable(input_file):
                        exe_content = f'{break_down_exe} -p="mAnalyzer" -f="{input_file}" -c="{customer}"'
                        self.logger.info(f"Executing: {exe_content}")
                        try:
                            result = run_process_with_timeout(exe_content, timeout_seconds)
                            if not result:
                                self.logger.error(f"Execution timed out or failed: {exe_content}")
                                terminate_word_excel()
                                time.sleep(2)
                                log_error(input_file, start_time, time.time(), self.error_log_file)
                                if os.path.exists(input_file):
                                    move_to_error(input_file, error_path)
                                input_files = []
                                continue
                        except Exception as e:
                            self.logger.error(f"Error executing {exe_content}: {str(e)}")
                            log_error(input_file, start_time, time.time(), self.error_log_file)
                            if os.path.exists(input_file):
                                move_to_error(input_file, error_path)
                            input_files = []
                    else:
                        log_error(input_file, start_time, time.time(), self.error_log_file)
                        if os.path.exists(input_file):
                            move_to_error(input_file, error_path)
                        input_files = []
                    gc.collect()


if __name__ == "__main__":
    watcher = Watcher()
    items = list(range(0, 57))
    l = len(items)
    configDetails = watcher.watcher_config
    pos = 0
    progress_bar = ProgressBar()
    signal.signal(signal.SIGINT, watcher.signal_handler)
    resource_monitor = ResourceMonitor()

    schedule.every(10).seconds.do(watcher.run_watcher, configDetails=configDetails)
    schedule.every(300).seconds.do(resource_monitor.log_resource_usage)

    while True:
        schedule.run_pending()
        pos = (pos + 1) % 51
        progress_bar.print(pos, l, prefix='Progress:', suffix='Watching...', length=50)
        time.sleep(1)
