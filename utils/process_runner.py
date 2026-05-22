import subprocess
import time
import psutil
from .file_utils import terminate_process

def run_process_with_timeout(command, timeout_seconds):
    try:
        process = subprocess.Popen(command, shell=True, text=True)
        try:
            process.wait(timeout=timeout_seconds)
        except subprocess.TimeoutExpired:
            terminate_process(process.pid)
            process.wait()
            return False
        return process.returncode == 0
    except Exception:
        return False
