import os
import glob
import shutil
import zipfile
import tarfile
import py7zr
from rarfile import is_rarfile
import psutil

def get_input_files(input_path):
    input_files = []
    for ext in ["*", "*.zip", "*.rar", "*.7z", "*.tar"]:
        input_files.extend(glob.glob(os.path.join(input_path, ext)))
    return input_files

def is_processable(file_path):
    return os.path.isdir(file_path) or any([
        zipfile.is_zipfile(file_path),
        tarfile.is_tarfile(file_path),
        is_sevenZfile(file_path),
        is_rarfile(file_path)
    ])

def move_to_error(file_path, error_path):
    if not os.path.exists(error_path):
        os.makedirs(error_path)
    file_name = os.path.basename(file_path)
    shutil.move(file_path, os.path.join(error_path, file_name))

def is_sevenZfile(file_path):
    try:
        with py7zr.SevenZipFile(file_path, 'r'):
            return True
    except py7zr.Bad7zFile:
        return False

def terminate_word_excel():
    word_processes = ["WINWORD.EXE", "EXCEL.EXE"]
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] in word_processes:
            try:
                proc.terminate()
                proc.wait(timeout=5)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                pass

def terminate_process(proc_pid):
    process = psutil.Process(proc_pid)
    for proc in process.children(recursive=True):
        proc.kill()
    process.kill()
