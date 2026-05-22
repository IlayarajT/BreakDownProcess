import os
import re
import shutil
import tarfile
import zipfile
import py7zr
from rarfile import RarFile, is_rarfile

class FileHandler:
    @staticmethod
    def normalize_path(path):
        path = re.sub(r"\\", "/", path)
        if not path.endswith("/"):
            path += "/"
        return path

    @staticmethod
    def slash_replace(filepath):
        return filepath.replace("\\", "/")

    def extract_rar(self, rar_path, output_dir):
        with RarFile(rar_path, 'r') as rar:
            try:
                rar.extractall(path=output_dir)
            except Exception:
                try:
                    rar.extractall()
                    cwd_folder = os.path.join(os.getcwd(), os.path.basename(output_dir))
                    if os.path.exists(output_dir):
                        shutil.rmtree(output_dir)
                    shutil.move(cwd_folder, output_dir)
                except Exception as e:
                    print(f"RAR extraction failed: {e}")
        os.remove(rar_path)
        self._flatten_directory(output_dir)
        self._process_nested_archives(output_dir)

    def extract_zip(self, zip_path, output_dir):
        with zipfile.ZipFile(zip_path, 'r') as zfile:
            try:
                zfile.extractall(path=output_dir)
            except Exception:
                try:
                    zfile.extractall()
                    cwd_folder = os.path.join(os.getcwd(), os.path.basename(output_dir))
                    if os.path.exists(output_dir):
                        shutil.rmtree(output_dir)
                    shutil.move(cwd_folder, output_dir)
                except Exception as e:
                    print(f"ZIP extraction failed: {e}")
        os.remove(zip_path)
        self._flatten_directory(output_dir)
        self._process_nested_archives(output_dir)

    def extract_tar(self, tar_path, output_dir):
        with tarfile.open(tar_path, 'r') as tfile:
            tfile.extractall(path=output_dir)
        os.remove(tar_path)
        self._process_nested_archives(output_dir)

    def extract_sevenz(self, file_path, output_dir):
        with py7zr.SevenZipFile(file_path, 'r') as archive:
            archive.extractall(path=output_dir)
        os.remove(file_path)

    def _flatten_directory(self, directory):
        for root, _, files in os.walk(directory):
            for file in files:
                src = os.path.join(root, file)
                dst = os.path.join(directory, file)
                if src != dst:
                    shutil.move(src, dst)

    def _process_nested_archives(self, directory):
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                if file.endswith('.zip'):
                    self.extract_zip(file_path, root)
                elif file.endswith('.tar'):
                    self.extract_tar(file_path, root)
                elif file.endswith('.rar'):
                    self.extract_rar(file_path, root)
                elif file.endswith('.7z'):
                    self.extract_sevenz(file_path, root)

    def is_sevenZfile(self, file_path):
        try:
            with py7zr.SevenZipFile(file_path, 'r') as archive:
                return True
        except py7zr.Bad7zFile:
            return False

    def move_to_process(self, filepath, unique_id, process_folder, error_folder):
        handlers = {
            'dir': self._handle_directory,
            'zip': self._handle_zip,
            'tar': self._handle_tar,
            '7z': self._handle_7z,
            'rar': self._handle_rar
        }
        
        file_type = self._identify_file_type(filepath)
        handler = handlers.get(file_type, lambda *_: (False, None))
        return handler(filepath, unique_id, process_folder, error_folder)

    def _identify_file_type(self, filepath):
        if os.path.isdir(filepath):
            return 'dir'
        elif zipfile.is_zipfile(filepath):
            return 'zip'
        elif tarfile.is_tarfile(filepath):
            return 'tar'
        elif self.is_sevenZfile(filepath):
            return '7z'
        elif is_rarfile(filepath):
            return 'rar'
        return None

    def _handle_directory(self, path, unique_id, process_folder, error_folder):
        folder_name = os.path.basename(path)
        process_path = os.path.join(process_folder, folder_name)
        
        if os.path.exists(process_path):
            shutil.rmtree(process_path)
        
        try:
            shutil.move(path, process_path)
            return True, process_path
        except Exception as e:
            return False, None

    def _handle_archive(self, path, unique_id, process_folder, error_folder, extract_func, file_type):
        base_name = os.path.basename(path).split('.')[0]
        process_path = os.path.join(process_folder, base_name)
        
        try:
            extract_func(path, process_path)
            return True, process_path
        except Exception as e:
            error_path = os.path.join(error_folder, base_name)
            shutil.move(path, error_path)
            return False, error_path

    def _handle_zip(self, path, unique_id, process_folder, error_folder):
        return self._handle_archive(path, unique_id, process_folder, error_folder, 
                                   self.extract_zip, 'zip')

    def _handle_tar(self, path, unique_id, process_folder, error_folder):
        return self._handle_archive(path, unique_id, process_folder, error_folder, 
                                   self.extract_tar, 'tar')

    def _handle_7z(self, path, unique_id, process_folder, error_folder):
        return self._handle_archive(path, unique_id, process_folder, error_folder, 
                                   self.extract_sevenz, '7z')

    def _handle_rar(self, path, unique_id, process_folder, error_folder):
        return self._handle_archive(path, unique_id, process_folder, error_folder, 
                                   self.extract_rar, 'rar')