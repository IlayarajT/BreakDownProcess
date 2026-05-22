import os
import glob
import shutil

source_root = r"V:\WUNDER"
exclude_folder = r"V:\WUNDER\23-Mar-2026"
destination = r"V:\Wunder\test_files"

os.makedirs(destination, exist_ok=True)

pattern = os.path.join(source_root, "*", "*", "04_SE", "*.docx")

for file_path in glob.glob(pattern):
    if file_path.startswith(exclude_folder):
        continue

    try:
        dest_file = os.path.join(destination, os.path.basename(file_path))
        shutil.copy2(file_path, dest_file)
        print(f"Copied: {file_path}")
    except Exception as e:
        print(f"Error copying {file_path}: {e}")

print("Completed.")