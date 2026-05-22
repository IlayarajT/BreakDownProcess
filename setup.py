import sys

from cx_Freeze import Executable, setup

try:
    from cx_Freeze.hooks import get_qt_plugins_paths
except ImportError:
    include_files = []
else:
    include_files = get_qt_plugins_paths("PyQt6", "platforms")
    includefiles = ['config/', 'jar/', 'SupportingFiles/', 'xsl/']
    excludes = ['cx_Freeze', 'pydoc_data', 'setuptools', 'distutils', 'tkinter']
    packages = ['PyQt6', 'PyQt6.sip']
base = None
if sys.platform == "win32":
    base = "Win32GUI"

build_exe_options = {
    "excludes": ["tkinter"],
    "include_files": include_files,
    "zip_include_packages": ["PyQt6"],
}

bdist_mac_options = {
    "bundle_name": "Test",
}

bdist_dmg_options = {
    "volume_label": "TEST",
}

executables = [Executable("main.py", base=None, target_name="test_pyqt6")]

setup(
    name="BreakDown",
    version="1.0.0",
    description="Break Down processes in C and M Digitals",
    options={
        "build_exe": build_exe_options,
        "bdist_mac": bdist_mac_options,
        "bdist_dmg": bdist_dmg_options,
    },
    executables=[Executable("main.py")],
)