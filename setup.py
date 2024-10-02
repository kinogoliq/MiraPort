# setup.py

import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["os", "tkinter", "openpyxl", "ttkbootstrap", "PIL"],
    "include_files": ["icons/", "templates/"],
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"
elif sys.platform == "darwin":
    base = "Console"

setup(
    name="ProformaApp",
    version="0.1",
    description="Proforma Application",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon="icons/app_icon.icns")],
)
