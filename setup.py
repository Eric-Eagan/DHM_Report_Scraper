import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["bs4", "datetime", "os", "requests", "shutil", "subprocess", "sys", "tkcalendar", "tkinter", "xlsxwriter"]}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(  name = "ArrivalReport",
        version = "0.1",
        description = "Create arrival report",
        options = {"build_exe": build_exe_options},
        executables = [Executable("main.py", base=base)])