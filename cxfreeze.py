import sys
from cx_Freeze import setup, Executable

include_files = [("assets/", "assets/")]

build_options = {
    "packages": ["pynput"],
    "excludes": [],
    "include_files": include_files,
    "include_msvcr": True,
}

base = "Win32GUI" if sys.platform == "win32" else None

executables = [
    Executable(
        "main.py",
        base=base,
        target_name="SnapToExcel",
        icon="./assets/icon.ico",
        copyright="Abu Ghalib",
        trademarks="SnapToExcel",
    )
]

setup(
    name="SnapToExcel",
    version="0.1.0",
    author="Abu Ghalib (abughalib64@gmail.com)",
    description="Screenshot to Excel File",
    options={"build_exe": build_options},
    executables=executables,
)
