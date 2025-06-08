from cx_Freeze import setup, Executable
import sys
import os

build_exe_options = {
    "packages": ["pandas", "openpyxl", "tkinter", "datetime"],
    "includes": ["openpyxl.cell", "openpyxl.styles", "openpyxl.utils", "openpyxl.workbook"],
    "include_files": [],  # Añade aquí archivos adicionales si son necesarios
    "excludes": []
}

base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="FiltradoContratos",
    version="1.0",
    description="Filtra fechas de contratos SENA en Excel.",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=base, icon="icono.ico")]
)