from cx_Freeze import setup, Executable
import sys
import os
import numpy
import pandas

build_exe_options = {
    "packages": [
        "pandas", 
        "numpy",
        "numpy.core",
        "numpy.core._methods",  # Agregado explícitamente
        "numpy.lib",
        "openpyxl", 
        "tkinter", 
        "datetime",
        "pytz",
        "dateutil",
    ],
    "includes": [
        "openpyxl.cell", 
        "openpyxl.styles", 
        "openpyxl.utils", 
        "openpyxl.workbook",
        "numpy.core._methods",
        "numpy.lib.format",
        "pandas._libs.tslibs.timedeltas",
        "pandas._libs.tslibs.timestamps",
    ],
    "include_files": [
        # Incluir explícitamente los directorios de numpy y pandas
        (numpy.__path__[0], "numpy"),
        (pandas.__path__[0], "pandas"),
    ],
    "excludes": [],
    "include_msvcr": True,
}

base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="FiltradoContratos",
    version="1.0",
    description="Filtra fechas de contratos en Excel.",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=base, icon="icono.ico")]
)