from cx_Freeze import setup, Executable

setup(
    name="FiltradoContratos",
    version="1.0",
    description="Filtra fechas de contratos SENA en Excel.",
    executables=[Executable("app.py", base="Win32GUI", icon="icono.ico")]
)