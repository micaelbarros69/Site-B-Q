from cx_Freeze import setup, Executable

setup(
    name="Anexador de Gpm",
    version="1.0",
    description="Descrição do Executável",
    executables=[Executable("index.py")]
)