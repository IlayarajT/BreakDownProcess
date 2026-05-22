from cx_Freeze import setup, Executable

setup(
    name="my_app",
    version="0.1",
    description="My Spacy Application",
    executables=[Executable("charConverter.py")]
)
