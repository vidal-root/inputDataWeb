import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.

packages = [
    'seleniumrequests',
    'datetime',
    'time',
    'selenium',
    'xlrd'
]

options = {
    'build_exe': {
        'packages': packages
    },
}

setup(
    name="App Preenche Web",
    version="0.1",
    description="Ele insere dados em paginas web",
    options=options,
    executables=[Executable("app.py", icon="icon.ico")]
)