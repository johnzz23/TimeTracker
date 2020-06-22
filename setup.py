from cx_Freeze import setup, Executable

build_exe_options = {"optimize": 2}

setup(name='TimeTracker',
      version = '0.2',
      description ='TimeTracker',
      executables = [Executable("TimeTracker.py", base = "Win32GUI", icon="Buttons/icon.ico")])
