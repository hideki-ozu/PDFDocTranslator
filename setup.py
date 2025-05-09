import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': ['httplib2'], 'excludes': []}

if sys.platform == "win32":
    base = "Win32GUI"
#base = 'gui'

executables = [
    Executable('main.py', base=base, target_name = 'PDFDocTranslator')
]

setup(name='PDFDocTranslator',
      version = '0.1',
      description = '',
      options = {'build_exe': build_options},
      executables = executables)
