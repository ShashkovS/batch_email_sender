import os
import sys
from cx_Freeze import setup, Executable

os.environ['TCL_LIBRARY'] = r'C:\Users\Serge\Miniconda3\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\Serge\Miniconda3\tcl\tk8.6'

includes = ['os', 'queue', 're', 'smtplib', 'subprocess', 'sys', 'traceback', 'os.path', 'typing', 'keyring', 'openpyxl',
            'email.header', 'email.mime.application', 'email.mime.multipart', 'email.mime.text', 'email.utils',
            'PyQt5.Qt', 'PyQt5.QtWidgetsQSizePolicy', 'QMenuBar', 'QStatusBar', 'PyQt5.QtCore', ]
excludes = []
packages = ['os', 'PyQt5', 'email']
path = []
build_exe_options = {
    'includes': includes,
    'excludes': excludes,
    'packages': packages,
    'path'    : path,
    #'dll_includes': ['msvcr100.dll'],
    'include_msvcr' : True,
    'include_files': [(r'C:\Windows\System32\msvcr100.dll', 'msvcr100.dll'),],
    'zip_include_packages': "*",
    'zip_exclude_packages': ""
}


base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

setup(
    name = "batch_email_sender",
    version = "0.1",
    description = "batch_email_sender",
    options = {'build_exe_options': build_exe_options},
    executables = [Executable("batch_email_sender.py", base=base)]
)