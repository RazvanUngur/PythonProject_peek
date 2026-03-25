# =============================================================================
# PeekApp.spec
# Rulare:
#   cd C:\Users\razvan.ungur\PycharmProjects\PythonProject_peek\peek_app
#   C:\Users\razvan.ungur\PycharmProjects\PythonProject_peek\.venv\Scripts\pyinstaller.exe PeekApp.spec --clean --noconfirm
# =============================================================================

import os
import sys
import glob
from PyInstaller.utils.hooks import collect_all

CTK_PATH  = r"C:\Users\razvan.ungur\PycharmProjects\PythonProject_peek\.venv\Lib\site-packages\customtkinter"
SITE_PKGS = r"C:\Users\razvan.ungur\PycharmProjects\PythonProject_peek\.venv\Lib\site-packages"

pil_datas, pil_binaries, pil_hiddenimports = collect_all('PIL')

# Colectăm TOATE folderele .dist-info din site-packages
# Pandas verifică la runtime versiunile dependențelor sale via importlib.metadata
# Fără aceste foldere, lansarea eșuează cu "Can't determine version for X"
dist_info_datas = []
for dist_info_dir in glob.glob(os.path.join(SITE_PKGS, '*.dist-info')):
    pkg_name = os.path.basename(dist_info_dir)
    dist_info_datas.append((dist_info_dir, pkg_name))

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=pil_binaries,
    datas=[
        ('assets', 'assets'),
        (CTK_PATH, 'customtkinter'),
    ] + pil_datas + dist_info_datas,
    hiddenimports=pil_hiddenimports + [
        'customtkinter',
        'PIL', 'PIL._tkinter_finder',
        'openpyxl', 'pandas',
        'pytz', 'dateutil', 'dateutil.tz', 'dateutil.tz.tz',
        'winreg', 'json', 'threading', 'calendar',
    ],
    hookspath=[],
    runtime_hooks=['hook_ctk_runtime.py'],
    excludes=[
        'matplotlib', 'scipy', 'unittest', 'xmlrpc',
        'ftplib', 'pydoc', 'tkinter.test',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PeekApp',
    debug=False,
    strip=False,
    upx=True,
    console=False,
    windowed=True,
    icon='assets\\icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='PeekApp',
)
