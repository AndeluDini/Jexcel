from PyInstaller.utils.hooks import collect_submodules
import os

# Since your spec file is in 'program', the project root is one level up.
BASE_DIR = os.path.abspath(os.path.join(os.getcwd(), '..'))

# Set the path to the ExcelToJson.py file inside the editor folder.
editor_script = os.path.join(BASE_DIR, 'editor', 'ExcelToJson.py')

a = Analysis(
    [editor_script],
    # Add the project root so that both 'editor' and 'jexcel' are visible.
    pathex=[BASE_DIR],
    binaries=[],
    datas=[],
    # Use collect_submodules to automatically include all jexcel submodules.
    hiddenimports=collect_submodules('jexcel'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['hooks/runtime_hook_jexcel.py'],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExcelToJson',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)