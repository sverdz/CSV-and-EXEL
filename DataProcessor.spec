# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec файл для Data Processor GUI
Використовується для збірки standalone EXE файлу
"""

block_cipher = None

a = Analysis(
    ['data_processor_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('data_processor.py', '.')],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'xlsxwriter',
        'tqdm',
        'numpy',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'IPython',
        'notebook',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DataProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Windowed application (без консолі)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Можете додати свою іконку: icon='app.ico'
)
