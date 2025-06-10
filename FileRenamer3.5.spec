# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['FileRenamer3.5.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['numpy'],  # 添加 numpy 为隐藏导入
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PIL', 'matplotlib', 'scipy', 'cv2', 'tkinter',
        'jupyter', 'IPython', 
        'pkg_resources.py2_warn', 'win32com', 'win32api', 
        'sqlite3', 'PyQt5.QtQml', 'PyQt5.QtNetwork', 'PyQt5.QtWebEngine',
        'PyQt5.QtWebEngineWidgets', 'PyQt5.QtOpenGL', 'PyQt5.QtSql',
        'PyQt5.QtTest', 'PyQt5.QtXml', 'PyQt5.QtXmlPatterns',
        'PyQt5.QtMultimedia', 'PyQt5.QtMultimediaWidgets',
        'PyQt5.QtPositioning', 'PyQt5.QtSensors', 'PyQt5.QtSerialPort'
    ],  # 从 excludes 中移除 numpy
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FileRenamer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico',
    version='file_version_info.txt'
)