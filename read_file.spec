# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['read_file.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'torchvision', 'tensorflow', 'sklearn', 'scipy', 'PIL', 'cv2', 'matplotlib', 'transformers', 'onnxruntime', 'IPython', 'jupyter', 'jedi', 'parso', 'pygments', 'openpyxl', 'fsspec', 'pydantic', 'jinja2', 'regex', 'yt_dlp', 'mutagen', 'brotli', 'secretstorage', 'curl_cffi', 'certifi', 'urllib3', 'requests', 'wcwidth', 'charset_normalizer', 'win32com'],
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
    name='read_file',
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
