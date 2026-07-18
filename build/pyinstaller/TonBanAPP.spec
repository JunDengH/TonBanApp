# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path


project_root = Path(SPECPATH).resolve().parents[1]
ocr_root = project_root / 'vendor' / 'tesseract'
ocr_binaries = [
    (str(ocr_root / 'tesseract.exe'), 'vendor\\tesseract'),
    *[(str(path), 'vendor\\tesseract') for path in ocr_root.glob('*.dll')],
]
ocr_datas = [
    (str(ocr_root / 'tessdata' / 'chi_sim.traineddata'), 'vendor\\tesseract\\tessdata'),
    (str(ocr_root / 'tessdata' / 'eng.traineddata'), 'vendor\\tesseract\\tessdata'),
    (str(ocr_root / 'tessdata' / 'configs'), 'vendor\\tesseract\\tessdata\\configs'),
    (str(ocr_root / 'tessdata' / 'tessconfigs'), 'vendor\\tesseract\\tessdata\\tessconfigs'),
    (str(ocr_root / 'LICENSE-TESSERACT.txt'), 'vendor\\tesseract'),
    (str(ocr_root / 'LICENSE-TESSDATA.txt'), 'vendor\\tesseract'),
    (str(ocr_root / 'README-UPSTREAM.md'), 'vendor\\tesseract'),
    (str(ocr_root / 'VENDOR.md'), 'vendor\\tesseract'),
]

a = Analysis(
    ['..\\..\\main.py'],
    pathex=[],
    binaries=ocr_binaries,
    # 发布 exe 不内置 data/config.json；开发期本地配置继续保留在项目目录。
    datas=ocr_datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name='TonBanAPP',
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
)
