# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

block_cipher = None

TESSERACT_DIR = Path(r"C:\Program Files\Tesseract-OCR")

# 收集 Tesseract 目錄下所有 .exe 和 .dll
tess_binaries = []
for f in TESSERACT_DIR.iterdir():
    if f.is_file() and f.suffix.lower() in (".exe", ".dll", ".jar"):
        tess_binaries.append((str(f), "tesseract"))

# tessdata（語言資料）
tessdata_files = []
for f in (TESSERACT_DIR / "tessdata").rglob("*"):
    if f.is_file():
        rel = f.relative_to(TESSERACT_DIR)
        tessdata_files.append((str(f), str(Path("tesseract") / rel.parent)))

a = Analysis(
    ["gui.py"],
    pathex=["."],
    binaries=tess_binaries,
    datas=[
        *tessdata_files,
    ],
    hiddenimports=[
        "pdfplumber",
        "fitz",
        "pytesseract",
        "cv2",
        "pandas",
        "img2table",
        "docx2pdf",
        "win32com",
        "win32com.client",
        "pywintypes",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "uvicorn", "fastapi", "starlette", "anyio",
        "multipart", "httpx", "websockets",
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
    [],
    exclude_binaries=True,
    name="PDF轉Word",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="PDF轉Word",
)
