# PyInstaller spec: one-file windowed Windows exe
from PyInstaller.utils.hooks import collect_all

datas, binaries, hidden = [], [], []
for pkg in ("pymupdf", "rapidocr_onnxruntime", "barcode", "acroforge", "reportlab", "pdfplumber", "pdfminer"):
    try:
        d, b, h = collect_all(pkg)
        datas += d; binaries += b; hidden += h
    except Exception:
        pass

a = Analysis(["app.py"], datas=datas, binaries=binaries, hiddenimports=hidden)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name="FormBuilder", console=False, upx=False)
