"""Recalculate a scratch copy with LibreOffice headless and list formula errors per sheet."""
import os, shutil, subprocess, sys, tempfile, openpyxl
src = sys.argv[1]
tmp = tempfile.mkdtemp(dir=sys.argv[2] if len(sys.argv) > 2 else None)
cp = os.path.join(tmp, "in.xlsx"); shutil.copy(src, cp)
subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir",
                os.path.join(tmp, "o"), cp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
wb = openpyxl.load_workbook(os.path.join(tmp, "o", "in.xlsx"), data_only=True)
f = openpyxl.load_workbook(src)
total = 0
for ws in wb:
    errs = [(c.coordinate, c.value) for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith("#")]
    nform = sum(1 for row in f[ws.title].iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    total += len(errs)
    print(f"{ws.title}: {nform} formulas, {len(errs)} errors", errs[:8])
print("TOTAL ERRORS", total)
