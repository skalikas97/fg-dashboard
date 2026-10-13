import sys, re, zipfile, openpyxl
p = sys.argv[1]
wb = openpyxl.load_workbook(p)
for ws in wb:
    pane = ws.sheet_view.pane
    print(ws.title, 'ySplit', pane.ySplit if pane else None, 'filter', ws.auto_filter.ref, 'print', ws.print_area, 'dims', ws.dimensions)
print('names', list(wb.defined_names.keys()))
bad = re.compile('[​-‏‪-‮⁠-⁤⁦-⁩﻿­͏᠎   -  　—–]')
hits = [(ws.title, c.coordinate) for ws in wb for row in ws.iter_rows() for c in row if isinstance(c.value, str) and bad.search(c.value)]
print('invisible/dash chars:', hits or 'none')
z = zipfile.ZipFile(p)
blob = b''.join(z.read(n) for n in z.namelist())
print('tool names:', [w for w in (b'openpyxl', b'Openpyxl', b'LibreOffice', b'python') if w in blob] or 'none')
print('theme colour refs in styles:', z.read('xl/styles.xml').count(b'theme='))
