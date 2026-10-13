import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], read_only=False)
for sn in ('SOW initial', 'SOW final'):
    ws = wb[sn]
    print(sn, repr(ws['A1'].value))
    for r in range(8, 360):
        c = ws.cell(r, 1)
        if c.font.b and c.fill.fill_type == 'solid' and isinstance(c.value, str) and c.value[:2].rstrip('.').isalpha():
            print('  ', r, repr(c.value))
wb2 = openpyxl.load_workbook(sys.argv[2])
print(repr(wb2['Daily SOW']['B2'].value))
for s in wb2.sheetnames: print(s, repr(wb2[s]['B2'].value))
