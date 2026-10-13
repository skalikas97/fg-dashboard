import sys, collections, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], read_only=True, data_only=True)
ws = wb['Daily SOW']
c = collections.defaultdict(list)
for r in ws.iter_rows(min_row=11, max_row=112, min_col=2, max_col=5, values_only=True):
    if r[2]: c[str(r[2]).strip()].append((r[0], (r[1] or '')[:45]))
for k, v in c.items():
    print(k, len(v)); [print('   ', x) for x in v[:6]]
