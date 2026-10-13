import json, sys, re
from openpyxl.utils import coordinate_to_tuple
s = json.load(open(sys.argv[1]))
r0, r1 = int(sys.argv[2]), int(sys.argv[3])
rows = {}
for k, d in s['cells'].items():
    r, c = coordinate_to_tuple(k)
    if r0 <= r <= r1 and ('v' in d or 'fill' in d):
        rows.setdefault(r, []).append((c, k, d))
for r in sorted(rows):
    parts = []
    for c, k, d in sorted(rows[r]):
        v = str(d.get('v', ''))[:28].replace('\n', '/')
        parts.append(f"{k}={v!s}" + (f"[{d['fill']}]" if d.get('fill') else '') + ('B' if d['font'][2] else ''))
    print(r, s['rows'].get(str(r), ''), ' | '.join(parts))
