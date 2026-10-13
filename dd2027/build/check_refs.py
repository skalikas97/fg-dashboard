import sys, zipfile, re
for path in sys.argv[1:]:
    z = zipfile.ZipFile(path)
    wb = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)) | {a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)}
    print('==', path.split('/')[-1])
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        t = rmap[rid].lstrip('/'); t = t if t.startswith('xl/') else 'xl/' + t
        with z.open(t) as f:
            head = f.read(6000).decode('utf8', 'ignore')
        m = re.search(r'<pane [^>]*', head)
        y = re.search(r'ySplit="([\d.]+)"', m.group(0)).group(1) if m and 'ySplit' in m.group(0) else '0'
        print(f'   {name!r:40} ySplit={y}')
    names = [n for n in z.namelist() if any(k in n for k in ('threadedComment','externalLink','comments','queryTable','ctrlProp'))]
    print('   parts:', len(names), sorted({re.sub(r'\d+','#',n) for n in names}))
