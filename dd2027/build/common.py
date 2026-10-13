"""Constants and finishing step shared by both AI DD2027 workbooks."""
import os
import zipfile
from datetime import date, datetime

AUTHOR = "Psyroukis Konstantinos"
REV0 = ("Rev00", "PSKO", date(2026, 10, 13))
REV_ROWS = 10            # T4: 10 pre-formatted Rev rows
DATE_FMT = "dd/mm/yyyy"
COST_CARRIER = "8038"    # TS 13/10: 0080 (docking) - 8038 (AL-IDRISI) - XXX (category of cost)

# Yard tariff letters (T3), verbatim from SO - SOW.xlsx, SOW initial / SOW final
CATEGORIES = [
    "A. PILOTAGE, MOORING AND UNMOORING", "B. DRYDOCKING", "C. SERVICES",
    "D. GENERAL", "E. SURFACE PREPARATIONS", "F. PAINT", "G. STEEL", "H. TANKS",
    "I. SEAWATER LINES", "J. HOPPER", "M. PROPELLERS AND TAILSHAFTS",
    "N. ENGINES", "Q. AUXILIARIES", "T. ANCHORS AND CHAINS", "X. MISCELLANEOUS",
]

APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>')


def finish(wb, out, patch=None):
    """Set document properties to the TS, save, blank app.xml (no tool names), optional zip patch."""
    now = datetime.now().replace(microsecond=0)
    p = wb.properties
    p.creator = p.lastModifiedBy = AUTHOR
    p.created = p.modified = now
    p.title = p.subject = p.description = p.keywords = p.category = None
    wb.save(out)
    parts = {"docProps/app.xml": APP_XML.encode()}
    rewrite(out, parts, patch)


def rewrite(path, replace, patch=None):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as src:
        files = {i.filename: src.read(i.filename) for i in src.infolist()}
    files.update(replace)
    if patch:
        patch(files)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        order = ["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]
        for name in order:
            dst.writestr(name, files[name])
    os.replace(tmp, path)
