"""Add the 'Display Week' scroll bar form control to the 3rd party Planner (T7).

openpyxl cannot write form controls, so the control is added at package level after saving,
with the same definition as in SANTIAGO DOCKING.xlsx (ctrlProp1.xml, vmlDrawing3.vml, drawing1.xml):
horizontal scroll bar linked to '3rd party Planner'!$G$4 (sheet named so every reader resolves it to this tab), min 1, max 100, increment 1, anchored J2:AB4.
The hidden text box of the reference drawing is not carried over.
"""
import re

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SHAPE_ID = 1025

CTRL_PROP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
             'objectType="Scroll" dx="22" fmlaLink="\'3rd party Planner\'!$G$4" horiz="1" max="100" min="1" page="0"/>')

VML = f"""<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype><v:shape id="_x0000_s{SHAPE_ID}" type="#_x0000_t201" style='position:absolute;
  margin-left:789pt;margin-top:46.5pt;width:227.25pt;height:17.25pt;z-index:1'
  strokecolor="windowText [64]" o:insetmode="auto">
  <o:lock v:ext="edit" rotation="t" text="t"/>
  <x:ClientData ObjectType="Scroll">
   <x:SizeWithCells/>
   <x:Anchor>
    9, 10, 1, 13, 27, 11, 3, 2</x:Anchor>
   <x:PrintObject>False</x:PrintObject>
   <x:FmlaLink>'3rd party Planner'!$G$4</x:FmlaLink>
   <x:Val>1</x:Val>
   <x:Min>1</x:Min>
   <x:Max>100</x:Max>
   <x:Inc>1</x:Inc>
   <x:Page>0</x:Page>
   <x:Horiz/>
   <x:Dx>22</x:Dx>
  </x:ClientData>
 </v:shape></xml>"""

ANCHOR = ("<xdr:from><xdr:col>9</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row>"
          "<xdr:rowOff>123825</xdr:rowOff></xdr:from><xdr:to><xdr:col>27</xdr:col><xdr:colOff>104775</xdr:colOff>"
          "<xdr:row>3</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:to>")

DRAWING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
           '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
           f'<xdr:twoCellAnchor editAs="oneCell">{ANCHOR}'
           f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{SHAPE_ID}" name="Scroll Bar 1" hidden="1">'
           '<a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}">'
           f'<a14:compatExt spid="_x0000_s{SHAPE_ID}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>'
           '<xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:miter lim="800000"/>'
           '<a:headEnd/><a:tailEnd/></a:ln></xdr:spPr></xdr:sp><xdr:clientData fPrintsWithSheet="0"/>'
           '</xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>')


def controls_xml(ctrl_rid):
    return ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            '<mc:Choice Requires="x14"><controls>'
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'<mc:Choice Requires="x14"><control shapeId="{SHAPE_ID}" r:id="{ctrl_rid}" name="Scroll Bar 1">'
            '<controlPr defaultSize="0" print="0" autoPict="0"><anchor moveWithCells="1">'
            f'{ANCHOR}</anchor></controlPr></control></mc:Choice></mc:AlternateContent>'
            '</controls></mc:Choice></mc:AlternateContent>')


def inject(files, sheet_index):
    sheet = f"xl/worksheets/sheet{sheet_index}.xml"
    rels = f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels"
    assert sheet in files, sheet
    for p in ("xl/drawings/drawing1.xml", "xl/drawings/vmlDrawing1.vml", "xl/ctrlProps/ctrlProp1.xml"):
        assert p not in files, p
    files["xl/drawings/drawing1.xml"] = DRAWING.encode()
    files["xl/drawings/vmlDrawing1.vml"] = VML.encode()
    files["xl/ctrlProps/ctrlProp1.xml"] = CTRL_PROP.encode()

    rel_xml = files.get(rels, b"").decode() or (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    used = {int(n) for n in re.findall(r'Id="rId(\d+)"', rel_xml)}
    nxt = max(used | {0})
    rid_d, rid_v, rid_c = (f"rId{nxt + i}" for i in (1, 2, 3))
    new = (f'<Relationship Id="{rid_d}" Type="{R_NS}/drawing" Target="../drawings/drawing1.xml"/>'
           f'<Relationship Id="{rid_v}" Type="{R_NS}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
           f'<Relationship Id="{rid_c}" Type="{R_NS}/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>')
    files[rels] = rel_xml.replace("</Relationships>", new + "</Relationships>").encode()

    xml = files[sheet].decode()
    root = re.search(r"<worksheet[^>]*>", xml).group(0)
    if "xmlns:r=" not in root:
        xml = xml.replace(root, root[:-1] + f' xmlns:r="{R_NS}">', 1)
    assert "<drawing " not in xml and "<legacyDrawing" not in xml and "<extLst>" not in xml.split("</sheetData>")[1]
    tail = f'<drawing r:id="{rid_d}"/><legacyDrawing r:id="{rid_v}"/>' + controls_xml(rid_c)
    xml = xml.replace("</worksheet>", tail + "</worksheet>")
    files[sheet] = xml.encode()

    ct = files["[Content_Types].xml"].decode()
    add = ""
    if 'Extension="vml"' not in ct:
        add += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
    add += ('<Override PartName="/xl/drawings/drawing1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            '<Override PartName="/xl/ctrlProps/ctrlProp1.xml" '
            'ContentType="application/vnd.ms-excel.controlproperties+xml"/>')
    files["[Content_Types].xml"] = ct.replace("</Types>", add + "</Types>").encode()
