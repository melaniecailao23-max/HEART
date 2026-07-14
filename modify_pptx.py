"""
Modify Cailao_Exp22_PPT_Ch03_ML2_Heart.pptx - SAFE version.
Only makes changes that won't corrupt the file.
Skips: OLE chart embed (requires PowerPoint), SmartArt diagram XML.
Focuses on: Table (Steps 6-9) and simple shapes (Steps 2-3).
"""
import zipfile, shutil, os, io
import xml.etree.ElementTree as ET
import subprocess

for prefix, uri in [
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('p', 'http://schemas.openxmlformats.org/presentationml/2006/main'),
    ('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('a16', 'http://schemas.microsoft.com/office/drawing/2014/main'),
]:
    ET.register_namespace(prefix, uri)

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

def emu(inches): return int(inches * 914400)

INPUT = '/projects/sandbox/HEART/Cailao_Exp22_PPT_Ch03_ML2_Heart.pptx'
TEMP = '/tmp/heart_pptx3'
HEART_IMG = '/projects/sandbox/HEART/Heart.jpg'
EXCEL_FILE = '/projects/sandbox/HEART/HeartData.xlsx'

# Start from original
result = subprocess.run(['git', 'show', 'main:Cailao_Exp22_PPT_Ch03_ML2_Heart.pptx'],
                      capture_output=True, cwd='/projects/sandbox/HEART')
if os.path.exists(TEMP): shutil.rmtree(TEMP)
os.makedirs(TEMP)
with zipfile.ZipFile(io.BytesIO(result.stdout), 'r') as z:
    z.extractall(TEMP)
print("Extracted original PPTX")

def read_xml(path):
    with open(os.path.join(TEMP, path), 'r', encoding='utf-8-sig') as f:
        return ET.fromstring(f.read())

def write_xml(root, path):
    s = ET.tostring(root, encoding='unicode', xml_declaration=False)
    s = '\ufeff<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + s
    with open(os.path.join(TEMP, path), 'w', encoding='utf-8') as f:
        f.write(s)

def write_rels(root, path):
    s = ET.tostring(root, encoding='unicode', xml_declaration=False)
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s
    with open(os.path.join(TEMP, path), 'w', encoding='utf-8') as f:
        f.write(s)

# ============================================================
# STEPS 2-3: Frame shapes on Slide 3
# ============================================================
print("Steps 2-3: Adding shapes on Slide 3...")

items = [
    "Family history of heart disease",
    "High blood pressure",
    "Overweight or obese",
    "Lack of exercise",
    "Use of tobacco products",
    "Diabetes or other diseases",
    "For women, use of birth control pills",
]

root = read_xml('ppt/slides/slide3.xml')
spTree = root.find(f'.//{P}spTree')

shape_w = emu(3.8)
shape_h = emu(0.62)
x_pos = emu(4.3)  # roughly centered
start_y = emu(1.55)
gap = emu(0.68)

for i, item in enumerate(items):
    y_pos = start_y + i * gap
    sp = ET.fromstring(f'''<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvSpPr><p:cNvPr id="{20+i}" name="Frame {20+i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{x_pos}" y="{y_pos}"/><a:ext cx="{shape_w}" cy="{shape_h}"/></a:xfrm>
<a:prstGeom prst="frame"><a:avLst><a:gd name="adj1" fmla="val 6944"/></a:avLst></a:prstGeom>
<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
<a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>
</p:spPr>
<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1100" dirty="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>{item}</a:t></a:r></a:p></p:txBody></p:sp>''')
    spTree.append(sp)

write_xml(root, 'ppt/slides/slide3.xml')
print("  Done")

# ============================================================
# STEPS 6-9: Table on Slide 6
# ============================================================
print("Steps 6-9: Table on Slide 6...")

# Copy Heart.jpg
shutil.copy2(HEART_IMG, os.path.join(TEMP, 'ppt/media/heart.jpg'))

# Update slide 6 rels
ET.register_namespace('', REL_NS)
s6_rels = read_xml('ppt/slides/_rels/slide6.xml.rels')
rel = ET.SubElement(s6_rels, 'Relationship')
rel.set('Id', 'rId2')
rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
rel.set('Target', '../media/heart.jpg')
write_rels(s6_rels, 'ppt/slides/_rels/slide6.xml.rels')

# Table dimensions
th = emu(4.5)
rh = th // 7
tw = emu(10.0)
c1w = emu(3.5)
c2w = emu(3.0)
c3w = tw - c1w - c2w
tx = emu(1.1)
ty = emu(1.8)

data = [
    ("Men", "", "Women"),
    ("Chest pain/discomfort", "", "Chest pressure"),
    ("Rapid or irregular heartbeat", "", "Fatigue for several days"),
    ("Dizzy or light-headed", "", "Anxiety and sleep disturbances"),
    ("Cold sweat", "", "Back, neck, arm, or jaw pain"),
    ("Stomach discomfort/indigestion", "", "Nausea"),
    ("Shortness of breath", "", "Shortness of breath"),
]

def cell(text, **attrs):
    extra = ''.join(f' {k}="{v}"' for k,v in attrs.items())
    c = f'<a:tc{extra}><a:txBody><a:bodyPr/><a:lstStyle/>'
    if text:
        c += f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" dirty="0"/><a:t>{text}</a:t></a:r></a:p>'
    else:
        c += '<a:p><a:endParaRPr lang="en-US" sz="1200"/></a:p>'
    c += '</a:txBody><a:tcPr anchor="ctr"'
    c += '/></a:tc>'
    return c

def img_cell(**attrs):
    extra = ''.join(f' {k}="{v}"' for k,v in attrs.items())
    c = f'<a:tc{extra}><a:txBody><a:bodyPr/><a:lstStyle/>'
    c += '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    c += '</a:txBody><a:tcPr anchor="ctr">'
    c += '<a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
    c += '</a:tcPr></a:tc>'
    return c

def vmerge_cell():
    return '<a:tc vMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>'

tbl = f'''<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{{5940675A-B579-460E-94D1-54222C63F5DA}}</a:tableStyleId></a:tblPr>
<a:tblGrid><a:gridCol w="{c1w}"/><a:gridCol w="{c2w}"/><a:gridCol w="{c3w}"/></a:tblGrid>'''

for ri, (t1, t2, t3) in enumerate(data):
    tbl += f'<a:tr h="{rh}">'
    tbl += cell(t1)
    if ri == 0:
        tbl += cell("")
    elif ri == 1:
        tbl += img_cell(rowSpan="6")
    else:
        tbl += vmerge_cell()
    tbl += cell(t3)
    tbl += '</a:tr>'

tbl += '</a:tbl>'

gf = f'''<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvGraphicFramePr><p:cNvPr id="10" name="Table 9"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvGraphicFramePr>
<p:xfrm><a:off x="{tx}" y="{ty}"/><a:ext cx="{tw}" cy="{th}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">{tbl}</a:graphicData></a:graphic>
</p:graphicFrame>'''

root = read_xml('ppt/slides/slide6.xml')
spTree = root.find(f'.//{P}spTree')

# Remove content placeholder
for sp in list(spTree):
    nvPr = sp.find(f'.//{P}nvPr')
    if nvPr is not None:
        ph = nvPr.find(f'{P}ph')
        if ph is not None and ph.get('idx') == '1' and ph.get('type') is None:
            spTree.remove(sp)
            break

spTree.append(ET.fromstring(gf))
write_xml(root, 'ppt/slides/slide6.xml')
print("  Done")

# ============================================================
# Content_Types update (just for jpg)
# ============================================================
print("Content_Types...")
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('', CT)
ct = read_xml('[Content_Types].xml')
has_jpg = any(d.get('Extension') == 'jpg' for d in ct)
if not has_jpg:
    d = ET.SubElement(ct, f'{{{CT}}}Default')
    d.set('Extension', 'jpg')
    d.set('ContentType', 'image/jpeg')

# Do NOT add xlsx or embedding - skip the OLE embed to avoid corruption
s = ET.tostring(ct, encoding='unicode', xml_declaration=False)
s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s
with open(os.path.join(TEMP, '[Content_Types].xml'), 'w', encoding='utf-8') as f:
    f.write(s)
print("  Done")

# ============================================================
# FINAL: Repackage
# ============================================================
print("Repackaging...")
if os.path.exists(INPUT): os.remove(INPUT)
with zipfile.ZipFile(INPUT, 'w', zipfile.ZIP_DEFLATED) as zf:
    for dirpath, dirs, files in os.walk(TEMP):
        for f in files:
            fp = os.path.join(dirpath, f)
            arcname = os.path.relpath(fp, TEMP)
            zf.write(fp, arcname)
print(f"Saved: {INPUT} ({os.path.getsize(INPUT)} bytes)")
print("\nDONE! Note: Steps 4-5 (chart embed) must be done manually in PowerPoint.")
