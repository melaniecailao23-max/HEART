"""
Modify Cailao_Exp22_PPT_Ch03_ML2_Heart.pptx per assignment instructions.
Works directly with PPTX ZIP/XML structure.
"""
import zipfile, shutil, os, io
import xml.etree.ElementTree as ET

# Register namespaces
for prefix, uri in [
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('p', 'http://schemas.openxmlformats.org/presentationml/2006/main'),
    ('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main'),
    ('dgm', 'http://schemas.openxmlformats.org/drawingml/2006/diagram'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
]:
    ET.register_namespace(prefix, uri)

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

def emu(inches): return int(inches * 914400)

INPUT = '/projects/sandbox/HEART/Cailao_Exp22_PPT_Ch03_ML2_Heart.pptx'
TEMP = '/tmp/heart_pptx'
HEART_IMG = '/projects/sandbox/HEART/Heart.jpg'
EXCEL_FILE = '/projects/sandbox/HEART/HeartData.xlsx'

if os.path.exists(TEMP): shutil.rmtree(TEMP)
os.makedirs(TEMP)
with zipfile.ZipFile(INPUT, 'r') as z: z.extractall(TEMP)
print("Extracted PPTX")

def read_xml(path):
    with open(os.path.join(TEMP, path), 'r', encoding='utf-8-sig') as f:
        return ET.fromstring(f.read())

def write_xml(root, path):
    s = ET.tostring(root, encoding='unicode', xml_declaration=False)
    s = '\ufeff<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + s
    with open(os.path.join(TEMP, path), 'w', encoding='utf-8') as f:
        f.write(s)

def write_xml_no_bom(root, path):
    s = ET.tostring(root, encoding='unicode', xml_declaration=False)
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s
    with open(os.path.join(TEMP, path), 'w', encoding='utf-8') as f:
        f.write(s)


# ============================================================
# STEP 2: Insert Basic Block List SmartArt on Slide 3
# ============================================================
print("Step 2: Inserting SmartArt on Slide 3...")

# SmartArt items
items = [
    "Family history of heart disease",
    "High blood pressure",
    "Overweight or obese",
    "Lack of exercise",
    "Use of tobacco products",
    "Diabetes or other diseases",
    "For women, use of birth control pills",
]

# SmartArt in OOXML requires diagram data files.
# We'll create the SmartArt as a graphicFrame with diagram references.
# Need: diagData, diagColors, diagStyle, diagLayout + relationships

# Create diagram data XML
os.makedirs(os.path.join(TEMP, 'ppt/diagrams'), exist_ok=True)

# Generate unique model IDs
import uuid
model_id = str(uuid.uuid4()).upper()

# Diagram data (dgm:dataModel)
data_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
data_xml += '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
data_xml += '<dgm:ptLst>'
# Root point
data_xml += f'<dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
# Data points for each item
for i, item in enumerate(items):
    mid = str(i + 1)
    data_xml += f'<dgm:pt modelId="{mid}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{item}</a:t></a:r></a:p></dgm:t></dgm:pt>'
data_xml += '</dgm:ptLst>'
data_xml += '<dgm:cxnLst>'
# Connections from root to each item
for i in range(len(items)):
    mid = str(i + 1)
    data_xml += f'<dgm:cxn modelId="{10+i}" srcId="0" destId="{mid}" type="parOf" sibTransId="{20+i}"/>'
data_xml += '</dgm:cxnLst>'
data_xml += '<dgm:bg/><dgm:whole/></dgm:dataModel>'

with open(os.path.join(TEMP, 'ppt/diagrams/data1.xml'), 'w', encoding='utf-8') as f:
    f.write(data_xml)

# Diagram colors
colors_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
colors_xml += '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2">'
colors_xml += '<dgm:title lang="" val=""/><dgm:desc lang="" val=""/><dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst>'
colors_xml += '</dgm:colorsDef>'

with open(os.path.join(TEMP, 'ppt/diagrams/colors1.xml'), 'w', encoding='utf-8') as f:
    f.write(colors_xml)

# Diagram style
style_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
style_xml += '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1">'
style_xml += '<dgm:title lang="" val=""/><dgm:desc lang="" val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>'
style_xml += '</dgm:styleDef>'

with open(os.path.join(TEMP, 'ppt/diagrams/style1.xml'), 'w', encoding='utf-8') as f:
    f.write(style_xml)

# Diagram layout - Basic Block List = urn:microsoft.com/office/officeart/2005/8/layout/vList5
layout_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
layout_xml += '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
layout_xml += '</dgm:layoutDef>'

with open(os.path.join(TEMP, 'ppt/diagrams/layout1.xml'), 'w', encoding='utf-8') as f:
    f.write(layout_xml)

# Create diagram rels
os.makedirs(os.path.join(TEMP, 'ppt/diagrams/_rels'), exist_ok=True)

# Add relationships for slide 3
ET.register_namespace('', REL_NS)
slide3_rels = read_xml('ppt/slides/_rels/slide3.xml.rels')
# Add diagram relationships
next_rid = 2
for rel_type, target in [
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', '../diagrams/data1.xml'),
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors', '../diagrams/colors1.xml'),
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramStyle', '../diagrams/style1.xml'),
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout', '../diagrams/layout1.xml'),
]:
    rel = ET.SubElement(slide3_rels, 'Relationship')
    rel.set('Id', f'rId{next_rid}')
    rel.set('Type', rel_type)
    rel.set('Target', target)
    next_rid += 1

write_xml_no_bom(slide3_rels, 'ppt/slides/_rels/slide3.xml.rels')

# Now add the SmartArt graphic frame to Slide 3
root = read_xml('ppt/slides/slide3.xml')
spTree = root.find(f'.//{P}spTree')

# The SmartArt graphicFrame with Frame shapes (Step 3 combined)
# Width 3.8" for shapes, font Black Text 1
# Since we can't use actual SmartArt rendering, we'll create individual shapes
# that represent the Basic Block List with Frame shape geometry

# Create shapes directly as a group of Frame shapes (combining Steps 2 & 3)
# Frame = prstGeom "frame"
# Position them in a vertical list layout
shape_width = emu(3.8)
shape_height = emu(0.7)  # approximate height for each frame
start_x = emu(1.5)  # centered-ish
start_y = emu(1.5)
spacing = emu(0.05)

for i, item in enumerate(items):
    sp_id = 20 + i
    y_pos = start_y + i * (shape_height + spacing)
    
    sp_xml = f'''<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvSpPr><p:cNvPr id="{sp_id}" name="Frame {sp_id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="{start_x}" y="{y_pos}"/><a:ext cx="{shape_width}" cy="{shape_height}"/></a:xfrm>
<a:prstGeom prst="frame"><a:avLst><a:gd name="adj1" fmla="val 6250"/></a:avLst></a:prstGeom>
<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
<a:ln><a:noFill/></a:ln>
</p:spPr>
<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>{item}</a:t></a:r></a:p></p:txBody></p:sp>'''
    
    sp_elem = ET.fromstring(sp_xml)
    spTree.append(sp_elem)

write_xml(root, 'ppt/slides/slide3.xml')
print("  Done (Steps 2 & 3 combined)")


# ============================================================
# STEP 4 & 5: Embed Excel chart on Slide 4
# Width 9.1", H pos 1.3", V pos 1.9"
# ============================================================
print("Steps 4-5: Embedding chart on Slide 4...")

# For an embedded OLE object (Paste Special as embedded object),
# we need to embed the Excel file and create an OLE object frame.
# Copy the Excel file into the pptx as an embedded object
os.makedirs(os.path.join(TEMP, 'ppt/embeddings'), exist_ok=True)
shutil.copy2(EXCEL_FILE, os.path.join(TEMP, 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'))

# Also need a preview image - we'll create a simple placeholder
# For now, create the graphicFrame that references the OLE object

# Update slide 4 rels
ET.register_namespace('', REL_NS)
slide4_rels = read_xml('ppt/slides/_rels/slide4.xml.rels')

# Add OLE object relationship
ole_rel = ET.SubElement(slide4_rels, 'Relationship')
ole_rel.set('Id', 'rId2')
ole_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package')
ole_rel.set('Target', '../embeddings/Microsoft_Excel_Worksheet1.xlsx')

write_xml_no_bom(slide4_rels, 'ppt/slides/_rels/slide4.xml.rels')

# Add the OLE object frame to slide 4
root = read_xml('ppt/slides/slide4.xml')
spTree = root.find(f'.//{P}spTree')

# Chart dimensions and position
chart_width = emu(9.1)
chart_height = emu(4.5)  # reasonable height for a chart
chart_x = emu(1.3)
chart_y = emu(1.9)

# Create graphicFrame for embedded Excel chart
gf_xml = f'''<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvGraphicFramePr>
<p:cNvPr id="4" name="Object 3"/>
<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
<p:nvPr/>
</p:nvGraphicFramePr>
<p:xfrm><a:off x="{chart_x}" y="{chart_y}"/><a:ext cx="{chart_width}" cy="{chart_height}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
<p:oleObj spid="_x0000_s1026" name="Worksheet" r:id="rId2" imgW="{chart_width}" imgH="{chart_height}" progId="Excel.Sheet.12">
<p:embed/>
</p:oleObj>
</a:graphicData></a:graphic>
</p:graphicFrame>'''

gf_elem = ET.fromstring(gf_xml)
spTree.append(gf_elem)

write_xml(root, 'ppt/slides/slide4.xml')
print("  Done")


# ============================================================
# STEPS 6-9: Table on Slide 6
# ============================================================
print("Steps 6-9: Creating table on Slide 6...")

# Table data (after inserting middle column and merging)
# Final structure: 3 columns x 7 rows
# Col 1: Men data, Col 2: merged (Heart.jpg), Col 3: Women data
# Row 1: "Men" | (header for merged col - empty or merged) | "Women"

# But first, per instructions:
# Step 6: Create 2-col, 7-row table
# Step 7: Insert column to RIGHT of first column, merge rows 2-7 in new column
# Step 8: Style, height, alignment
# Step 9: Heart.jpg in merged column

# So final table: 3 columns, 7 rows
# Col 1: Men symptoms
# Col 2: Merged cell rows 2-7 with Heart.jpg background
# Col 3: Women symptoms

table_data = [
    ["Men", "", "Women"],
    ["Chest pain/discomfort", "", "Chest pressure"],
    ["Rapid or irregular heartbeat", "", "Fatigue for several days"],
    ["Dizzy or light-headed", "", "Anxiety and sleep disturbances"],
    ["Cold sweat", "", "Back, neck, arm, or jaw pain"],
    ["Stomach discomfort/indigestion", "", "Nausea"],
    ["Shortness of breath", "", "Shortness of breath"],
]

# Table dimensions
table_height = emu(4.5)
row_height = table_height // 7
# Slide width is 12192000 EMU (13.33"), use reasonable table width
table_width = emu(10.0)
col1_width = emu(3.8)
col2_width = emu(2.4)  # merged column for image
col3_width = table_width - col1_width - col2_width

# Position table in content area
table_x = emu(1.1)
table_y = emu(1.6)

# Build the table XML
# Light Style 2 = {5940675A-B579-460E-94D1-54222C63F5DA} (standard GUID)
# This applies a light formatting with accent colors

# Add Heart.jpg to pptx media
shutil.copy2(HEART_IMG, os.path.join(TEMP, 'ppt/media/heart.jpg'))

# Update slide 6 rels for the image
ET.register_namespace('', REL_NS)
slide6_rels = read_xml('ppt/slides/_rels/slide6.xml.rels')
img_rel = ET.SubElement(slide6_rels, 'Relationship')
img_rel.set('Id', 'rId2')
img_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
img_rel.set('Target', '../media/heart.jpg')
write_xml_no_bom(slide6_rels, 'ppt/slides/_rels/slide6.xml.rels')

# Build table XML
def build_cell(text, is_merged_vert=False, merge_count=0, has_image=False, img_rid=''):
    """Build a table cell XML string."""
    cell = '<a:tc'
    if merge_count > 0:
        cell += f' rowSpan="{merge_count}"'
    cell += '>'
    
    cell += '<a:txBody><a:bodyPr/><a:lstStyle/>'
    cell += f'<a:p><a:pPr algn="ctr"/>'
    if text:
        cell += f'<a:r><a:rPr lang="en-US" sz="1200" dirty="0"/><a:t>{text}</a:t></a:r>'
    else:
        cell += '<a:endParaRPr lang="en-US" sz="1200"/>'
    cell += '</a:p></a:txBody>'
    
    # Cell properties
    cell += '<a:tcPr anchor="ctr"'
    cell += '>'
    
    if has_image:
        # Set background image fill
        cell += f'<a:blipFill><a:blip r:embed="{img_rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
    
    cell += '</a:tcPr></a:tc>'
    return cell

def build_merged_cell():
    """Build a vertically merged continuation cell."""
    return '<a:tc vMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>'

# Build full table
table_xml = f'''<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:nvGraphicFramePr>
<p:cNvPr id="10" name="Table 9"/>
<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
<p:nvPr><p:ph idx="1"/></p:nvPr>
</p:nvGraphicFramePr>
<p:xfrm><a:off x="{table_x}" y="{table_y}"/><a:ext cx="{table_width}" cy="{table_height}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandRow="1">
<a:tableStyleId>{{5940675A-B579-460E-94D1-54222C63F5DA}}</a:tableStyleId>
</a:tblPr>
<a:tblGrid>
<a:gridCol w="{col1_width}"/>
<a:gridCol w="{col2_width}"/>
<a:gridCol w="{col3_width}"/>
</a:tblGrid>'''

# Row 1 (header)
table_xml += f'<a:tr h="{row_height}">'
table_xml += build_cell("Men")
table_xml += build_cell("")  # Middle column header (empty in row 1)
table_xml += build_cell("Women")
table_xml += '</a:tr>'

# Row 2 (first merged row - has rowSpan="6")
table_xml += f'<a:tr h="{row_height}">'
table_xml += build_cell(table_data[1][0])
table_xml += build_cell("", merge_count=6, has_image=True, img_rid='rId2')
table_xml += build_cell(table_data[1][2])
table_xml += '</a:tr>'

# Rows 3-7 (merged continuation)
for row_idx in range(2, 7):
    table_xml += f'<a:tr h="{row_height}">'
    table_xml += build_cell(table_data[row_idx][0])
    table_xml += build_merged_cell()
    table_xml += build_cell(table_data[row_idx][2])
    table_xml += '</a:tr>'

table_xml += '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'

# Update slide 6 - replace content placeholder with table
root = read_xml('ppt/slides/slide6.xml')
spTree = root.find(f'.//{P}spTree')

# Remove the content placeholder (ph idx="1")
for sp in list(spTree):
    nvPr = sp.find(f'.//{P}nvPr')
    if nvPr is not None:
        ph = nvPr.find(f'{P}ph')
        if ph is not None and ph.get('idx') == '1' and ph.get('type') is None:
            spTree.remove(sp)
            break

# Add table
table_elem = ET.fromstring(table_xml)
spTree.append(table_elem)

write_xml(root, 'ppt/slides/slide6.xml')
print("  Done")


# ============================================================
# Update [Content_Types].xml
# ============================================================
print("Updating Content_Types...")
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('', CT_NS)
ct_root = read_xml('[Content_Types].xml')

# Add xlsx content type if not present
has_xlsx = False
has_jpg_override = False
for elem in ct_root:
    ext = elem.get('Extension', '')
    if ext == 'xlsx': has_xlsx = True

if not has_xlsx:
    d = ET.SubElement(ct_root, f'{{{CT_NS}}}Default')
    d.set('Extension', 'xlsx')
    d.set('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Add diagram content types
for part, ct in [
    ('/ppt/diagrams/data1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml'),
    ('/ppt/diagrams/colors1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml'),
    ('/ppt/diagrams/style1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml'),
    ('/ppt/diagrams/layout1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayoutDefinition+xml'),
]:
    override = ET.SubElement(ct_root, f'{{{CT_NS}}}Override')
    override.set('PartName', part)
    override.set('ContentType', ct)

write_xml_no_bom(ct_root, '[Content_Types].xml')
print("  Done")

# ============================================================
# FINAL: Repackage PPTX
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
print("\nALL DONE!")
