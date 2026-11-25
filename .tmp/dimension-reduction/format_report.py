from pathlib import Path
import shutil
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

base = Path(__file__).resolve().parents[2]
path = base / '降维PCA与UMAP工作单报告.docx'
backup = path.parent / '.tmp/dimension-reduction/report-before-formatting.docx'
backup.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(path, backup)
doc = Document(path)

def set_fonts(run, east_asia, size_pt, bold=None):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:ascii'), 'Times New Roman')
    rfonts.set(qn('w:hAnsi'), 'Times New Roman')
    rfonts.set(qn('w:eastAsia'), east_asia)
    rfonts.set(qn('w:cs'), 'Times New Roman')

def set_style_font(style, east_asia, size_pt, bold=None):
    style.font.name = 'Times New Roman'
    style.font.size = Pt(size_pt)
    if bold is not None:
        style.font.bold = bold
    rpr = style._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:ascii'), 'Times New Roman')
    rfonts.set(qn('w:hAnsi'), 'Times New Roman')
    rfonts.set(qn('w:eastAsia'), east_asia)
    rfonts.set(qn('w:cs'), 'Times New Roman')

# Reclassify numbered headings by their explicit section-number structure.
for p in doc.paragraphs:
    if re.match(r'^\d+\.\d+\s', p.text.strip()):
        p.style = doc.styles['Heading 2']
    elif re.match(r'^\d+\s', p.text.strip()):
        p.style = doc.styles['Heading 1']

normal = doc.styles['Normal']
set_style_font(normal, '宋体', 12)
normal.paragraph_format.first_line_indent = Pt(24)
normal.paragraph_format.line_spacing = 1.25

set_style_font(doc.styles['Heading 1'], '黑体', 15, True)  # 小三
doc.styles['Heading 1'].paragraph_format.first_line_indent = Pt(0)
doc.styles['Heading 1'].paragraph_format.space_before = Pt(12)
doc.styles['Heading 1'].paragraph_format.space_after = Pt(6)
doc.styles['Heading 1'].paragraph_format.keep_with_next = True

set_style_font(doc.styles['Heading 2'], '黑体', 14, True)  # 四号
doc.styles['Heading 2'].paragraph_format.first_line_indent = Pt(0)
doc.styles['Heading 2'].paragraph_format.space_before = Pt(8)
doc.styles['Heading 2'].paragraph_format.space_after = Pt(4)
doc.styles['Heading 2'].paragraph_format.keep_with_next = True

set_style_font(doc.styles['Title'], '黑体', 18, True)

for p in doc.paragraphs:
    name = p.style.name
    if name == 'Title':
        east, size = '黑体', 18
        p.paragraph_format.first_line_indent = Pt(0)
    elif name == 'Heading 1':
        east, size = '黑体', 15
    elif name == 'Heading 2':
        east, size = '黑体', 14
    else:
        east, size = '宋体', 12
        p.paragraph_format.line_spacing = 1.25
        # Figure captions remain centered; prose paragraphs receive a two-character indent.
        if p.text.strip().startswith(('图 1', '图 2')):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
        else:
            p.paragraph_format.first_line_indent = Pt(24)
    for run in p.runs:
        set_fonts(run, east, size)

def set_border(parent, edge, val, sz=8, color='000000'):
    tag = qn('w:' + edge)
    el = parent.find(tag)
    if el is None:
        el = OxmlElement('w:' + edge)
        parent.append(el)
    el.set(qn('w:val'), val)
    if val != 'nil':
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)

for table in doc.tables:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn('w:tblBorders'))
    if borders is None:
        borders = OxmlElement('w:tblBorders')
        tbl_pr.append(borders)
    for edge in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
        set_border(borders, edge, 'nil')
    for ridx, row in enumerate(table.rows):
        if ridx == 0:
            tr_pr = row._tr.get_or_add_trPr()
            repeat = OxmlElement('w:tblHeader')
            repeat.set(qn('w:val'), 'true')
            tr_pr.append(repeat)
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tc_pr = cell._tc.get_or_add_tcPr()
            shd = tc_pr.find(qn('w:shd'))
            if shd is not None:
                tc_pr.remove(shd)
            cborders = tc_pr.find(qn('w:tcBorders'))
            if cborders is None:
                cborders = OxmlElement('w:tcBorders')
                tc_pr.append(cborders)
            for edge in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
                set_border(cborders, edge, 'nil')
            if ridx == 0:
                set_border(cborders, 'top', 'single', 12)
            if ridx == 0:
                set_border(cborders, 'bottom', 'single', 8)
            if ridx == len(table.rows) - 1:
                set_border(cborders, 'bottom', 'single', 12)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                for run in p.runs:
                    set_fonts(run, '宋体', 12, True if ridx == 0 else None)

# Ensure any existing headers and footers also follow the Latin/Chinese typeface rule.
for section in doc.sections:
    for container in (section.header, section.footer):
        for p in container.paragraphs:
            for run in p.runs:
                set_fonts(run, '宋体', 12)

doc.save(path)
print(f'Formatted: {path}')
print(f'Backup: {backup}')
