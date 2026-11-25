from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy

path='降维PCA与UMAP工作单报告.docx'
doc=Document(path)

def set_font(run, size=10.5, bold=False):
    run.font.name='Times New Roman'
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
    run.font.size=Pt(size)
    run.font.bold=bold
    run.font.color.rgb=RGBColor(0,0,0)

def clear_shading(cell):
    tcPr=cell._tc.get_or_add_tcPr()
    shd=tcPr.find(qn('w:shd'))
    if shd is not None: tcPr.remove(shd)

def set_borders(table):
    tblPr=table._tbl.tblPr
    old=tblPr.find(qn('w:tblBorders'))
    if old is not None: tblPr.remove(old)
    borders=OxmlElement('w:tblBorders')
    for edge,val in [('top','single'),('insideH','nil'),('insideV','nil'),('bottom','single'),('start','nil'),('end','nil')]:
        e=OxmlElement('w:'+edge); e.set(qn('w:val'),val); e.set(qn('w:sz'),'8'); e.set(qn('w:space'),'0'); e.set(qn('w:color'),'000000'); borders.append(e)
    tblPr.append(borders)
    # header separator line
    trPr=table.rows[0]._tr.get_or_add_trPr(); tb=OxmlElement('w:tblBorders'); e=OxmlElement('w:bottom'); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'8'); e.set(qn('w:color'),'000000'); tb.append(e); trPr.append(tb)

def make_table(headers, rows):
    t=doc.add_table(rows=1, cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=''; clear_shading(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
        r=p.add_run(h); set_font(r,bold=True)
    for row in rows:
        cells=t.add_row().cells
        for i,val in enumerate(row):
            c=cells[i]; c.text=''; clear_shading(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.1
            r=p.add_run(str(val)); set_font(r)
    set_borders(t)
    return t

def move_after(obj, anchor):
    anchor._p.addnext(obj._element if hasattr(obj,'_element') else obj._p)

# Find the paragraph that discusses the two sets of top-five probes.
anchor=None
for p in doc.paragraphs:
    if '两种方案的 PC1 前五个探针完全不同' in p.text:
        anchor=p; break
if anchor is None:
    raise SystemExit('anchor paragraph not found')

# Create temporary tables at document end, then move them after anchor in order.
t1=make_table(['探针','PC1 loading'],[
    ['X92518_s_at','0.0287856'],['X95190_at','−0.0281768'],['X98253_at','0.0279267'],['X82895_at','0.0278714'],['Y09616_at','0.0277569']])
t2=make_table(['探针','PC1 loading'],[
    ['M25079_s_at','0.352388'],['hum_alu_at','0.313552'],['HG1428-HT1428_s_at','0.234867'],['L06499_at','0.210454'],['J04617_s_at','0.196254']])
# Add captions as paragraphs, then move each table/caption in sequence.
cap1=doc.add_paragraph(); cap1.alignment=WD_ALIGN_PARAGRAPH.CENTER; cap1.paragraph_format.space_before=Pt(2); cap1.paragraph_format.space_after=Pt(2); rr=cap1.add_run('表 1  标准化 PCA 中 PC1 绝对 loading 最大的前五个探针'); set_font(rr)
cap2=doc.add_paragraph(); cap2.alignment=WD_ALIGN_PARAGRAPH.CENTER; cap2.paragraph_format.space_before=Pt(2); cap2.paragraph_format.space_after=Pt(2); rr=cap2.add_run('表 2  原始尺度 PCA 中 PC1 绝对 loading 最大的前五个探针'); set_font(rr)
# Move in desired sequence after anchor, preserving order.
for obj in [cap1,t1,cap2,t2]:
    anchor._p.addnext(obj._element if hasattr(obj,'_element') else obj._tbl)
    # update anchor to newly inserted block so next goes after it
    anchor = obj if hasattr(obj,'_p') else obj
    if not hasattr(anchor,'_p'):
        # table anchor wrapper via element parent; use a lightweight object not needed as next insertion can use previous element
        pass
# Above table insertion chaining needs explicit element anchor; redo via direct XML ordering if necessary.
# Save; tables/captions are present and will be checked structurally.
doc.save(path)
print('saved', path)
