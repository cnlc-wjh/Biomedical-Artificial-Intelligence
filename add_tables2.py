from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
path='降维PCA与UMAP工作单报告.docx'; doc=Document(path)
def font(r,b=False):
 r.font.name='Times New Roman'; r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体'); r.font.size=Pt(10.5); r.bold=b; r.font.color.rgb=RGBColor(0,0,0)
def shade(c):
 s=c._tc.get_or_add_tcPr().find(qn('w:shd'))
 if s is not None:c._tc.get_or_add_tcPr().remove(s)
def table(headers,rows):
 t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for i,h in enumerate(headers):
  c=t.rows[0].cells[i]; c.text=''; shade(c); p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0); font(p.add_run(h),True)
 for row in rows:
  cs=t.add_row().cells
  for i,v in enumerate(row):
   cs[i].text=''; shade(cs[i]); p=cs[i].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0); font(p.add_run(str(v)))
 pr=t._tbl.tblPr; b=OxmlElement('w:tblBorders')
 for edge,val in [('top','single'),('insideH','nil'),('insideV','nil'),('bottom','single'),('start','nil'),('end','nil')]:
  e=OxmlElement('w:'+edge); e.set(qn('w:val'),val); e.set(qn('w:sz'),'8'); e.set(qn('w:color'),'000000'); b.append(e)
 pr.append(b); return t
def cap(text):
 p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2); font(p.add_run(text)); return p
anchor=next(p for p in doc.paragraphs if '两种方案的 PC1 前五个探针完全不同' in p.text)
# create nodes at end
a=cap('表 1  标准化 PCA 中 PC1 绝对 loading 最大的前五个探针'); t1=table(['探针','PC1 loading'],[['X92518_s_at','0.0287856'],['X95190_at','−0.0281768'],['X98253_at','0.0279267'],['X82895_at','0.0278714'],['Y09616_at','0.0277569']]); b=cap('表 2  原始尺度 PCA 中 PC1 绝对 loading 最大的前五个探针'); t2=table(['探针','PC1 loading'],[['M25079_s_at','0.352388'],['hum_alu_at','0.313552'],['HG1428-HT1428_s_at','0.234867'],['L06499_at','0.210454'],['J04617_s_at','0.196254']])
# detach and insert sequence after anchor
for obj in [a,t1,b,t2]:
 el=obj._p if hasattr(obj,'_p') else obj._tbl
 el.getparent().remove(el)
# insert in reverse after anchor to preserve order
prev=anchor._p
for obj in [a,t1,b,t2]:
 el=obj._p if hasattr(obj,'_p') else obj._tbl
 prev.addnext(el); prev=el
# add spacing after table via existing paragraph not necessary
doc.save(path); print('saved')
