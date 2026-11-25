from pathlib import Path
from docx import Document
import json
from zipfile import ZipFile
from lxml import etree

base=Path(__file__).resolve().parent.parent
s=json.loads((base/'results/summary.json').read_text(encoding='utf-8'))
assert s['primary']['n1']+s['primary']['n0']==865
assert sum(x['original_n'] for x in s['descriptives'])==920
assert sum(x['missing_n'] for x in s['descriptives'])==55
assert s['primary']['ci_low']<s['primary']['delta']<s['primary']['ci_high']<0
assert abs(s['adjusted_models'][0]['estimate']-s['primary']['delta'])<1e-10
d=Document(base/'统计检验与假设设计报告.docx')
text='\n'.join(p.text for p in d.paragraphs)+'\n'+'\n'.join(c.text for t in d.tables for r in t.rows for c in r.cells)
for token in ['1.3016','−20.5389','第一类错误','另','四句结论']:
    # “另”可以出现在课后设计说明；用于检查正文未意外截断。
    assert token in text,token
assert len(d.inline_shapes)==2
assert len(d.tables)==8
sec=d.sections[0]
usable=sec.page_width-sec.left_margin-sec.right_margin
for t in d.tables:
    assert sum(c.width for c in t.columns) <= usable+1000
    assert t.rows[0]._tr.trPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblHeader') is not None
with ZipFile(base/'统计检验与假设设计报告.docx') as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            etree.fromstring(z.read(name))
for filename in ['hypothesis_analysis.py','build_hypothesis_report.py']:
    compile((base/filename).read_text(encoding='utf-8'),filename,'exec')
print('PASS: numerical consistency, report sections, all XML, images, table widths, repeated headers, Python syntax')
print('DOCX:',(base/'统计检验与假设设计报告.docx').stat().st_size,'bytes')
