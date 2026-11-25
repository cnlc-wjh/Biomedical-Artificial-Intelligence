import importlib.util
from docx import Document
from pathlib import Path
import json
for m in ['win32com','fitz','reportlab','pdf2image','PIL']:
    print(m, bool(importlib.util.find_spec(m)))
d=Document(Path.cwd().parent/'4. Dimension-reduction/降维PCA与UMAP报告.docx')
for i,p in enumerate(d.paragraphs[:16]):
    print('PARA',i,p.text[:90], 'style',p.style.name,'format',p.paragraph_format.first_line_indent,p.paragraph_format.line_spacing,p.paragraph_format.space_after,'runs',[(r.text[:20], r.font.name,r.font.size.pt if r.font.size else None,r.bold, r._r.rPr.xml if r._r.rPr is not None else '') for r in p.runs[:1]])
