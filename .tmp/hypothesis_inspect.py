from pathlib import Path
from docx import Document
import pandas as pd
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re

root = Path.cwd()
pptx_path = root / "V3/slides/hypothesis-test-pbl.pptx"
ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
with ZipFile(pptx_path) as z:
    slide_names = sorted([n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)], key=lambda x:int(re.search(r"slide(\d+)", x).group(1)))
    print("SLIDES", len(slide_names))
    for i, name in enumerate(slide_names, 1):
        print(f"--- SLIDE {i} ---")
        rootxml = ET.fromstring(z.read(name))
        texts = ["".join(t.text or "" for t in tx.findall(".//a:t", ns)) for tx in rootxml.findall(".//a:p", ns)]
        for t in texts:
            if t.strip(): print(repr(t))

df = pd.read_csv(root / "V3/data/heart-disease-hypothesis.csv")
print("SHAPE", df.shape)
print("COLUMNS", df.columns.tolist())
print("DTYPES", df.dtypes.to_dict())
print("HEAD")
print(df.head(5).to_string(index=False))
print("COUNTS", df["disease"].value_counts(dropna=False).to_dict())
print("MISSING", df.isna().sum().to_dict())
print("GROUP")
print(df.groupby("disease")["thalach"].agg(["count", "mean", "std", "median"]).to_string())
print("SITES", df["site"].value_counts().to_dict())

ref = Path(r"D:\学习\博一\ai课课件\4. Dimension-reduction\降维PCA与UMAP报告.docx")
doc = Document(ref)
sec = doc.sections[0]
print("SECTION_CM", *[round(getattr(sec, k).cm, 2) for k in ["page_width", "page_height", "top_margin", "bottom_margin", "left_margin", "right_margin"]])
print("STYLES")
for name in ["Normal", "Title", "Heading 2", "Heading 3"]:
    st = doc.styles[name]
    print(name, "font", st.font.name, "size", st.font.size.pt if st.font.size else None, "bold", st.font.bold)
print("TABLES", len(doc.tables))
for i, table in enumerate(doc.tables):
    print("TABLE", i, "rows", len(table.rows), "cols", len(table.columns), "style", table.style.name)
    print("HEADER", [cell.text for cell in table.rows[0].cells])
print("HEADINGS")
for para in doc.paragraphs:
    if para.style.name.startswith("Heading"):
        print(para.text, para.style.name)
