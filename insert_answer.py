from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from copy import deepcopy
p='降维PCA与UMAP工作单报告.docx'
d=Document(p)
target='PC1 的 loading 是各探针在第一主成分方向上的权重；绝对值大表示该探针对该方向贡献较大，正负号表示与该方向同向或反向。它不等于差异表达显著性、疾病因果作用或分类重要性。整体翻转 PC1 的符号不会改变任何解释率和样本间关系。'
newtext='中心化后的 72 个样本至多有 71 个非零主成分。原因是每个探针列在中心化后都满足 72 个样本值之和为 0，这构成一个线性约束，使中心化矩阵的秩最多为 72−1=71。因此，其协方差矩阵最多有 71 个非零特征值，PCA 最多得到 71 个非零主成分；这不是因为探针数量不足，而是因为中心化消除了样本均值方向。'
idx=None
for i,para in enumerate(d.paragraphs):
    if para.text==target:
        idx=i; break
if idx is None:
    raise SystemExit('target not found')
para=d.paragraphs[idx]
# build paragraph after target, copying paragraph properties exactly
# Above branch can be unsafe; replace with explicit new paragraph element
# safer: create via add_paragraph then move after target
newp=d.add_paragraph()
newp._p.getparent().remove(newp._p)
para._p.addnext(newp._p)
# copy paragraph properties
if para._p.pPr is not None:
    newp._p.insert(0, deepcopy(para._p.pPr))
# add run and copy first run formatting
r=newp.add_run(newtext)
if para.runs:
    if para.runs[0]._r.rPr is not None:
        r._r.insert(0, deepcopy(para.runs[0]._r.rPr))
# ensure requested body font/black while preserving paragraph structure
r.font.name='Times New Roman'; r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体'); r.font.size=Pt(12); r.font.color.rgb=RGBColor(0,0,0)
d.save(p)
print('inserted', idx, 'paragraphs', len(d.paragraphs))
