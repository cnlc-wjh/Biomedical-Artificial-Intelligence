from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE = Path(__file__).resolve().parent
RES = BASE / '.tmp/dimension-reduction/student-results'
OUT = BASE / '降维PCA与UMAP工作单报告.docx'

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr(); shd = tcPr.find(qn('w:shd'))
    if shd is None: shd = OxmlElement('w:shd'); tcPr.append(shd)
    shd.set(qn('w:fill'), fill)

def set_cell_margins(cell, top=100, start=110, bottom=100, end=110):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr(); tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None: tcMar = OxmlElement('w:tcMar'); tcPr.append(tcMar)
    for m, v in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tcMar.find(qn(f'w:{m}'))
        if node is None: node = OxmlElement(f'w:{m}'); tcMar.append(node)
        node.set(qn('w:w'), str(v)); node.set(qn('w:type'), 'dxa')

def add_table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers)); t.style = 'Normal Table'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=str(h); set_cell_shading(c,'D9EAF2'); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; set_cell_margins(c)
        for r in c.paragraphs[0].runs: r.bold=True; r.font.size=Pt(9)
    for ridx,row in enumerate(rows):
        cells=t.add_row().cells
        for i,v in enumerate(row):
            cells[i].text=str(v); cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; set_cell_margins(cells[i])
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after=Pt(0)
                for r in p.runs: r.font.size=Pt(9)
            if ridx % 2: set_cell_shading(cells[i],'F5F8FA')
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Cm(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)
    return t

def add_body(doc, text, bold_prefix=None):
    p=doc.add_paragraph(style='Normal'); p.paragraph_format.first_line_indent=Cm(0.74); p.paragraph_format.line_spacing=1.25
    if bold_prefix and text.startswith(bold_prefix):
        p.add_run(bold_prefix).bold=True; p.add_run(text[len(bold_prefix):])
    else: p.add_run(text)
    return p

def add_heading(doc, text, level=2):
    p=doc.add_paragraph(style=f'Heading {level}'); p.paragraph_format.keep_with_next=True; p.add_run(text); return p

doc=Document()
sec=doc.sections[0]; sec.top_margin=Cm(2.54); sec.bottom_margin=Cm(2.54); sec.left_margin=Cm(3.18); sec.right_margin=Cm(3.18)
styles=doc.styles
styles['Normal'].font.name='宋体'; styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体'); styles['Normal'].font.size=Pt(10.5)
for name,size in [('Title',18),('Heading 2',14),('Heading 3',11.5)]:
    styles[name].font.name='黑体'; styles[name]._element.rPr.rFonts.set(qn('w:eastAsia'),'黑体'); styles[name].font.size=Pt(size); styles[name].font.color.rgb=None

p=doc.add_paragraph(style='Title'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run('降维、PCA 与 UMAP：白血病表达数据分析报告')
add_heading(doc,'1 摘要',2)
add_body(doc,'本报告使用 72 份白血病微阵列表达样本及 7,129 个探针，考察标准化、PCA 与 UMAP 在描述高维样本结构和支持后续预测时分别保留了什么信息。数据包含 47 份 ALL 和 25 份 AML；原始划分为训练集 38 份、测试集 34 份。标准化 PCA 的 PC1 和 PC2 分别解释 14.9216% 和 9.4462% 的方差，前两维合计 24.3678%，累计达到 80% 和 90% 分别需要 37 和 51 个成分。基线 UMAP 与标准化原空间的二维 5 近邻平均重合为 41.11%；改变邻居数、min_dist 或随机种子后，局部邻域会变化。训练集内三折比较中，标准化后直接使用 1 近邻的平均验证 balanced accuracy 为 0.8426，加入 PCA(10) 后为 0.7639。结果支持把二维图作为探索线索，但不能凭图证明疾病机制、真实亚型数量或新患者预测能力。')
add_heading(doc,'2 研究问题',2)
add_body(doc,'本研究围绕三个问题展开：第一，PCA 和 UMAP 把 7,129 维表达矩阵压缩到二维时，各自保留什么、隐藏什么？第二，标准化会如何改变 PCA 的主要方向？第三，二维图中的样本分离和近邻关系，在参数变化及预测验证中有多稳定？')
add_body(doc,'每行代表一份白血病样本，标签为 ALL 或 AML；每个表达列对应一个探针。探针不等同于不同基因，也不保证彼此独立。教学编号 G001–G072 只能用于追踪样本，不能证明存在 72 位互不重复的患者；本数据还缺少可用于排查批次效应的字段。')
add_heading(doc,'3 实验设计',2)
add_heading(doc,'3.1 数据准备与标准化',3)
add_body(doc,'排除 sample_id、original_split 和 diagnosis 三个元数据列，得到 72 × 7,129 的表达矩阵。描述性降维中，对每个探针按列拟合 StandardScaler；标准化后的列均值绝对值最大为 5.98×10⁻¹⁶，列总体标准差与 1 的差的绝对值最大为 3.33×10⁻¹⁶。标准化表示“偏离该探针自身均值多少个标准差”，与芯片数据上游归一化不是同一步。')
add_heading(doc,'3.2 PCA 对照',3)
add_body(doc,'分别对标准化矩阵和只中心化的原尺度矩阵拟合 full SVD PCA。记录 PC1、PC2 的解释率、累计解释率、PC1 绝对权重最大的五个探针，并保存样本坐标。原尺度对照用于观察尺度选择如何改变主方向；PCA 优化的是整体方差，不使用诊断标签。')
add_heading(doc,'3.3 UMAP 对照',3)
add_body(doc,"使用同一个标准化矩阵拟合二维 UMAP，欧氏距离、n_jobs=1。基线为 n_neighbors=15、min_dist=0.1、random_state=42；三个单因素对照分别将 n_neighbors 改为 5、min_dist 改为 0.5、random_state 改为 7。拟合后统一比较每个样本的二维 5 近邻。")
add_heading(doc,'3.4 预测验证',3)
add_body(doc,'仅在原始 38 个训练样本内使用 StratifiedKFold(n_splits=3, shuffle=True, random_state=42)，比较 StandardScaler → 1 近邻与 StandardScaler → PCA(10) → 1 近邻。每一折在训练部分独立拟合完整 Pipeline，再转换验证部分；原始 34 个测试样本没有参与本次分数选择。')
add_heading(doc,'4 结果',2)
add_heading(doc,'4.1 PCA 的方差保留',3)
add_table(doc,['输入方案','PC1','PC2','前两项合计','达到 80%','达到 90%'],[['按列标准化','14.9216%','9.4462%','24.3678%','37','51'],['原尺度，仅中心化','17.4278%','10.3386%','27.7664%','20','34']],[3.4,2,2,3,2.3,2.3])
add_body(doc,'标准化 PCA 的二维图只显示约 24.37% 的输入总方差，其余约 75.63% 位于没有画出的主成分方向。被省略的方向不会要求样本在二维图中重合；样本可以在 PC1、PC2 上分开，同时在 PC3 及后续方向上还有额外差异。原尺度 PCA 的解释率不能直接与标准化方案判定优劣，因为两者参照的是不同的预处理空间。')
pic=doc.add_paragraph(); pic.alignment=WD_ALIGN_PARAGRAPH.CENTER; pic.add_run().add_picture(str(RES/'pca-umap.png'),width=Inches(6.25))
cap=doc.add_paragraph('图 1  标准化 PCA、原尺度 PCA 及四组 UMAP 对照'); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
add_body(doc,'PC1 是全部探针的线性加权方向，使样本投影方差最大。标准化 PCA 的 PC1 前五个绝对权重探针为 X92518_s_at、X95190_at、X98253_at、X82895_at 和 Y09616_at；原尺度 PCA 的前五个为 M25079_s_at、hum_alu_at、HG1428-HT1428_s_at、L06499_at 和 J04617_s_at，两组没有重合。原尺度高权重探针的原始标准差约为 6,847–15,783，而标准化方案所列探针约为 327–531，说明尺度处理明显改变了 PCA 关注的方向。权重不等于致病作用，PC1 也不是专门最大化 ALL/AML 区分的方向。')
pic=doc.add_paragraph(); pic.alignment=WD_ALIGN_PARAGRAPH.CENTER; pic.add_run().add_picture(str(RES/'pca-variance.png'),width=Inches(5.8))
cap=doc.add_paragraph('图 2  两种 PCA 输入方案的累计解释率'); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
add_heading(doc,'4.2 UMAP 的局部关系与稳定性',3)
add_table(doc,['设置','n_neighbors','min_dist','seed','原空间 5 近邻保留','与基线二维近邻重合'],[['baseline','15','0.1','42','41.11%','100.00%'],['neighbors5','5','0.1','42','45.56%','55.28%'],['mindist05','15','0.5','42','40.00%','64.17%'],['seed7','15','0.1','7','40.56%','60.28%']],[2.3,2,1.6,1.2,3.0,3.1])
add_body(doc,'基线 UMAP 的 41.11% 是标准化原空间与二维图的 5 近邻名单平均重合比例，不是坐标或距离保留比例，也不是分类准确率。改变参数或随机种子会改变二维布局，因此读 UMAP 时应追踪同一样本的邻居，而不是比较不同图的坐标数字。四组设置中，G004 的共同二维近邻为 G001、G007、G008、G027、G049；G010 的四组邻居交集为空，说明不同样本的稳定性不同。')
add_body(doc,'PCA 主要保留线性方向上的方差，可计算解释率并查看 loading；UMAP 主要近似保留局部邻域，坐标轴没有固定生物学含义，远处岛屿之间的距离、方向和面积也不能直接解释为疾病距离、阶段或群体比例。')
add_heading(doc,'4.3 训练集内交叉验证',3)
add_table(doc,['验证折','标准化 → 1 近邻','标准化 → PCA(10) → 1 近邻'],[['1','0.819444','0.694444'],['2','0.875000','0.763889'],['3','0.833333','0.833333'],['三折平均','0.842593','0.763889']],[3.0,4.1,5.0])
add_body(doc,'balanced accuracy 是 ALL 与 AML 两类召回率的平均值，不是总正确率。每折训练分数均为 1.0 对 1 近邻的训练记忆并不意外，不能作为泛化证据。在这次 38 个训练样本、固定三折和 1 近邻设定下，PCA(10) 没有提高平均验证指标；样本很少，不能推广为 PCA 永远无用或必然有害。')
add_heading(doc,'5 解释与限制',2)
add_body(doc,'二维图上的样本靠近或分开，只能说明它们在这次表示中呈现出相应的数值关系。靠近的样本可能共享生物学表达特征，也可能受到同一批次、仪器、试剂或其他技术因素的影响。本数据没有批次字段，因此无法排除批次混杂。即使某些局部邻域在四组 UMAP 中稳定，也只能支持候选结构；不能凭“看到一团点”认定真实亚型数量。')
add_body(doc,'同一诊断类别内部的候选亚型，表示可能存在更具体、稳定的共同生物学特征，不等同于已经证明患病原因不同。要提出生物学亚型，需要在独立样本中复现，并结合额外的分子、临床或功能证据。')
add_body(doc,'本报告的预测比较只使用训练集内交叉验证，未报告原始测试集最终分数；描述性 PCA/UMAP 使用全部 72 个样本拟合，因此不能把图上的分离直接转化为新患者预测能力。标准化、PCA 和 UMAP 在预测任务中都必须在训练折内拟合，以避免验证信息泄露。')
add_heading(doc,'6 结论',2)
add_body(doc,'在本次白血病表达数据中，标准化 PCA 的前两维只解释 24.3678% 的方差，UMAP 的局部邻居关系对参数和随机种子有一定敏感性；固定训练集内的比较也没有显示 PCA(10) 提高 1 近邻验证表现。证据支持的结论是：PCA 和 UMAP 可以帮助探索高维样本结构，但二维分离只是候选线索，不能单独证明疾病机制、真实亚型数量或新患者预测性能。下一步应补充患者独立性、中心和批次信息，在独立队列中预先固定完整 Pipeline 并验证。')
add_heading(doc,'7 数据与复现',2)
add_body(doc,'数据为课程目录中的 Golub 白血病微阵列表达数据 OpenML 数据集 1104（version 1）副本。输入 CSV 的 SHA-256 为 915bce00295ac470f6bd970a4032730a170a080642a5da53ab0bcb0545156f19。运行环境为 conda highdim，Python 3.12.15；numpy 2.3.5、pandas 2.3.3、scipy 1.17.1、scikit-learn 1.8.0、umap-learn 0.5.12、matplotlib 3.10.8。')
add_body(doc,r'完整分析脚本为 .tmp/dimension-reduction/student_analysis.py；运行摘要、交叉验证表、PCA 解释率、loading、样本坐标、稳定邻居表和图像均保存在 .tmp/dimension-reduction/student-results/。PowerShell 运行命令为：& D:\anaconda\envs\highdim\python.exe .tmp\dimension-reduction\student_analysis.py。')
add_heading(doc,'8 参考文献',2)
add_body(doc,'Golub TR, et al. Molecular Classification of Cancer: Class Discovery and Class Prediction by Gene Expression Monitoring. Science. 1999;286(5439):531–537. doi:10.1126/science.286.5439.531.')
add_body(doc,'McInnes L, Healy J, Melville J. UMAP: Uniform Manifold Approximation and Projection for Dimension Reduction. arXiv:1802.03426.')
doc.save(OUT)
print(OUT)
