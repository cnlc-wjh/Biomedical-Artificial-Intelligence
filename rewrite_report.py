from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE=Path(__file__).resolve().parent
RES=BASE/'.tmp/dimension-reduction/student-results'
OUT=BASE/'降维PCA与UMAP工作单报告.docx'

# ---------- typography ----------
def set_run_font(run, east='宋体', size=12, bold=None, color='000000'):
    run.font.name='Times New Roman'
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), east)
    run.font.size=Pt(size)
    run.font.color.rgb=RGBColor.from_string(color)
    if bold is not None: run.bold=bold

def format_para(p, first=True, align=None, before=0, after=6, line=1.25):
    pf=p.paragraph_format
    if first: pf.first_line_indent=Cm(0.74)
    pf.space_before=Pt(before); pf.space_after=Pt(after); pf.line_spacing=line
    if align is not None: p.alignment=align

def add_text(doc,text,bold_prefix=None,first=True,align=None,after=6):
    p=doc.add_paragraph(); format_para(p,first,align,after=after)
    if bold_prefix and text.startswith(bold_prefix):
        r=p.add_run(bold_prefix); set_run_font(r,bold=True)
        r=p.add_run(text[len(bold_prefix):]); set_run_font(r)
    else:
        r=p.add_run(text); set_run_font(r)
    return p

def heading(doc,text,level=1):
    p=doc.add_paragraph(); format_para(p,False,after=6,before=10,line=1.15)
    r=p.add_run(text); set_run_font(r,'黑体',15 if level==1 else 12,bold=True)
    p.paragraph_format.keep_with_next=True
    return p

def set_cell_borders(table):
    tblPr=table._tbl.tblPr
    old=tblPr.find(qn('w:tblBorders'))
    if old is not None: tblPr.remove(old)
    borders=OxmlElement('w:tblBorders')
    # three-line table: top, header bottom, bottom; inner vertical/horizontal absent
    for edge, val in [('top','single'),('insideH','nil'),('insideV','nil'),('bottom','single'),('start','nil'),('end','nil')]:
        el=OxmlElement(f'w:{edge}'); el.set(qn('w:val'),val); el.set(qn('w:sz'),'8'); el.set(qn('w:space'),'0'); el.set(qn('w:color'),'000000'); borders.append(el)
    tblPr.append(borders)
    # header bottom border
    trPr=table.rows[0]._tr.get_or_add_trPr();
    tb=OxmlElement('w:tblBorders')
    el=OxmlElement('w:bottom'); el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'8'); el.set(qn('w:color'),'000000'); tb.append(el); trPr.append(tb)

def clear_shading(cell):
    tcPr=cell._tc.get_or_add_tcPr(); shd=tcPr.find(qn('w:shd'))
    if shd is not None: tcPr.remove(shd)

def add_table(doc,headers,rows,widths=None):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=True
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=''; clear_shading(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
        r=p.add_run(str(h)); set_run_font(r,'宋体',10.5,bold=True)
    for row in rows:
        cs=t.add_row().cells
        for i,v in enumerate(row):
            c=cs[i]; c.text=''; clear_shading(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.1
            r=p.add_run(str(v)); set_run_font(r,'宋体',10.5)
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Cm(w)
    set_cell_borders(t); doc.add_paragraph().paragraph_format.space_after=Pt(2)
    return t

def add_fig(doc,path,caption,width=6.1):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(2)
    run=p.add_run(); run.add_picture(str(path),width=Inches(width))
    cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_after=Pt(8)
    r=cp.add_run(caption); set_run_font(r,'宋体',10.5)

# ---------- document ----------
doc=Document(); sec=doc.sections[0]; sec.top_margin=Cm(2.54); sec.bottom_margin=Cm(2.54); sec.left_margin=Cm(3.18); sec.right_margin=Cm(3.18)
# default styles
for name,east,size in [('Normal','宋体',12),('Title','黑体',22)]:
    s=doc.styles[name]; s.font.name='Times New Roman'; s._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),east); s.font.size=Pt(size); s.font.color.rgb=RGBColor(0,0,0)

p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(20)
r=p.add_run('降维 PCA 与 UMAP 白血病表达数据分析报告'); set_run_font(r,'黑体',22,bold=True)

heading(doc,'1 摘要',1)
add_text(doc,'本报告使用 72 个白血病样本的 7,129 个探针表达值，按照工作单要求完成数据准备、标准化 PCA、原始尺度 PCA 对照、UMAP 参数对照和训练集内交叉验证。报告先说明 PCA 与 UMAP 的计算原理，再逐项给出实验设计、运行结果和结果边界。标准化 PCA 的 PC1 和 PC2 分别解释 14.9216% 和 9.4462% 的方差，前两者合计 24.3678%；达到 80% 和 90% 累计解释率分别需要 37 和 51 个主成分。UMAP 基线设置下，二维嵌入与原始标准化空间的 5 近邻平均重合率为 41.11%。在 38 个训练样本上进行三折交叉验证时，标准化后直接使用 1 近邻的平均 balanced accuracy 为 0.842593，加入 10 个 PCA 成分后的平均值为 0.763889。上述结果支持将二维图用于结构探索，但不足以单独证明疾病机制、真实生物学亚型或新患者预测能力。')
heading(doc,'2 研究问题',1)
add_text(doc,'本报告围绕三个问题展开：第一，PCA 和 UMAP 将 7,129 维表达矩阵压缩到二维时各自保留了什么信息；第二，按列标准化与仅中心化会怎样改变 PCA 的主方向；第三，UMAP 的局部邻域在参数和随机种子改变后是否稳定，以及二维分离能否支持生物学或预测性结论。')
add_text(doc,'每一行代表一个样本，每一列代表一个探针；sample_id、original_split 和 diagnosis 是元数据列，未作为表达特征输入。diagnosis 仅用于结果图的事后着色，不参与 PCA 或 UMAP 的拟合。')

heading(doc,'3 实验结果',1)
heading(doc,'3.1 PCA 与 UMAP 的计算原理',2)
add_text(doc,'PCA（主成分分析）先对每个探针的表达列进行中心化；本实验还将每列除以其标准差，使每个探针在相同的尺度上参与计算。设标准化后的矩阵为 Z，PCA 寻找单位方向 w1，使样本投影 t1=Zw1 的方差最大；随后在与 w1 正交的条件下寻找 w2，并依次得到后续主成分。实际计算可通过 Z=UΣVᵀ 的奇异值分解完成，样本坐标由投影得到，某个主成分的 explained variance ratio 等于该方向方差占输入矩阵总方差的比例。')
add_text(doc,'PC1 坐标表示样本在第一条最大方差方向上的投影位置，数值的正负只是相对于所选方向的坐标；将某个主成分方向和所有样本得分同时乘以 -1，只会翻转坐标轴，不会改变方差解释率或样本间距离。PCA 是无监督的，PC1 最大化的是总体表达方差，不等于最大化 ALL/AML 分类差异，也不等于某一个基因的作用。')
add_text(doc,'UMAP（统一流形近似与投影）先在高维空间中为每个样本寻找近邻，并把这些近邻关系表示为带权图；随后在二维空间中迭代调整坐标，使高维近邻尽量仍然靠近，同时允许远处样本分开。n_neighbors 决定近邻图使用的邻域范围，min_dist 控制二维点可以压缩到多紧，random_state 决定随机初始化和迭代中的可复现性。UMAP1 和 UMAP2 只是嵌入坐标，没有 PCA 那样的方差解释率，图上的岛屿距离、方向和面积不能直接当作生物学量。')

heading(doc,'3.2 数据准备与拟合范围',2)
add_text(doc,'实验数据为 72×7,129 的表达矩阵。描述性降维在全部 72 个样本上拟合，用于回答“当前样本结构如何呈现”；诊断标签只在拟合完成后用于着色。面向新样本的预测流程则先划分训练集和测试集，在每个训练折内部拟合 StandardScaler、PCA 和分类器，再将已拟合的变换应用到验证折，避免测试信息进入拟合过程。')
add_table(doc,['分析目的','拟合范围','主要输出','评价方式'],[
['当前结构描述','全部 72 个样本','PCA/UMAP 坐标、解释率、邻域表','图形与参数对照'],
['新样本预测验证','训练集 38 个样本内三折','Scaler→PCA（可选）→1 近邻','balanced accuracy；测试集不参与选择']],[3.0,3.2,4.2,4.2])

heading(doc,'3.3 PCA：标准化与原始尺度对照',2)
add_text(doc,'实验设计：从表达矩阵中排除三列元数据；标准化方案使用 StandardScaler 按列计算均值和总体标准差（ddof=0），得到每列均值接近 0、总体标准差接近 1 的矩阵。随后使用 full SVD 拟合 PCA，记录 PC1、PC2 的 explained variance ratio、累计解释率、达到 80%/90% 的成分数和 PC1 绝对权重最大的五个探针。对照方案仅中心化而不除以标准差，用来观察量纲和原始方差大小对主方向的影响。')
add_table(doc,['输入方案','PC1','PC2','前两项合计','达到 80%','达到 90%'],[
['按列标准化','14.9216%','9.4462%','24.3678%','37','51'],['原始尺度，仅中心化','17.4278%','10.3386%','27.7664%','20','34']],[3.5,2.0,2.0,2.6,2.2,2.2])
add_text(doc,'结果：标准化 PCA 的前两条轴只能显示输入总方差的 24.3678%，其余方差位于未绘制的 PC3 及后续方向；这些方向不需要与 PC1、PC2 “重合”，而是与前两条主成分正交。原始尺度方案的解释率更高并不表示更好，因为它让原始波动较大的探针获得更大的影响。两种方案的 PC1 前五个探针完全不同：标准化方案为 X92518_s_at、X95190_at、X98253_at、X82895_at、Y09616_at；原始尺度方案为 M25079_s_at、hum_alu_at、HG1428-HT1428_s_at、L06499_at、J04617_s_at。该变化说明尺度处理改变了 PCA 优化的方差结构。')
add_fig(doc,RES/'pca-comparison.png','图 1  标准化 PCA 与原始尺度 PCA 的 PC1-PC2 坐标对照',6.2)
add_fig(doc,RES/'pca-variance.png','图 2  两种 PCA 输入方案的累计方差解释率',5.9)
add_text(doc,'PC1 的 loading 是各探针在第一主成分方向上的权重；绝对值大表示该探针对该方向贡献较大，正负号表示与该方向同向或反向。它不等于差异表达显著性、疾病因果作用或分类重要性。整体翻转 PC1 的符号不会改变任何解释率和样本间关系。')

heading(doc,'3.4 UMAP：局部邻域与参数对照',2)
add_text(doc,'实验设计：沿用任务 3.3 的标准化矩阵，不按诊断标签筛选探针。基线设置为 n_neighbors=15、min_dist=0.1、metric=euclidean、random_state=42、n_jobs=1；然后分别只改变 n_neighbors 为 5、只改变 min_dist 为 0.5、只改变随机种子为 7。每次保存 72 个样本的坐标，并用同一组样本编号追踪邻域变化。')
add_table(doc,['设置','n_neighbors','min_dist','seed','与原始空间 5 近邻重合','与基线二维近邻重合'],[
['baseline','15','0.1','42','41.11%','100.00%'],['neighbors5','5','0.1','42','45.56%','55.28%'],['mindist05','15','0.5','42','40.00%','64.17%'],['seed7','15','0.1','7','40.56%','60.28%']],[2.4,2.0,1.6,1.2,3.1,3.1])
add_text(doc,'结果：四种设置下原始高维空间的 5 近邻保留率约为 40.00%—45.56%，说明二维图只保留了部分局部关系。改变 n_neighbors、min_dist 或随机种子都会改变若干样本的邻域和整体布局；因此应比较同一样本的邻居，而不能把不同图上的 UMAP 坐标数值直接对齐。UMAP 图中出现一团点只能说明这些样本在当前嵌入中相对接近，可能来自共享表达模式，也可能受到批次、测量和预处理因素影响。')
add_fig(doc,RES/'umap-settings.png','图 3  UMAP 基线及三个单因素设置的样本布局',6.2)
add_text(doc,'PCA 和 UMAP 的信息侧重点不同：PCA 保留总体线性方差方向，坐标、loading 和解释率可追踪；UMAP 优先展示局部邻域，坐标轴没有固定生物学含义，也没有 explained variance。UMAP 岛屿数量不能直接当作亚型数量，岛屿间距离和面积也不能直接解释为疾病距离、阶段或群体比例。')

heading(doc,'3.5 预测用途的交叉验证',2)
add_text(doc,'实验设计：只使用原始训练集的 38 个样本，采用 StratifiedKFold(n_splits=3, shuffle=True, random_state=42)。在每个训练折中独立拟合完整 Pipeline，比较“StandardScaler→1 近邻”和“StandardScaler→PCA(10)→1 近邻”；验证折只接受训练折已经拟合好的变换。原始 34 个测试样本不参与本次模型选择。')
add_table(doc,['验证折','标准化→1 近邻','标准化→PCA(10)→1 近邻'],[['1','0.819444','0.694444'],['2','0.875000','0.763889'],['3','0.833333','0.833333'],['三折平均','0.842593','0.763889']],[3.2,4.0,5.0])
add_text(doc,'结果：在这组样本、固定三折划分和 1 近邻设置下，加入 PCA(10) 没有提高平均 balanced accuracy。由于训练样本数量较少，这个结果只能说明本次比较中没有观察到 PCA(10) 的预测收益，不能推出 PCA 在所有数据集或分类器中都无效。交叉验证也不能把描述性降维图自动变成临床预测证据。')

heading(doc,'3.6 结果边界与可支持的结论',2)
add_text(doc,'本实验可以支持的结论是：标准化会改变 PCA 关注的方差方向；标准化 PCA 前两轴只覆盖 24.3678% 的输入方差；UMAP 的局部邻域对参数和随机种子有敏感性；在当前训练集内，PCA(10) 未改善所选 1 近邻模型的三折平均 balanced accuracy。')
add_text(doc,'本实验不能单独支持的结论是：二维图分开就等于发现疾病机制；每个点群就是一个真实生物学亚型；图中岛屿数等于亚型数；当前 72 个样本上的分离可以直接预测新患者；数据已经排除了批次效应。若要把候选结构升级为生物医学结论，还需要独立队列复现、批次信息和外部临床或分子证据。')

heading(doc,'4 结论',1)
add_text(doc,'本报告完成了工作单要求的 PCA 与 UMAP 降维实验，并将描述性降维与预测验证分开。PCA 提供了可计算的方差解释率和 loading，适合检查整体线性结构；UMAP 提供了局部邻域的探索性可视化，但布局会随参数和随机种子变化。因而，当前结果适合提出后续验证假设，不足以独立证明疾病机制、真实亚型数量或新样本预测能力。')
heading(doc,'5 数据与复现信息',1)
add_text(doc,'输入数据为项目目录中的 Golub 白血病表达矩阵，包含 72 个样本和 7,129 个探针。分析环境为 conda highdim，Python 3.12.15；numpy 2.3.5、pandas 2.3.3、scipy 1.17.1、scikit-learn 1.8.0、umap-learn 0.5.12、matplotlib 3.10.8。坐标表、loading、方差解释率、邻域对照表和图像保存在 .tmp/dimension-reduction/student-results/；完整运行脚本为 .tmp/dimension-reduction/student_analysis.py。')

# enforce all existing runs black and correct English font; table text already set
for p in doc.paragraphs:
    for r in p.runs:
        r.font.color.rgb=RGBColor(0,0,0)
        if r._element.get_or_add_rPr().rFonts is None: r._element.get_or_add_rPr().get_or_add_rFonts()
        r.font.name='Times New Roman'
# prevent widows/orphans
for p in doc.paragraphs:
    p.paragraph_format.widow_control=True

doc.save(OUT)
print(OUT)
