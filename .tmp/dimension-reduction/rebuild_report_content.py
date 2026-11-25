from pathlib import Path
import shutil
import json
import re
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parents[2]
OUTDIR = BASE / '.tmp/dimension-reduction/student-results'
SOURCE_DOCX = BASE / '降维PCA与UMAP工作单报告.docx'
DOCX = BASE / '降维PCA与UMAP工作单报告_修订版.docx'
BACKUP = BASE / '.tmp/dimension-reduction/report-before-content-revision.docx'
shutil.copy2(SOURCE_DOCX, BACKUP)
summary = json.loads((OUTDIR / 'summary.json').read_text(encoding='utf-8'))

plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.labelcolor': 'black', 'text.color': 'black', 'axes.edgecolor': 'black', 'xtick.color': 'black', 'ytick.color': 'black'})
COLORS = {'ALL': '#387ca3', 'AML': '#d57839'}

def make_pca_figure():
    fig, axes = plt.subplots(1, 2, figsize=(9.1, 4.0))
    for ax, kind in zip(axes, ['standardized', 'raw']):
        data = pd.read_csv(OUTDIR / f'pca-{kind}-coordinates.csv')
        for label in ['ALL', 'AML']:
            part = data[data.diagnosis == label]
            ax.scatter(part.PC1, part.PC2, s=22, color=COLORS[label], label=label, alpha=.85)
        info = summary['pca'][kind]
        ax.set_xlabel(f"PC1 ({info['pc1']:.1%})")
        ax.set_ylabel(f"PC2 ({info['pc2']:.1%})")
        ax.set_title('PCA ' + kind)
        if kind == 'raw': ax.ticklabel_format(axis='both', style='sci', scilimits=(0, 0))
        ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(OUTDIR / 'pca-comparison.png', dpi=200, bbox_inches='tight')
    plt.close(fig)

def make_umap_figure():
    fig, axes = plt.subplots(2, 2, figsize=(8.2, 6.6))
    keys = ['baseline', 'neighbors5', 'mindist05', 'seed7']
    titles = ['Baseline', 'n_neighbors = 5', 'min_dist = 0.5', 'random_state = 7']
    for ax, key, title in zip(axes.flat, keys, titles):
        data = pd.read_csv(OUTDIR / f'umap-{key}-coordinates.csv')
        for label in ['ALL', 'AML']:
            part = data[data.diagnosis == label]
            ax.scatter(part.UMAP1, part.UMAP2, s=20, color=COLORS[label], label=label, alpha=.85)
        ax.set(xlabel='UMAP1', ylabel='UMAP2', title=title)
        ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(OUTDIR / 'umap-settings.png', dpi=200, bbox_inches='tight')
    plt.close(fig)

make_pca_figure(); make_umap_figure()

doc = Document(SOURCE_DOCX)
doc.paragraphs[0].style = doc.styles['Title']
doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.paragraphs[0].paragraph_format.first_line_indent = Pt(0)
body = doc._element.body
paragraphs = doc.paragraphs
section3 = next(p for p in paragraphs if p.text.strip() == '3 实验设计')
section3.text = '3 任务解答与实验结果'

# Preserve the existing title, summary and research-question sections, then replace
# the old generic experiment-design section with task-by-task results.
seen = False
for child in list(body):
    if child is section3._p:
        seen = True
        continue
    if seen and child.tag != qn('w:sectPr'):
        body.remove(child)

def add_paragraph(text='', style='Normal', before=None, bold_prefix=None):
    if before is None:
        p = doc.add_paragraph(style=style)
    else:
        p = before.insert_paragraph_before(style=style)
    if bold_prefix and text.startswith(bold_prefix):
        p.add_run(bold_prefix).bold = True
        p.add_run(text[len(bold_prefix):])
    else:
        p.add_run(text)
    return p

def add_before(text, style='Normal', bold_prefix=None):
    return add_paragraph(text, style, before=section3, bold_prefix=bold_prefix)

def add_heading(text, level):
    return doc.add_paragraph(text, style=f'Heading {level}')

def set_border(parent, edge, val, sz=8, color='000000'):
    tag = qn('w:' + edge)
    el = parent.find(tag)
    if el is None:
        el = OxmlElement('w:' + edge); parent.append(el)
    el.set(qn('w:val'), val)
    if val != 'nil':
        el.set(qn('w:sz'), str(sz)); el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)

def add_table(headers, rows, font_size=10.5, widths=None):
    t = doc.add_table(rows=1, cols=len(headers)); t.style = 'Normal Table'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True
    for i, value in enumerate(headers): t.rows[0].cells[i].text = str(value)
    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row): cells[i].text = str(value)
    tblpr = t._tbl.tblPr
    tblb = tblpr.find(qn('w:tblBorders'))
    if tblb is None: tblb = OxmlElement('w:tblBorders'); tblpr.append(tblb)
    for edge in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'): set_border(tblb, edge, 'nil')
    for ri, row in enumerate(t.rows):
        if ri == 0:
            trpr = row._tr.get_or_add_trPr(); repeat = OxmlElement('w:tblHeader'); repeat.set(qn('w:val'), 'true'); trpr.append(repeat)
        for ci, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tcpr = cell._tc.get_or_add_tcPr()
            shade = tcpr.find(qn('w:shd'))
            if shade is not None: tcpr.remove(shade)
            cb = tcpr.find(qn('w:tcBorders'))
            if cb is None: cb = OxmlElement('w:tcBorders'); tcpr.append(cb)
            for edge in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'): set_border(cb, edge, 'nil')
            if ri == 0: set_border(cb, 'top', 'single', 12); set_border(cb, 'bottom', 'single', 8)
            if ri == len(t.rows)-1: set_border(cb, 'bottom', 'single', 12)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
                for run in p.runs:
                    set_run(run, '宋体', font_size, bold=(ri == 0))
            if widths and ci < len(widths): cell.width = Inches(widths[ci])
    doc.add_paragraph()
    return t

def set_run(run, east_asia, size, bold=None):
    run.font.name = 'Times New Roman'; run.font.size = Pt(size); run.font.color.rgb = __import__('docx').shared.RGBColor(0,0,0)
    if bold is not None: run.bold = bold
    rpr = run._element.get_or_add_rPr(); rf = rpr.rFonts
    if rf is None: rf = OxmlElement('w:rFonts'); rpr.insert(0, rf)
    rf.set(qn('w:ascii'), 'Times New Roman'); rf.set(qn('w:hAnsi'), 'Times New Roman'); rf.set(qn('w:eastAsia'), east_asia); rf.set(qn('w:cs'), 'Times New Roman')

def add_labelled(label, text):
    p = doc.add_paragraph(style='Normal'); p.add_run(label).bold = True; p.add_run(text); return p

def add_image(path, width=6.2):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent = Pt(0); p.add_run().add_picture(str(path), width=Inches(width)); return p

def add_caption(text):
    p = doc.add_paragraph(style='Normal'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent = Pt(0); p.add_run(text); return p

# Add the user's unrecorded pre-experiment fields and concise method primer before results.
add_before('实验前个人判断：________；信心（0—100）：________；最想补充的一项证据：________。这些是实验前记录，请按当时真实想法补填。')
add_before('PCA 与 UMAP 简介及计算原理', 'Heading 2')
add_before('PCA（主成分分析）把多个原始探针转换为较少的综合坐标。先按列中心化；需要时再标准化为 zᵢⱼ=(xᵢⱼ−μⱼ)/sⱼ。第一主成分选择单位向量 w₁，使所有样本投影 t₁=Zw₁ 的方差最大；后续主成分与前面方向正交，并依次解释剩余方差。计算上可对中心化矩阵做奇异值分解 Z=UΣVᵀ；样本坐标由投影得到，解释率由各方向方差除以总方差得到。')
add_before('UMAP（统一流形近似与投影）不直接寻找最大方差方向。它先根据高维样本距离为每个样本找近邻，形成带权的邻居关系图，再给每个样本安排二维坐标。算法迭代调整坐标，让高维中关系较强的近邻在二维中尽量靠近，同时避免所有点挤在一起。n_neighbors 影响高维邻域尺度，min_dist 影响二维点的紧密程度；随机种子可能改变初始化和最终布局。UMAP 坐标不是探针线性组合，也没有 PCA 同义的方差解释率。')
add_before('两种方法的共同点是压缩表示、都可能丢失信息；区别是 PCA 优先呈现线性投影方差，UMAP 优先近似呈现局部邻域。以下按工作单任务逐项报告实验设计、结果与解释。')

# Task 1
add_heading('3.1 任务一：明确降维目的与拟合范围', 2)
add_labelled('实验设计：', '依据工作单，将五种降维目标分别配对需要的输出和评价方式；再区分描述当前 72 个样本与预测未见样本两条流程。工作单列出的交叉验证比较属于可选验证，本次实际完成。')
add_table(['目标','需要的输出','如何评价是否有用'],[
 ['可视化','二维/三维样本坐标和标注图','能否观察、追踪关系；改变设置后主要关系是否稳健，不能只看图是否漂亮。'],
 ['去噪','保留信号的低维表示或重建数据','重复测量是否更一致、真实信号是否保留，或独立下游验证是否改善；仅图更平滑不足以证明。'],
 ['压缩特征','较少坐标及可应用的变换规则','权衡保留维数、解释率、重建误差、存储和计算成本。'],
 ['发现样本结构','候选邻域、梯度或分组','检查跨参数、批次和独立样本的稳定性，并结合外部信息。'],
 ['后续模型输入','训练集拟合的降维变换和下游模型','用无泄露交叉验证和独立测试评估预测表现与计算成本。']], font_size=10.5, widths=[1.15,2.25,3.15])
add_labelled('描述当前数据的流程：', '全部 72 个样本表达矩阵 → 在这 72 个样本拟合标准化 → PCA/UMAP 拟合并生成坐标 → 拟合完成后按诊断着色 → 描述当前样本结构。')
add_labelled('预测新样本的流程：', '先锁定训练/测试划分 → 每个训练折内拟合标准化、PCA（若用）与分类器 → 用已拟合流程转换并预测验证折 → 固定方案后以训练数据重拟合 → 最后评估独立测试集。')
add_labelled('为什么全数据 PCA 仍会泄露：', 'PCA 不用标签，但会从输入学习均值、协方差和主成分方向；StandardScaler 也从输入学习均值和尺度。若验证折参与这些拟合，验证数据的分布信息已进入模型流程。因此预测任务须把全部拟合步骤放在 Pipeline 中并在每个训练折内拟合。')
add_labelled('本次可选验证设计：', '仅使用原始 38 个训练样本，固定 StratifiedKFold(3, shuffle=True, random_state=42)，比较 StandardScaler→1近邻 与 StandardScaler→PCA(10, svd_solver="full")→1近邻，以 balanced accuracy 评价；原始 34 个测试样本未用于评分或调参。')
add_table(['验证折','标准化→1近邻','标准化→PCA(10)→1近邻'],[
 ['1','0.819444','0.694444'],['2','0.875000','0.763889'],['3','0.833333','0.833333'],['三折平均','0.842593','0.763889']], font_size=10.5, widths=[1.2,2.3,3.0])
add_labelled('结果解释：', '本次固定三折比较中，加入 PCA(10) 后平均验证 balanced accuracy 较低；这不表示 PCA 总会降低预测表现。样本量小，结果依赖当前划分和模型。balanced accuracy 是 ALL 与 AML 两类召回率的平均值，不是总正确率。两套 1 近邻模型的训练分数均为 1.0，训练满分不代表新样本泛化。')

# Task 2
add_heading('3.2 任务二：PCA 保留了什么，又丢掉了什么', 2)
add_labelled('实验设计：', '使用全部 72 行、7,129 个探针，不纳入 sample_id、original_split、diagnosis。按列标准化并检查均值与总体标准差；对标准化矩阵做 PCA，同时对仅中心化、未除以标准差的原尺度矩阵做对照。记录解释率、累计解释率、PC1 五个最大绝对权重探针及坐标。该任务描述当前数据，不是预测评估。')
add_labelled('预处理结果：', 'StandardScaler 的变换是 zᵢⱼ=(xᵢⱼ−μⱼ)/σⱼ。标准化后列均值绝对值最大为 5.98×10⁻¹⁶，各列总体标准差与 1 的差的绝对值最大为 3.33×10⁻¹⁶，属于浮点误差。此按列标准化不是微阵列上游归一化；数据含负值，本课不对其直接取 log 或 log1p。')
add_table(['PCA 输入','PC1 解释率','PC2 解释率','前两项合计','达到 80% 最少成分','达到 90% 最少成分'],[
 ['标准化','14.9216%','9.4462%','24.3678%','37','51'],['原尺度，仅中心化','17.4278%','10.3386%','27.7664%','20','34']], font_size=9.5, widths=[1.3,1.25,1.25,1.5,1.55,1.55])
add_image(OUTDIR/'pca-comparison.png', 6.35)
add_caption('图 1  标准化 PCA 与原尺度 PCA；轴括号为各自输入空间的解释率')
add_image(OUTDIR/'pca-variance.png', 5.8)
add_caption('图 2  两种 PCA 输入的累计解释率')
add_labelled('图和解释率：', '标准化 PCA 前两维合计解释 24.3678% 的总方差；其余约 75.63% 位于 PC3 及后续方向，二维图未显示。被省略的方差不等于样本在 PC1、PC2 上完全重合；样本可在二维图上分开，同时在未显示方向上仍有差异。两种输入的总方差基准不同，不能只凭解释率更高断定某一预处理更好。')
load = pd.read_csv(OUTDIR/'pca-loadings.csv')
rows=[]
for _, r in load.iterrows(): rows.append([r['input'],r['probe'],f"{r['weight']:+.8f}",f"{r['raw_sd']:.2f}"])
add_labelled('PC1 权重表：', '本报告将 pca.components_.T 定义为主成分方向权重（loading）；表中按 PC1 权重绝对值排序。')
add_table(['输入方案','探针','PC1 带符号权重','原始尺度标准差'],rows,font_size=9.5,widths=[1.45,2.25,1.7,1.75])
add_labelled('PC1 如何计算：', '对样本 i，PC1 得分是全部 7,129 个输入特征的加权和：PC1ᵢ=Σⱼ wⱼzᵢⱼ。标准化 PCA 中 X92518_s_at 的权重约 +0.02878558，它只贡献总和的一项；高权重不等于致病、差异显著或分类最重要。')
add_labelled('尺度对照解释：', '两种 PCA 的前五探针完全不同。原尺度前五探针的原始标准差约 6,846.94–15,782.80；标准化方案所列前五约 326.91–531.14。原尺度中的大波动特征更容易影响总方差方向；标准化让各探针以自身标准差为单位比较，但也可能相对放大低变异噪声。')
add_labelled('概念题结果：', 'PC1 在单位长度方向中最大化样本投影方差，不是使用 ALL/AML 标签找出的最佳分类方向，也不是单一基因。若把所有主成分权重和得分同时乘以 −1，只会翻转坐标轴符号，不改变解释率和样本几何关系。中心化后 72 个样本偏差之和为零，样本行至少有一个线性依赖关系，因此矩阵秩最多 min(72−1,7129)=71，最多有 71 个非零主成分。')
add_labelled('一句解释 PC1：', '标准化 PC1 是全部探针共同构成的线性方向，解释本数据总方差的 14.9216%。')
add_labelled('一句不能推出的结论：', '不能由 PC1 权重或二维图中的诊断趋势推断该探针致病、发现真实疾病机制，或保证对新患者预测准确。')

# Task 3
add_heading('3.3 任务三：UMAP 上的结构可以相信到什么程度', 2)
add_labelled('实验设计：', '沿用任务二同一个 72×7,129 标准化矩阵，不按诊断筛选探针，拟合二维 UMAP；度量为欧氏距离，n_jobs=1。基线设 n_neighbors=15、min_dist=0.1、random_state=42，再各自只改变邻居数、min_dist 或随机种子。各次坐标均与 sample_id 对应保存，避免只挑最清楚的一张图。')
add_table(['设置','n_neighbors','min_dist','种子','原空间5近邻保留','与基线二维5近邻重合'],[
 ['基线','15','0.1','42','41.11%','100.00%'],['仅改邻居数','5','0.1','42','45.56%','55.28%'],['仅改 min_dist','15','0.5','42','40.00%','64.17%'],['仅改随机种子','15','0.1','7','40.56%','60.28%']],font_size=9.5,widths=[1.55,1.15,1.05,0.85,1.8,2.0])
add_image(OUTDIR/'umap-settings.png', 5.7)
add_caption('图 3  基线 UMAP 与三个单因素参数/随机种子对照')
add_labelled('指标如何读：', '“原空间5近邻保留”是标准化 7,129 维空间与相应二维图的近邻名单平均重合；“与基线重合”是两张二维 UMAP 图的近邻名单平均重合。二者均按每个样本各取 5 个近邻、计算交集占 5 的比例后再对 72 个样本取平均。它们是本报告的描述性辅助量，不是分类准确率、UMAP 方差解释率或标准 trustworthiness 指标。')
stable = pd.read_csv(OUTDIR/'umap-stable-neighbors.csv')
g4 = stable.loc[stable.sample_id=='G004'].iloc[0]
g8 = stable.loc[stable.sample_id=='G008'].iloc[0]
g10 = stable.loc[stable.sample_id=='G010'].iloc[0]
add_labelled('稳定邻域结果：', f"四组二维嵌入的 5 近邻交集中，G004 的共同邻居为 {g4.all_settings_neighbors.replace(';', '、')}；G008 的共同邻居为 {g8.all_settings_neighbors.replace(';', '、')}。G010 的共同邻居数为 {int(g10['count'])}，说明不同样本的邻域稳定程度不一。这里的稳定只针对本次四种设置。")
add_labelled('PCA 与 UMAP 比较：', '下表按工作单要求概括两种方法；预测用途都必须在训练折内拟合变换。')
add_table(['比较项','PCA','UMAP'],[
 ['线性/非线性','线性投影','非线性嵌入'],
 ['主要保留信息','依次最大化正交方向的投影方差；固定维数下最小化平方重建误差','建立高维近邻图并优化低维布局，侧重局部邻域关系'],
 ['坐标轴含义','特征的带权线性组合，可查看权重与得分','嵌入坐标，没有固定的基因或生物学单位'],
 ['方差解释率','可计算并需注明输入预处理','没有与 PCA 等价的 explained variance ratio'],
 ['局部与全局距离','二维只保留部分线性子空间；截断后距离仍可能失真','邻近关系是局部线索；远处岛屿距离、方向、面积不能直接作生物学量'],
 ['参数与随机性','对标准化和成分数敏感；本次 full SVD 确定到符号翻转','受 n_neighbors、min_dist、度量和随机种子影响'],
 ['预测时避免泄露','Scaler 与 PCA 只在每折训练数据拟合，再变换验证数据','Scaler 与 UMAP 也须训练折拟合；仅用已拟合模型转换验证数据']],font_size=10,widths=[1.55,2.75,2.75])
add_labelled('岛屿与单细胞展示：', '本工作单没有定义聚类算法或岛屿边界，因此图上“有几团”不是正式簇数。岛屿数量、相互距离、方向和面积会受布局、参数与随机性影响，不能直接当作亚型数、疾病阶段或群体比例。UMAP 常用于单细胞或复杂表型的探索性展示，因为许多局部非线性关系不容易由两条线性 PCA 轴呈现；它帮助提出结构假设，但不替代验证。')
add_labelled('批次边界：', '本数据没有可用批次字段，无法检验或调整批次效应，因此不能宣称已排除批次影响。样本相近或分离可能来自生物差异，也可能来自技术处理差异。')
add_labelled('图能支持／不能支持：', '（1）支持：当前设定下有些样本在局部邻域中较相近；不能推出所有原空间邻居都被完整保留。（2）支持：诊断着色可帮助探索表达排列是否与已知标签相关；不能凭分离证明疾病机制或真实亚型数量。（3）支持：图可帮助提出后续建模或实验问题；不能据此报告新患者预测能力，也不能排除批次影响。')

# Task 4 and required homework responses
add_heading('3.4 任务四：研究团队结论与三道作业', 2)
add_labelled('主张：', '二维降维图显示当前样本中存在值得进一步检验的表达结构，但不能单独证明疾病机制或预测新患者。')
add_labelled('证据：', '标准化 PCA 前两维合计解释 24.3678% 的方差；只将 UMAP 的 n_neighbors 从 15 改为 5 后，与基线二维 5 近邻平均重合为 55.28%。')
add_labelled('推理：', '二维投影遗漏大量高维变异，UMAP 的部分近邻也会随设定变化，因此图中的分离是候选线索，不是完整结构或因果证据。')
add_labelled('边界与后续验证：', '本数据仅有 72 个样本且缺少批次字段。应补充批次和患者独立性信息，在预先固定完整 Pipeline 的独立队列中验证预测；若提出机制，还需针对候选分子机制做额外实验验证。')
add_labelled('作业 1：PCA 第一主成分代表什么？', '它是使样本投影方差最大的单位线性方向，样本得分由所有输入探针加权得到。本次标准化 PC1 解释总方差 14.9216%；探针 X92518_s_at 的方向权重约为 +0.02878558，只是 PC1 计算中的一项，不代表该探针对疾病有因果作用。')
add_labelled('作业 2：为什么 PCA 前考虑标准化？', 'PCA 依赖方差，原尺度波动大的探针更容易影响主方向。本次标准化与原尺度 PCA 的 PC1 前五探针完全不同，说明尺度处理改变了关注方向。标准化让探针按各自标准差比较，代价是低变异噪声可能被相对放大；须结合数据和分析目标，不以解释率单项选优。')
add_labelled('作业 3：UMAP 图能说明什么、不能说明什么？', 'UMAP 可以展示当前预处理和设置下的局部邻域，帮助提出候选结构。本次仅改随机种子后，与基线二维 5 近邻平均重合为 60.28%，说明部分邻居保持、布局并非唯一。图不能单独证明机制、真实亚型数或新患者预测准确率，也不能把全局岛屿距离当成生物距离。')

# Reproducibility section with actual versions and output list.
add_heading('4 数据与复现', 1)
add_labelled('数据与拟合范围：', 'Golub 白血病表达数据课程副本，72 个样本、7,129 个探针；ALL 47、AML 25。描述性 PCA/UMAP 使用全部 72 个样本，诊断只在拟合后用于着色；交叉验证只用原始 38 个训练样本，每折 Pipeline 内拟合。未用原始测试集分数选方案。')
add_labelled('软件版本：', 'Python 3.12.15；NumPy 2.3.5；pandas 2.3.3；SciPy 1.17.1；scikit-learn 1.8.0；umap-learn 0.5.12；Matplotlib 3.10.8。')
add_labelled('参数记录：', '标准化为按列中心化并按总体标准差缩放（ddof=0）；PCA 使用 full SVD；UMAP 基线为 n_neighbors=15、min_dist=0.1、metric=euclidean、random_state=42、n_jobs=1。对照分别改 n_neighbors=5、min_dist=0.5、random_state=7。交叉验证为三折分层，shuffle=True、种子 42。')
add_labelled('可复现文件：', '完整分析代码为 .tmp/dimension-reduction/student_analysis.py；运行摘要为 summary.json；交叉验证表为 cross-validation.csv；PCA 解释率与权重表为 pca-variance.csv、pca-loadings.csv；样本坐标和 UMAP 稳定邻居表均保存在 .tmp/dimension-reduction/student-results/。运行图和表可由代码重新生成。')
add_labelled('输入文件校验：', f"leukemia.csv 的 SHA-256 为 {summary['csv_sha256']}。")
add_heading('5 参考资料', 1)
add_paragraph('Golub TR, et al. Molecular Classification of Cancer: Class Discovery and Class Prediction by Gene Expression Monitoring. Science. 1999;286(5439):531–537. doi:10.1126/science.286.5439.531.')
add_paragraph('McInnes L, Healy J, Melville J. UMAP: Uniform Manifold Approximation and Projection for Dimension Reduction. arXiv:1802.03426.')

# Apply requested typography and black text globally; retain the section titles in parts 1 and 2.
styles = doc.styles
def set_style(style, east, size, bold=False):
    style.font.name='Times New Roman'; style.font.size=Pt(size); style.font.bold=bold; style.font.color.rgb=__import__('docx').shared.RGBColor(0,0,0)
    rpr=style._element.get_or_add_rPr(); rf=rpr.rFonts
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for key,val in [('ascii','Times New Roman'),('hAnsi','Times New Roman'),('cs','Times New Roman'),('eastAsia',east)]: rf.set(qn('w:'+key),val)
set_style(styles['Normal'],'宋体',12)
styles['Normal'].paragraph_format.line_spacing=1.25
styles['Normal'].paragraph_format.first_line_indent=Pt(24)
set_style(styles['Heading 1'],'黑体',15,True)
styles['Heading 1'].paragraph_format.first_line_indent=Pt(0); styles['Heading 1'].paragraph_format.keep_with_next=True
styles['Heading 1'].paragraph_format.space_before=Pt(11); styles['Heading 1'].paragraph_format.space_after=Pt(5)
set_style(styles['Heading 2'],'黑体',12,True)
styles['Heading 2'].paragraph_format.first_line_indent=Pt(0); styles['Heading 2'].paragraph_format.keep_with_next=True
styles['Heading 2'].paragraph_format.space_before=Pt(8); styles['Heading 2'].paragraph_format.space_after=Pt(4)
set_style(styles['Title'],'黑体',22,True)

for p in doc.paragraphs:
    name=p.style.name
    if name=='Title': east,size='黑体',22; p.paragraph_format.first_line_indent=Pt(0)
    elif name=='Heading 1': east,size='黑体',15; p.paragraph_format.first_line_indent=Pt(0)
    elif name=='Heading 2': east,size='黑体',12; p.paragraph_format.first_line_indent=Pt(0)
    else:
        east,size='宋体',12
        if p.alignment == WD_ALIGN_PARAGRAPH.CENTER: p.paragraph_format.first_line_indent=Pt(0)
        else: p.paragraph_format.first_line_indent=Pt(24)
        p.paragraph_format.line_spacing=1.25
    for run in p.runs: set_run(run,east,size,bold=run.bold if run.bold is not None else (name in ('Title','Heading 1','Heading 2')))

# Set every table cell to centered text, no fill, black font, and true three-line rules.
for table in doc.tables:
    table.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblpr=table._tbl.tblPr; tblb=tblpr.find(qn('w:tblBorders'))
    if tblb is None: tblb=OxmlElement('w:tblBorders'); tblpr.append(tblb)
    for edge in ('top','bottom','left','right','insideH','insideV'): set_border(tblb,edge,'nil')
    for ri,row in enumerate(table.rows):
        if ri==0:
            trpr=row._tr.get_or_add_trPr(); rep=OxmlElement('w:tblHeader'); rep.set(qn('w:val'),'true'); trpr.append(rep)
        for cell in row.cells:
            tcpr=cell._tc.get_or_add_tcPr(); shd=tcpr.find(qn('w:shd'))
            if shd is not None: tcpr.remove(shd)
            cb=tcpr.find(qn('w:tcBorders'))
            if cb is None: cb=OxmlElement('w:tcBorders'); tcpr.append(cb)
            for edge in ('top','bottom','left','right','insideH','insideV'): set_border(cb,edge,'nil')
            if ri==0: set_border(cb,'top','single',12); set_border(cb,'bottom','single',8)
            if ri==len(table.rows)-1: set_border(cb,'bottom','single',12)
            for p in cell.paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent=Pt(0); p.paragraph_format.line_spacing=1.0; p.paragraph_format.space_after=Pt(0)
                for run in p.runs: set_run(run,'宋体',10.5,bold=(ri==0))

# Ensure all references and any existing title properties resolve to black.
for section in doc.sections:
    for part in (section.header,section.footer):
        for p in part.paragraphs:
            for run in p.runs: set_run(run,'宋体',12)

doc.save(DOCX)
print(DOCX)
print('paragraphs',len(doc.paragraphs),'tables',len(doc.tables),'figures',len(doc.inline_shapes))
