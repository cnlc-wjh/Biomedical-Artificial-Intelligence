from pathlib import Path
import json,hashlib
from copy import deepcopy
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

ROOT=Path(__file__).resolve().parents[2]
SRC=ROOT/'作业报告/多重检验校正报告.docx'
OUT=ROOT/'作业报告/多重检验校正报告_图文终稿.docx'
TMP=ROOT/'.tmp/multiple-test'
source_hash=hashlib.sha256(SRC.read_bytes()).hexdigest()
doc=Document(SRC)
ps=list(doc.paragraphs)
replacements={
3:'本报告比较 Golub 数据中 AML 与 ALL 两种白血病亚型的平均表达。发现阶段包括 38 份样本，共检验 7,129 个探针。未经校正时，1,694 个探针的 p≤0.05；对全部探针校正后，Bonferroni 显著 26 个，BH 显著 535 个。',
4:'随后从 BH 显著结果中按既定规则选取 10 个候选，在剩余 34 份历史留出样本中检查。7 个候选同时满足方向一致和候选范围内 Bonferroni 校正后 p≤0.05。这些结果可作为后续研究线索，但尚不足以证明致病机制或临床价值。',
8:'数据共 72 份样本、7,129 个探针。sample_id 为样本编号，original_split 标记发现或留出部分，diagnosis 标记诊断组别，三列均不参与探针检验。每个探针是一项表达测量，不一定对应一个独立基因；数据中没有健康对照。',
9:'核查未发现缺失值、非有限值、重复编号、完全重复的表达行或常数探针。表达值范围为 −28,400 至 71,369。本次保留原有数值，不取对数，也不将负值作为缺失值处理。现有信息不足以排除批次、中心及样本来源的影响。',
10:'每个探针均检验两组总体平均表达是否相同，采用双侧 Welch t 检验，允许两组方差不同。差值统一记为 AML−ALL。发现阶段仅使用 train 样本；全部 7,129 个探针共同构成检验家族，即需要一起校正的一组检验。显著性水平为 0.05。',
13:'本研究的原假设是 AML 与 ALL 的总体平均表达相同。p 值衡量：在这一假设和检验模型成立时，出现当前或更极端检验结果的概率。双侧检验同时考察表达升高和降低。',
14:'例如，原始 p＝0.03 达到单项 0.05 显著性标准，提示存在反对均值相等的统计证据。它不表示原假设有 3% 的概率成立，也不表示该探针有 97% 的概率存在真实差异。',
15:'由于本研究同时检验数千个探针，不能仅凭原始 p≤0.05 确认结果，需要进行多重校正。相反，p>0.05 只表示本次证据未达到标准，不能证明两组没有差异。',
17:'若 1,000 个变量实际上均无差异，每项检验的假阳性概率为 0.05，则平均会出现 1,000×0.05＝50 个原始显著结果。50 是重复实验的平均数，单次结果可以高于或低于它。这个期望不要求各检验独立；相同条件下检验 5,000 项，平均假发现数为 250。',
18:'因此，案例中“1,000 个基因有 53 个原始显著结果”不足以支持“发现 53 个疾病相关基因”。还需检查完整校正结果、差异大小和验证证据；也不能仅凭这一数量认定 53 个结果全部错误。',
20:'用 R 表示一次分析判定显著的总数，V 表示其中的假发现数，S 表示真发现数，因此 R＝V＋S。有发现时，错误发现比例 FDP＝V/R；没有发现时记为 0。FWER 是至少出现一个假发现的概率，FDR 是多次重复实验中 FDP 的平均水平。',
21:'例如，某次模拟发现 20 项，其中 2 项为假，则该次 FDP＝10%，并且发生了至少一次假发现。估计 FDR 时，先计算每次 FDP，再对这些比例求平均，不能通常用平均 V 除以平均 R 代替。',
22:'检出力表示真实信号中有多少被检出。例如，100 个真信号中检出 30 个，检出力为 30%。全零实验没有真信号，检出力不适用，记为 N/A。真实数据没有真值清单，因此无法直接数出真实假发现或计算真实检出力。',
26:'Bonferroni 将原始 p 乘以检验总数 m，结果超过 1 时记为 1，再与 0.05 比较。',
27:'也可直接将原始 p 与 0.05/m 比较，二者给出相同结论。本研究 m＝7,129，因此原始 p 的判定界限为 7.0136×10⁻⁶。校正后 p 仍与 0.05 比较，无需再次除以 m。',
28:'在单项 p 值有效的条件下，Bonferroni 将整组至少出现一次假发现的概率控制在目标水平内，不要求检验相互独立。其标准较严格，也可能漏掉较弱的真实差异。',
30:'BH 先将原始 p 从小到大排序，再寻找满足 p₍k₎≤k×0.05/m 的最大位置 k，纳入前 k 项。排序界限随位置变化，不能只将每行单独比较后决定是否纳入。',
32:'BH 校正后 p 可按上式计算，再与 0.05 比较，得到相同的发现集合。BH 关注显著结果中假发现比例的长期平均值。在单项检验有效且检验独立或满足适当正依赖条件时，它能提供相应控制；真实探针的相关结构仍需审查。',
35:'表2中，Bonferroni 校正后只有前 2 项≤0.05；BH 的最大满足位置为第 4 项，因此前 4 项显著。两种方法都使用 0.05 判断校正后 p，但控制的错误不同，结果无需相同。',
36:'本报告计算的是 BH 校正后 p。它不能直接当作某个结果为假的概率，也不等同于另行估计零假设比例的 Storey q-value。',
40:'图1与表3显示，两种校正均减少了显著探针数。BH 最后纳入第 535 项，其原始 p＝0.00374664，低于该位置界限 0.00375228。三种数量的差值不能当作真实假阳性数，535 个显著探针也不能直接称为 535 个疾病基因。',
42:'以原始 p 最小的探针 U22376_cds2_s_at 为例，表4报告发现阶段的两组统计量，图2展示发现和留出样本中的实际测量值。',
45:'该探针在发现样本中，AML 平均表达比 ALL 低 3,160.93。95%置信区间未跨越 0，与负向差异一致。此区间只针对该项检验，未对全部探针及筛选过程作覆盖率调整。表达值沿用数据原有尺度，不能将均值差写成倍数变化。',
47:'除原始尺度的均值差外，本报告用 Cohen’s d 描述标准化差异。它将均值差除以合并组内标准差，便于比较差异相对于组内波动的大小。',
50:'本例合并组内标准差约为 1,710.07，Cohen’s d≈−1.8484，即 AML 均值比 ALL 低约 1.85 个合并组内标准差。负号来自 AML−ALL 的方向约定。',
51:'p 值反映统计证据，d 描述差异大小。d 不是表达倍数或真实性概率，不能单凭其数值判断临床意义。两组标准差相差较大，因此合并标准差仅作描述；Welch 检验仍按各组方差分别计算标准误。',
53:'该探针原始 p 在全部 7,129 项中排名第一，Bonferroni 结果为：',
55:'BH 会比较从当前位置到最后一项的调整量 m×p₍j₎/j。本次最小调整量恰好出现在第 1 项，因此首位探针的 BH 与 Bonferroni 校正后 p 均为 5.9803×10⁻⁵。这是该数据的结果，不表示两种方法相同，也不保证排名第一时总会相等。',
56:'表2第 2 项可作对照：原始 p＝0.009，Bonferroni 校正后为 0.045，BH 为 0.0225。二者的计算规则及控制目标均不同。',
58:'每次模拟包含 1,000 个独立正态变量，两组样本量为 11 和 27，组内标准差设为 1。随机种子为 42，每种设置重复 1,000 次。模拟数据已知正态且等方差，采用双侧等方差 t 检验；真实数据采用 Welch 检验。每次均对同一组 p 值比较三种规则。',
61:'实验A的全部变量都没有真实差异，因此所有显著结果均为假发现。未经校正时，每次平均发现 49.944 项；Bonferroni 和 BH 分别为 0.051 和 0.052 项。此处有发现时 FDP＝1，无发现时为0，所以 FDR 与 FWER 相同。0.051 是有限重复的估计，与0.05的微小偏差不足以说明控制失效。',
64:'实验B中，前100个变量存在真实差异，其余900个无差异。δ表示两组真实均值差，分别设为1.0和0.5。图3显示：δ＝1.0时，BH的FDR约4.61%，检出力约28.65%；Bonferroni的检出力约5.96%。BH保留更多真信号，但其至少一次假发现的发生率更高，FWER为75.4%。',
65:'当均值差减为0.5、组内波动保持不变时，BH和Bonferroni检出力分别降至约0.335%和0.226%。真实信号仍有100个，因此检验不显著不能证明信号不存在。不同效应设置复用相同随机噪声，以便比较信号强弱的影响。',
68:'先从7,129项中选出原始p最小的20项，再只对20项校正，会忽略前面的筛选过程。本例完整校正后p约为5.9803×10⁻⁵，错误地只校正前20项则约为1.6777×10⁻⁷。更小的数值来自缩小了校正范围，并非增加了证据。',
69:'如果20项是在查看当前数据前由外部依据确定，且研究只检验这些项目，才可能使用m＝20。不能因为某个结果在两种流程中都显著，就认为事后缩小范围合理。',
71:'发现阶段在BH显著探针中，按原始p、探针名升序固定前10项及其差异方向，再检查test数据。留出阶段只检验这10项，并在候选范围内作Bonferroni校正。校正后p≤0.05且方向一致，才计为通过。本次为课堂数据的复现，不构成新开展的盲法验证。',
74:'10项在留出样本中的差异方向均与发现阶段一致，其中7项通过校正后的显著性标准。首位探针的留出差值为−1,766.73，Bonferroni校正后p＝0.023186。其余3项未达到标准，不能据此证明没有差异，也不应在查看结果后替换候选。',
75:'留出阶段的m＝10只适用于已固定的候选，不能用于重新校正发现阶段。test来自同一研究的历史样本；独立中心或新队列的验证仍需另行开展。',
78:'对于探索性组学筛选，目标是形成后续验证的候选名单，可考虑用BH控制FDR，同时报告效应和区间。它通常能保留更多候选，但后续研究仍需承担误选项目的成本。',
79:'对于确认性研究，若任意一个主要终点误报都会影响结论，应优先考虑FWER控制。例如预先确定5个主要终点时，可用Bonferroni将各原始p与0.01比较。具体方案应在查看结果前确定。',
80:'亚组分析、多模型尝试及选择性报告也会增加检验机会。校正范围应覆盖支撑研究结论的相关检验，不能在看到结果后只保留显著部分。',
82:'多重校正只处理多次检验带来的错误风险，不能消除批次差异、混杂或不独立样本的影响。本研究样本较少，探针之间也可能相关；现有字段不足以全面检查这些问题。本次未进行批次调整、稳健性重分析或外部验证。',
83:'按最小p值筛出的候选，其发现阶段效应可能被高估；表4的逐项置信区间也未处理筛选影响。后续应核实探针注释和测量质量，在独立数据中复测差异，并审查混杂因素。机制及临床用途还需专门研究。',
85:'在38份发现样本中，535个探针通过BH校正。固定10个候选后，7个在历史留出样本中方向一致且通过候选范围内Bonferroni校正。这些结果支持进一步研究AML与ALL的表达差异，但不能据此确认独立疾病基因数量、因果机制或临床有效性。'
}

def write(p,text):
    # Keep the user's paragraph formatting; replace only the requested wording.
    rp=deepcopy(p.runs[0]._r.rPr) if p.runs and p.runs[0]._r.rPr is not None else None
    p.clear();r=p.add_run(text)
    if rp is not None:r._r.insert(0,rp)
    r.font.name='Times New Roman';r.font.size=Pt(10.5);r.font.bold=False
    rf=r._r.get_or_add_rPr().get_or_add_rFonts();rf.set(qn('w:eastAsia'),'宋体')
    for key in list(rf.attrib):
        if 'theme' in key.lower():del rf.attrib[key]
for i,text in replacements.items():write(ps[i],text)

# Preserve user edits in the tables and bibliography; streamline section labels.
titles={2:'1 摘要',5:'2 数据与方法',11:'3 显著性与多重检验',12:'3.1 p值的含义',16:'3.2 多次检验与假阳性',19:'3.3 错误率与检出力',23:'4 校正方法与手算示例',24:'4.1 Bonferroni校正',29:'4.2 BH校正',37:'5 表达数据分析结果',41:'5.1 候选探针的表达差异',46:'5.2 标准化效应量',52:'5.3 两种校正结果相等的原因',57:'6 模拟实验结果',66:'7 候选筛选与留出验证',67:'7.1 校正范围与选择偏倚',70:'7.2 留出样本检验',76:'8 讨论与结论',77:'8.1 方法的选择',81:'8.2 研究局限',84:'8.3 结论'}
for i,title in titles.items():
    p=ps[i];p.text=title
    if i==46:p.style=doc.styles['Heading 2']
    size=15 if p.style.name=='Heading 1' else 12
    for r in p.runs:
        r.font.name='Times New Roman';r.font.size=Pt(size);r.font.bold=True;r.font.color.rgb=RGBColor(0,0,0)
        rf=r._r.get_or_add_rPr().get_or_add_rFonts();rf.set(qn('w:eastAsia'),'黑体')
        for key in list(rf.attrib):
            if 'theme' in key.lower():del rf.attrib[key]

def after(anchor,text='',style='Normal'):
    el=OxmlElement('w:p');anchor._p.addnext(el);p=Paragraph(el,anchor._parent);p.style=doc.styles[style]
    if text:write(p,text)
    return p

def figure(anchor,name,caption,note):
    p=after(anchor);p.paragraph_format.first_line_indent=Cm(0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next=True;p.paragraph_format.space_before=Pt(6)
    p.add_run().add_picture(str(TMP/'submission-figures'/name),width=Cm(15.0))
    cap=after(p,caption,'Caption');cap.alignment=WD_ALIGN_PARAGRAPH.CENTER;cap.paragraph_format.first_line_indent=Cm(0);cap.paragraph_format.keep_with_next=True
    n=after(cap,note);n.paragraph_format.space_after=Pt(6)
    return n

figure(ps[40],'discovery.png','图1 不同判定规则下的显著探针数','三种规则使用相同的38份发现样本和7,129个探针。显著数反映规则差异，不能直接用于计算真实假发现数。')
figure(ps[45],'probe_expression.png','图2 候选探针在发现及留出样本中的表达分布','每个点代表一份样本，黑色横线表示组均值。两部分数据中AML的平均表达均低于ALL。图示探针为U22376_cds2_s_at，表达值未经额外变换。')
figure(ps[65],'simulation.png','图3 信号强弱对错误发现率和检出力的影响','图中为1,000次重复模拟的汇总，δ为两组真实均值差，噪声标准差固定为1。信号较弱时，三种方法的检出力均降低；左图虚线标出5%的FDR目标水平。')
figure(ps[75],'validation.png','图4 候选探针在发现及留出样本中的均值差','编号按表7顺序排列；绿色表示满足留出判据，灰色表示未满足。所有点位于第一或第三象限，说明方向均一致；方向一致仍不等于达到校正后的显著性标准。')

# Keep original page breaks at major sections, but let the longer illustrated
# result sections flow naturally. Prevent tables and their captions from separating.
for p in doc.paragraphs:
    if p.style.name=='Heading 2':p.paragraph_format.page_break_before=False
    if p.text.startswith(('6 模拟','7 候选','8 讨论')):p.paragraph_format.page_break_before=False
    p.paragraph_format.widow_control=True
    if p._p.xpath('.//m:oMath'):p.paragraph_format.keep_with_next=True
for t in doc.tables:
    for ri,row in enumerate(t.rows):
        pr=row._tr.get_or_add_trPr()
        if pr.find(qn('w:cantSplit')) is None:pr.append(OxmlElement('w:cantSplit'))
        for cell in row.cells:
            for p in cell.paragraphs:p.paragraph_format.keep_with_next=ri<len(t.rows)-1
    header=t.rows[0]._tr.get_or_add_trPr()
    if header.find(qn('w:tblHeader')) is None:header.append(OxmlElement('w:tblHeader'))
for st in doc.styles:
    for rf in st.element.xpath('./w:rPr/w:rFonts'):
        for key in list(rf.attrib):
            if 'theme' in key.lower():del rf.attrib[key]
doc.save(OUT)
assert hashlib.sha256(SRC.read_bytes()).hexdigest()==source_hash
(TMP/'submission-edit-log.json').write_text(json.dumps({'source':str(SRC),'source_sha256':source_hash,'output':str(OUT),'rewritten_paragraphs':len(replacements),'figures':4,'source_tables_preserved':[[[c.text for c in r.cells] for r in t.rows] for t in doc.tables]==[[[c.text for c in r.cells] for r in t.rows] for t in Document(SRC).tables],'bibliography_preserved':all(p.text in [q.text for q in doc.paragraphs] for p in Document(SRC).paragraphs if p.text.startswith('['))},ensure_ascii=False,indent=2),encoding='utf-8')
print(OUT)
