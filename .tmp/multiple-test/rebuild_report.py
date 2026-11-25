from pathlib import Path
from copy import deepcopy
import json, csv, hashlib, shutil, zipfile
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[2]
TMP=ROOT/'.tmp/multiple-test'
RESULT=TMP/'student-results'
OUT=ROOT/'作业报告'
OUT.mkdir(exist_ok=True)
REF=ROOT.parent/'5. Hypothesis test/统计检验与假设设计报告-吴际恒.docx'
DOC=OUT/'多重检验校正报告_字体公式修正版.docx'
s=json.loads((RESULT/'analysis_summary.json').read_text(encoding='utf-8'))
rows=list(csv.DictReader((RESULT/'discovery_all.csv').open(encoding='utf-8-sig')))
valid=list(csv.DictReader((RESULT/'validation.csv').open(encoding='utf-8-sig')))
sims=list(csv.DictReader((RESULT/'simulation_summary.csv').open(encoding='utf-8-sig')))
e=s['example']
ref_hash=hashlib.sha256(REF.read_bytes()).hexdigest()
doc=Document(REF)
# A complete content replacement in a copy of the course reference, retaining section and style definitions.
for child in list(doc._element.body):
    if child.tag!=qn('w:sectPr'): doc._element.body.remove(child)
for rel in list(doc.part.rels.values()):
    if any(rel.reltype.endswith('/'+t) for t in ['image','hyperlink','comments']):
        doc.part.drop_rel(rel.rId)
for sec in doc.sections:
    for part in [sec.header,sec.footer]:
        for child in list(part._element): part._element.remove(child)
    sec.footer_distance=Cm(1.25)

def font(run,size=10.5,bold=False,east='宋体'):
    run.font.name='Times New Roman'; run.font.size=Pt(size)
    run.font.bold=bold; run.font.color.rgb=RGBColor(0,0,0)
    rf=run._element.get_or_add_rPr().get_or_add_rFonts()
    for key in list(rf.attrib):
        if 'theme' in key.lower():del rf.attrib[key]
    rf.set(qn('w:eastAsia'),east)
    return run

for name,size,east,bold in [('Normal',10.5,'宋体',False),('Title',20,'黑体',True),('Heading 1',15,'黑体',True),('Heading 2',12,'黑体',True),('Caption',10.5,'宋体',False)]:
    st=doc.styles[name]; st.font.name='Times New Roman'; st.font.size=Pt(size)
    st.font.color.rgb=RGBColor(0,0,0); st.font.bold=bold
    st.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),east)
    f=st.paragraph_format; f.line_spacing=1.25
    f.space_before=Pt(0); f.space_after=Pt(5)
    f.first_line_indent=Cm(.74) if name=='Normal' else Cm(0)
    if name.startswith('Heading'): f.keep_with_next=True; f.space_before=Pt(9)
    for el in st.element.xpath('./w:pPr/w:pBdr'): el.getparent().remove(el)

# Theme bindings take precedence over explicit font names in Word. Remove them
# from all inherited and linked styles rather than setting only eastAsia.
for st in doc.styles:
    for rf in st.element.xpath('./w:rPr/w:rFonts'):
        for key in list(rf.attrib):
            if 'theme' in key.lower():del rf.attrib[key]
for name in ['Title','Heading 1','Heading 2']:
    st=doc.styles[name]
    st.font.italic=False
    szcs=st.element.get_or_add_rPr().find(qn('w:szCs'))
    if szcs is not None:szcs.set(qn('w:val'),str(round(st.font.size.pt*2)))

settings=doc.settings.element
for tag in ['hideSpellingErrors','hideGrammaticalErrors']:
    node=settings.find(qn('w:'+tag))
    if node is None: node=OxmlElement('w:'+tag); settings.append(node)
    node.set(qn('w:val'),'true')
doc.core_properties.title='Golub白血病表达数据多重检验校正报告'
doc.core_properties.subject='Bonferroni与BH校正 效应大小与留出验证'
doc.core_properties.author=''
doc.core_properties.comments=''

def para(text='',style=None,indent=True,size=10.5):
    p=doc.add_paragraph(style=style)
    if style is None:
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        if not indent:p.paragraph_format.first_line_indent=Cm(0)
        font(p.add_run(text),size)
    else:p.add_run(text)
    p.paragraph_format.widow_control=True
    return p

def heading(text,level=1,new=False):
    p=para(text,'Heading '+str(level))
    for r in p.runs:font(r,15 if level==1 else 12,True,'黑体');r.font.italic=False
    if new:p.paragraph_format.page_break_before=True
    return p

def mr(text,roman=False):
    r=OxmlElement('m:r')
    if roman:
        pr=OxmlElement('m:rPr');sty=OxmlElement('m:sty');sty.set(qn('m:val'),'p');pr.append(sty);r.append(pr)
    t=OxmlElement('m:t');t.text=text;r.append(t);return r

def sub(base,idx):
    node=OxmlElement('m:sSub');e=OxmlElement('m:e');e.append(mr(base));s=OxmlElement('m:sub');s.append(mr(idx,True));node.extend([e,s]);return node

def frac(num,den):
    f=OxmlElement('m:f')
    for tag,seq in [('num',num),('den',den)]:
        el=OxmlElement('m:'+tag)
        for x in seq:el.append(mr(x) if isinstance(x,str) else x)
        f.append(el)
    return f

def formula(text):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent=Cm(0); p.paragraph_format.space_after=Pt(7)
    math=OxmlElement('m:oMath')
    if text.startswith('pBonf'):
        nodes=[sub('p','Bonf'),mr(' = '),mr('min',True),mr('(m × p, 1) ;   '),sub('p','Bonf'),mr(' ≤ α')]
    elif text.startswith('pBH'):
        nodes=[sub('p','BH,(i)'),mr(' = '),mr('min',True),mr('{1, '),sub('min','j ≥ i'),mr(' '),frac([mr('m × '),sub('p','(j)')],[mr('j')]),mr('}')]
    elif text.startswith('sₚ'):
        rad=OxmlElement('m:rad');pr=OxmlElement('m:radPr');hide=OxmlElement('m:degHide');hide.set(qn('m:val'),'1');pr.append(hide);rad.append(pr);rad.append(OxmlElement('m:deg'));ee=OxmlElement('m:e')
        ee.append(frac([mr('('),sub('n','AML'),mr(' − 1)'),sub('s','AML'),mr('² + ('),sub('n','ALL'),mr(' − 1)'),sub('s','ALL'),mr('²')],[sub('n','AML'),mr(' + '),sub('n','ALL'),mr(' − 2')]))
        rad.append(ee);nodes=[sub('s','p'),mr(' = '),rad]
    elif text.startswith('d＝'):
        nodes=[mr('d = '),frac([mr('均值',True),sub('x̄','AML'),mr(' − '),sub('x̄','ALL')],[sub('s','p')])]
        # x-bar already denotes the sample mean; omit duplicated Chinese label.
        nodes=[mr('d = '),frac([sub('x̄','AML'),mr(' − '),sub('x̄','ALL')],[sub('s','p')])]
    else:nodes=[mr(text)]
    for node in nodes:math.append(node)
    p._p.append(math)
    return p

def caption(text):
    p=para(text,'Caption'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next=True
    return p

def table(headers,data,widths,size=10.5,compact=False):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False
    pr=t._tbl.tblPr
    for old in pr.findall(qn('w:tblBorders')):pr.remove(old)
    borders=OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        el=OxmlElement('w:'+side); el.set(qn('w:val'),'single' if side in ['top','bottom'] else 'nil')
        if side in ['top','bottom']:el.set(qn('w:sz'),'8'); el.set(qn('w:color'),'000000')
        borders.append(el)
    pr.append(borders)
    for col,w in zip(t.columns,widths):col.width=Cm(w)
    for ridx,values in enumerate([headers]+data):
        row=t.rows[0] if ridx==0 else t.add_row()
        trpr=row._tr.get_or_add_trPr(); trpr.append(OxmlElement('w:cantSplit'))
        if ridx==0:trpr.append(OxmlElement('w:tblHeader'))
        for j,(cell,text,w) in enumerate(zip(row.cells,values,widths)):
            cell.width=Cm(w); cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cp=cell._tc.get_or_add_tcPr(); margins=OxmlElement('w:tcMar')
            for side in ['top','bottom','start','end']:
                el=OxmlElement('w:'+side); el.set(qn('w:w'),('40' if compact else '80') if side in ['top','bottom'] else '70'); el.set(qn('w:type'),'dxa'); margins.append(el)
            cp.append(margins)
            shd=cp.find(qn('w:shd'))
            if shd is not None:cp.remove(shd)
            if ridx==0:
                bd=OxmlElement('w:tcBorders'); bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'4'); bot.set(qn('w:color'),'000000'); bd.append(bot); cp.append(bd)
            p=cell.paragraphs[0]; p.paragraph_format.first_line_indent=Cm(0); p.paragraph_format.line_spacing=1.05 if compact else 1.15
            p.paragraph_format.space_after=Pt(0); p.paragraph_format.space_before=Pt(0)
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT if j==0 else WD_ALIGN_PARAGRAPH.CENTER
            font(p.add_run(str(text)),size,ridx==0)
    p=doc.add_paragraph(); p.paragraph_format.first_line_indent=Cm(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.space_before=Pt(0)
    font(p.add_run(''),3)
    return t

# Page 1
p=para('Golub白血病表达数据','Title'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p=para('多重检验校正报告','Title'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
heading('1 摘要与研究问题')
para('在 Golub 白血病表达数据的 38 份发现样本中，对 7,129 个探针分别进行双侧 Welch t 检验。原始 p≤0.05 的探针有 1,694 个；按完整家族校正后，Bonferroni 显著 26 个，BH 显著 535 个。三种数量对应不同判断规则，不能把数量之差直接当作真实假阳性数。')
para('本报告研究 AML 与 ALL 两种白血病亚型之间的平均表达差异。按课堂规则从发现部分锁定 10 个候选后，7 个在历史留出样本中同时满足候选家族 Bonferroni≤0.05 和方向一致。结果提供进一步验证的探索性证据，尚不构成疾病机制、因果关系或临床应用的证明。')
heading('2 数据与分析协议')
caption('表1 发现与留出样本的构成')
table(['数据部分','ALL','AML','合计','用途'],[['train',27,11,38,'发现与锁定候选'],['test',20,14,34,'检查锁定候选']],[2.7,1.4,1.4,1.4,8.33])
para('CSV 共 72 行、7,132 列。sample_id、original_split、diagnosis 为说明字段，其余 7,129 列构成发现检验家族，包含控制探针。探针是测量特征，不等同于独立基因；数据没有健康对照。')
para('本次核查未发现缺失、非有限值、重复编号、完全重复的表达行或常数探针。表达范围为 −28,400 至 71,369。保留原处理尺度，不把负值视为缺失，不直接取对数。现有字段不足以全面核验患者独立性、批次和来源中心混杂。')
para('对探针 j，H₀ 为 μAML,j＝μALL,j，H₁ 为两者不等；均值差统一定义为 AML−ALL。发现部分只使用 train，逐项采用允许方差不同的 Welch t 检验。阈值取 0.05；Bonferroni 与 BH 均在完整 7,129 项上计算。规则来自课堂协议，本次是已知课堂结果后的复现，不声称完成预注册或首次盲法验证。')

# Page 2
heading('3 p值与多重检验的含义',new=True)
heading('3.1 p小于等于0.05表示什么',2)
para('p 值是：在原假设与检验模型成立的前提下，得到当前这样或更极端检验统计量的概率。此处原假设是两组总体平均表达相同，而不是两组每个人的数值都相同。双侧检验同时考虑 AML 高于或低于 ALL 的情况。')
para('例如 p＝0.03，表示假定总体均值相等并采用当前模型时，当前或更极端的统计量出现概率为 3%。按单项 0.05 规则应拒绝原假设，可以写“观察到反对均值相等的统计证据”。不能写“原假设只有 3% 的概率为真”，也不能写“该探针有 97% 的概率是真实差异”。')
para('一个探针原始 p≤0.05，称为未经多重校正的单项显著。研究同时检验 7,129 项时，这一标签尚未控制整组搜索带来的错误风险。p>0.05 则是证据不足以按该规则拒绝原假设，不能据此证明总体完全相同。')
heading('3.2 为什么1000次检验可能约有50个显著',2)
para('假设 1,000 个原假设全都真实成立，且每个检验的假阳性概率都是 0.05，则平均假发现数 E[V]＝1,000×0.05＝50。这是重复实验的平均数，不要求某次恰好出现 50 个；期望的线性性不要求各检验独立。若改为 5,000 项，相同条件下期望为 250。')
para('因此，假设案例中“1,000 个基因有 53 个原始显著”不足以宣布发现 53 个疾病相关基因。应补充完整家族校正、效应、数据质量和验证证据；也不能反过来断言这 53 个全是假阳性。')
heading('3.3 区分错误个数与错误率',2)
para('设 R 为宣布的全部发现，V 为其中假发现，S 为真发现，故 R＝V＋S。某一次实验的 FDP＝V/max(R,1)；没有发现时记为 0。FWER＝P(V>0)，关注整个家族是否至少错一次；FDR＝E[FDP]，关注长期平均错误比例。')
para('例如一次模拟宣布 20 项，其中 2 项是假的，该次 FDP＝10%，且发生了“至少错一次”。单次记录不能直接给出长期 FDR。重复模拟应先计算每次 FDP，再取平均，通常不能用平均 V 除以平均 R 替代。')
para('若真信号共有 100 个而检出 30 个，检出力为 30%。全零实验没有真信号，检出力记 N/A。真实 Golub 数据没有已知真值，所以本报告只能记录发现数，不能直接计算真实 V、FDP 或真实检出力。')

# Page 3
heading('4 两种校正方法与手算例子',new=True)
heading('4.1 Bonferroni控制家族错误率',2)
formula('pBonf＝min(m×p, 1)；拒绝条件为 pBonf≤α')
para('min 表示取较小值：在 m×p 与 1 之间取较小者。这里的 1 是校正后 p 的上限。例如 m×p＝2.4 时记为1；m×p＝0.006 时保留0.006。')
para('等价地，可以保持原始 p 不变，与 α/m 比较。m＝7,129、α＝0.05 时，原始 p 的界限为 7.0136064×10⁻⁶。原始 p 与校正后 p 是两种表示同一决策的方式，不应拿校正后 p 再去比较 0.05/m，否则相当于重复校正。')
para('Bonferroni 在单项 p 值有效时控制 FWER，不要求检验彼此独立。它重视避免整组中出现任何假发现，代价可能是漏掉较弱的真实差异。')
heading('4.2 BH控制错误发现比例的期望',2)
para('把 p 从小到大排为 p₍₁₎,…,p₍ₘ₎，找到满足 p₍k₎≤k×q/m 的最大 k，拒绝排序前 k 项。本报告 q＝0.05。不能把“每一项各自比较本项排序界限”当作完整算法。')
formula('pBH,(i)＝min{1, min[j≥i](m×p(j)/j)}')
para('上式先对每个排序位置 j 算 m×p₍j₎/j，再从第 i 项到最后一项取最小值，超过 1 则记为 1，最后映回原探针顺序。BH 校正后 p≤0.05 与上述逐步规则给出相同的发现集合。独立或适当正依赖等条件下，BH 提供 FDR 控制；不能对任意探针相关结构无条件承诺。')
caption('表2 教学p值的手算结果  家族大小为5')
table(['排序i','原始p','Bonferroni p','BH排序界限','BH校正p'],[[1,'.001','.005','.010','.005'],[2,'.009','.045','.020','.0225'],[3,'.020','.100','.030','.0333'],[4,'.040','.200','.040','.050'],[5,'.200','1.000','.050','.200']],[1.4,2.4,3.8,3.8,3.83])
para('Bonferroni 发现前 2 项；BH 的最大满足位置为 k＝4，发现前 4 项。这说明控制目标和原始 p 的判定规则不同。两种校正后 p 都与数值 0.05 比较，并不意味着两种方法控制同一种错误。')
para('本报告所称 BH 校正后 p 不自动等于 Storey q-value。Storey 方法还涉及估计真实零假设比例 π₀；本次未实施这一估计。二者都不能解读为“某个发现为假的概率”。')

# Page 4
heading('5 真实数据的发现与效应',new=True)
caption('表3 同一完整家族下的发现数')
table(['判断规则','发现数','控制目标'],[['原始 p≤0.05','1,694','仅单项显著性'],['Bonferroni 校正p≤0.05','26','FWER'],['BH 校正p≤0.05','535','FDR']],[7.0,2.1,6.13])
para('BH 的最大拒绝位置为 535，该位置原始 p＝0.0037466415，小于排序界限 535×0.05/7,129＝0.0037522794。535 是当前数据和方法得到的发现数，不能将 1,694−535 当成已知假阳性，也不能认为 535 项均已得到生物学确认。')
heading('5.1 首位候选的完整证据',2)
para('选择按原始 p 排序第 1 的 U22376_cds2_s_at 作为例子。该编号是探针名称，以下数值均来自发现部分。')
caption('表4 U22376_cds2_s_at的描述统计与推断结果')
table(['指标','结果'],[['AML 样本量；均值；标准差','11；702.36；360.34'],['ALL 样本量；均值；标准差','27；3,863.30；1,999.78'],['均值差 AML−ALL','−3,160.93'],['逐项95% Welch置信区间','[−3,977.85, −2,344.02]'],['合并组内标准差','1,710.07'],['描述性 Cohen’s d','−1.8484'],['原始p','8.3887×10⁻⁹'],['Bonferroni 校正p','5.9803×10⁻⁵'],['BH 校正p','5.9803×10⁻⁵']],[7.4,7.83])
para('负均值差说明 AML 的平均测量值更低。区间在本模型下支持负向均值差，但它是未经同时覆盖和筛选后修正的逐项区间；不能当作对 7,129 项的联合保证。表达测量的上游单位不完整，均值差不能称为表达倍数变化或 log fold change。')

# Page 5
heading('5.2 Cohen d的计算和解释',new=True)
para('均值差保留原始测量尺度。为描述差异相对于组内波动有多大，可以除以合并组内标准差 sₚ，得到无量纲的 Cohen’s d。本报告采用两独立组的描述性合并标准差版本。')
formula('sₚ＝√{[(nAML−1)sAML²＋(nALL−1)sALL²]/(nAML＋nALL−2)}')
formula('d＝(均值AML−均值ALL)/sₚ')
para('本例 sₚ＝√[(10×360.3391²＋26×1,999.7849²)/36]≈1,710.0688；因此 d＝−3,160.9327/1,710.0688≈−1.8484。其含义是 AML 均值比 ALL 低约 1.85 个合并组内标准差。负号只标记所定义比较方向；调换组别，d 的符号也会反转。')
para('p 值回答在原假设下当前证据有多极端，d 描述标准化差异大小，两者不能相互替代。d 不是变化比例、表达倍数或真实性概率，也不能仅凭其绝对值判断临床价值。两组标准差差别较大，因此该 d 仅作描述；并不意味着 Welch 检验额外要求方差相等。')
heading('5.3 为什么本例两个校正后p恰好相同',2)
para('本例原始 p 在完整家族中排名第一。Bonferroni 的计算直接为：')
formula('7,129×8.3887146203×10⁻⁹＝5.9803146528×10⁻⁵')
para('BH 不是只把这个 p 乘以 7,129。它要检查全部排序位置的 m×p₍j₎/j。本次逐项核查发现，这一调整量在 j＝1 时达到全组最小值 5.9803146528×10⁻⁵，因此第 1 项的 BH 校正结果恰好等于 Bonferroni 结果。第 2 项对应调整量为 0.0003500928，也大于该最小值。')
para('所以，截图中的约 0.0000598 是计算结果，不是两种方法的原始 p 阈值。相等只说明本探针本次落在相同调整量上，不代表两种方法等价。即使排名第 1，若后面出现更小的 m×p₍j₎/j，第 1 项的 BH 校正后 p 也会更小。')
para('例如表2第 2 项原始 p＝0.009，Bonferroni 校正后 p＝0.045，而 BH 为 0.0225。这个对照显示，同样用 0.05 判断校正后 p 时，两种方法仍可得到不同调整值和不同发现集合。')

# Page 6
heading('6 已知真值的模拟比较',new=True)
para('每次生成 1,000 个独立正态变量，模拟组大小为 11 与 27，噪声标准差为 1，随机种子 42，各设置重复 1,000 次。模拟已知正态且等方差，使用双侧等方差 t 检验；真实数据采用 Welch 检验。三种规则在同一批原始 p 上比较。')
caption('表5 实验A 全部没有真信号时的发现与波动')
table(['方法','首次R','平均R','R标准差','R的5%至95%分位','FWER及FDR'],[['raw','42','49.944','6.762','[39, 61]','1.000'],['Bonferroni','0','0.051','0.220','[0, 1]','0.051'],['BH','0','0.052','0.227','[0, 1]','0.051']],[2.5,1.2,1.7,2.0,4.2,3.63],10)
para('实验A所有发现都是假的。有发现时 FDP＝1，无发现时为 0，所以此情境 FDR＝FWER，检出力记 N/A。校正方法首次没有发现不等于永不出错；平均 R 也不是至少错一次的概率。0.051 是有限重复的估计，其 Monte Carlo 标准误约 0.007，不能据此断言理论控制失效。')
caption('表6 实验B 有100个真信号时的方法比较')
data=[]
for r in sims:
    if int(r['signals'])==0:continue
    data.append([r['delta'],r['method'],f"{float(r['mean_R']):.3f}",f"{float(r['mean_V']):.3f}",f"{float(r['FWER']):.3f}",f"{float(r['FDR']):.4f}",f"{float(r['power']):.4f}"])
table(['δ','方法','平均R','平均V','FWER','FDR','检出力'],data,[.7,3.0,2.0,2.0,1.9,1.9,1.73],10)
para('δ＝1.0 时，BH 的 FWER＝0.754，但 FDR≈0.0461，两者可同时成立：许多重复实验会至少错一次，平均错误比例仍可能较低。Bonferroni 检出力约 5.96%，BH 约 28.65%，反映控制目标与漏检的权衡。')
para('信号减弱为 δ＝0.5 后，BH 检出力降为约 0.335%，Bonferroni 约 0.226%。模拟中真信号仍有 100 个，因此“没有显著”不能说明“没有效应”。不同效应设置复用相同噪声，便于比较；随机种子是复现起点，不是样本量或重复次数。')

# Page 7
heading('7 选择偏倚与历史留出验证',new=True)
heading('7.1 为什么不能先选前20项再校正',2)
para('若在 7,129 项中先按原始 p 选最小的 20 项，再把家族写成 m＝20，便忽略了这些项目是从大规模搜索中选出的事实。本例完整家族 Bonferroni 与 BH 均约 5.9803×10⁻⁵，错误前20流程均约 1.6777×10⁻⁷。数值变小并未增加可靠证据。')
para('只有这 20 项确实由外部依据在查看当前数据前指定，且研究只检验该家族，才可能合理使用 m＝20。即使首位探针在正确与错误流程中都显著，仍不能证明错误流程有效。')
heading('7.2 锁定名单和验证结果',2)
para('按课堂规则，在 train 的 BH 显著者中按原始 p、探针名升序锁定前 10 项并保存发现方向，随后才访问 test 的候选数据。验证家族为这 10 项，逐项 Welch 检验；候选家族 Bonferroni≤0.05 且方向相同才记为满足判据。本次按该顺序重新运行，但结果已在课堂材料中公开。')
caption('表7 锁定候选在历史留出样本中的结果')
data=[]
for r in valid:
    data.append([r['probe'],f"{float(r['discovery_delta']):.2f}",f"{float(r['validation_delta']):.2f}",f"{float(r['bonferroni']):.6f}",'是' if r['passes']=='True' else '否'])
table(['锁定探针','发现差值','留出差值','留出Bonferroni p','通过'],data,[4.6,2.5,2.5,4.2,1.43],9.5)
para('10 项方向全部一致，其中 7 项同时通过显著性判据。U22376_cds2_s_at 的留出差值为 −1,766.73，候选家族 Bonferroni p＝0.023186。未通过的 3 项不能据此判为无效，也不得在看到结果后替换名单。')
para('m＝10 仅用于已锁定候选在留出部分的检查，不可反过来用于发现部分。历史留出样本来自同一研究，不能称为新收集的外部独立验证。')

# Page 8
heading('8 方法选择与结论边界',new=True)
heading('8.1 根据研究目的选择控制目标',2)
para('探索性组学筛选希望形成可供后续研究的候选集合，可在说明有效检验与依赖条件的前提下采用 BH 控制 FDR，并报告效应、区间及验证计划。发现数较多不是方法更正确的证明，错误发现仍会消耗后续资源。')
para('确认性临床研究若预先指定 5 个主要终点，任何一项误报都可能影响主要主张，可以优先控制这 5 项的 FWER，例如采用 Bonferroni：各原始 p 与 0.05/5＝0.01 比较。实际方案还应结合终点结构与研究设计预先确定，不能看结果后选择最有利的方法。')
para('分析 4 个亚组、尝试 3 个模型后只报告最显著者，也增加选择机会，应在方案中说明相应假设和选择流程。检验家族服务于研究主张，不能靠少报告几个结果消除已经发生的搜索。')
heading('8.2 校正之后仍然缺少什么',2)
para('多重校正不能修复批次混杂、异常值影响、不独立样本或不合适的单项检验。真实表达数据存在小样本和方差差异，且探针相关结构未被全面验证。本次未实施批次调整、稳健性重分析或外部队列检验，因此不报告这些分析的结果。')
para('发现部分按极小 p 筛选候选，效应可能受到选择偏倚影响。表4的逐项区间不提供多重比较或筛选后的覆盖率保证。后续需核查探针注释与质量、在独立数据中复测效应并评估混杂；机制研究和临床评价另需相应设计。')
heading('8.3 可支持的最终结论',2)
para('本次在 38 份发现样本的 7,129 个探针中得到 535 个 BH 显著结果。按照课堂规则锁定 10 项后，其中 7 项在历史留出样本中方向一致并通过候选家族 Bonferroni 判据。这些发现支持进一步研究 AML 与 ALL 的表达差异，但不证明存在 535 个独立疾病基因，也不能证明因果机制或临床有效性。')
heading('参考文献',2)
for text in [
'[1] Benjamini Y, Hochberg Y. Controlling the false discovery rate: a practical and powerful approach to multiple testing. JRSS B. 1995;57:289–300. DOI: 10.1111/j.2517-6161.1995.tb02031.x.',
'[2] Storey JD. A direct approach to false discovery rates. JRSS B. 2002;64:479–498. DOI: 10.1111/1467-9868.00346.',
'[3] Golub TR, et al. Molecular classification of cancer. Science. 1999;286:531–537. DOI: 10.1126/science.286.5439.531.',
'[4] 本课工作单、27页PPT、数据说明及配套Notebook。数据来源为 OpenML leukemia 1104 v1；本报告使用课程本地CSV快照。',
'[5] SciPy 文档 false_discovery_control。https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.false_discovery_control.html']:
    p=para(text,indent=False,size=9);p.paragraph_format.space_after=Pt(3);p.alignment=WD_ALIGN_PARAGRAPH.LEFT

# Page 9
heading('附录 复现代码与阅读方法',new=True)
heading('A 运行入口和输出',2)
para('配套“复现分析.py”以脚本所在位置定位本项目数据，不修改原始 CSV。将脚本保留在“作业报告”目录，在项目根目录运行下列命令；依赖 NumPy、pandas 和 SciPy。')
p=para('python 作业报告/复现分析.py',indent=False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
para('结果写入 .tmp/multiple-test/student-results/，包括全部发现结果、锁定名单、留出检查、模拟汇总，以及保存数据哈希、环境和关键数值的 analysis_summary.json。')
heading('B 逐段读懂主要代码',2)
table(['代码','在本报告中做什么'],[['pd.read_csv(...)','把CSV读成行列数据表'],['df.original_split.eq("train")','只选择发现部分的38份样本'],['train.diagnosis.eq("AML")','按诊断选择AML组；ALL同理'],['stats.ttest_ind(a,b,axis=0,\nequal_var=False)','逐列进行Welch检验；axis=0表示按列比较'],['np.minimum(m*p,1)','将原始p乘完整家族大小，最大记1'],['stats.false_discovery_control(p,\nmethod="bh")','对完整p数组做BH校正'],['delta / pooled_sd','均值差除合并组内标准差，得到d'],['(adjusted_p <= .05).sum()','统计有多少项通过指定规则']],[7.0,8.23],10,compact=True)
para('a、b 分别是 AML、ALL 的样本×探针矩阵；axis=0 沿样本方向计算，每个探针得到一个结果。sum() 把满足条件的 True 当作1相加，得到发现数。m 必须是完整家族大小。')
heading('C 数值复核和环境记录',2)
para('本次使用手写 Welch 统计量、自由度和双侧尾概率复核 SciPy p 值，并使用排序与反向累计最小值复核 SciPy BH 结果，逐项一致。发现数为 1,694、26、535；完整复跑模拟与留出检查亦与本报告一致。')
para('分析环境：'+ '；'.join(k+' '+v for k,v in s['versions'].items())+'。文档排版使用单独的文档运行环境。',indent=False,size=9.5)
para('数据 SHA-256：'+s['data_sha256'],indent=False,size=9)
para('这份哈希用于确认输入文件相同；种子42用于复现随机模拟。实际软件版本或数据改变时，应重新核对结果，不应只修改表格中的数字。')

# Consistent editable page number and non-proofed technical tokens.
for sec in doc.sections:
    p=sec.footer.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent=Cm(0)
    field=OxmlElement('w:fldSimple'); field.set(qn('w:instr'),'PAGE')
    r=OxmlElement('w:r'); tx=OxmlElement('w:t'); tx.text='1'; r.append(tx);field.append(r);p._p.append(field)
for el in doc.element.xpath('//w:r'):
    rp=el.find(qn('w:rPr'))
    if rp is None:rp=OxmlElement('w:rPr');el.insert(0,rp)
    if rp.find(qn('w:noProof')) is None:rp.append(OxmlElement('w:noProof'))
doc.save(DOC)
assert hashlib.sha256(REF.read_bytes()).hexdigest()==ref_hash
(TMP/'reference_evidence.json').write_text(json.dumps({'reference':str(REF),'sha256':ref_hash,'new_document':str(DOC),'table_count':len(doc.tables),'section_count':len(doc.sections)},ensure_ascii=False,indent=2),encoding='utf-8')
source=(TMP/'rebuild_analysis.py').read_text(encoding='utf-8')
source=source.split('with ZipFile(ROOT/')[0]
source=source.replace('Path(__file__).resolve().parents[2]','Path(__file__).resolve().parent.parent')
source=source.replace('from zipfile import ZipFile\n','').replace('import xml.etree.ElementTree as etree\n','')
(OUT/'复现分析.py').write_text(source,encoding='utf-8')
(OUT/'运行说明.txt').write_text('多重检验报告配套复现\n\n在项目根目录运行：\npython 作业报告/复现分析.py\n\n依赖：numpy、pandas、scipy。\n输入为 data/leukemia.csv；不修改原始数据。\n结果位于 .tmp/multiple-test/student-results/。\n完整结果和环境记录已纳入压缩包。\n\n本次为公开课堂结果后的复现，不构成首次盲法验证。\n',encoding='utf-8')
with zipfile.ZipFile(OUT/'多重检验校正_代码与结果.zip','w',zipfile.ZIP_DEFLATED) as z:
    for name in ['复现分析.py','运行说明.txt']:z.write(OUT/name,name)
    for f in RESULT.iterdir():
        if f.suffix in ['.csv','.json']:z.write(f,'student-results/'+f.name)
print(DOC)
print('Tables:',len(doc.tables),'Paragraphs:',len(doc.paragraphs))
