from pathlib import Path
import sys,json,hashlib,platform
import numpy as np
import pandas as pd
import sklearn, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from docx import Document
from docx.shared import Cm,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parent
BASE=ROOT/'7. Naive-Bayes'
OUT=ROOT/'报告成果'; OUT.mkdir(exist_ok=True)
FIG=OUT/'figures'; FIG.mkdir(exist_ok=True)
sys.path.insert(0,str(BASE/'scripts'))
from lab_utils import FEATURES,fit_model,scores,threshold_table,choose_threshold,hand_calculation,reliability,wilson
train=pd.read_csv(BASE/'data/student/train.csv'); val=pd.read_csv(BASE/'data/student/validation.csv')
target='bladder_inflammation'
model=fit_model(train)
pv=model.predict_proba(val[FEATURES])[:,1]
tt=threshold_table(val[target],pv); threshold=choose_threshold(tt)
assert threshold==.5
# 阈值仅由验证集确定，随后才读取现有测试标签复算；不声称完成过教师释放或前瞻性盲评。
test=pd.read_csv(BASE/'data/student/locked-test-features.csv').merge(pd.read_csv(BASE/'data/teacher/locked-test-labels.csv'),on='sample_id',validate='one_to_one')
pt=model.predict_proba(test[FEATURES])[:,1]
pt_train=model.predict_proba(train[FEATURES])[:,1]
metrics={k:scores(d[target],p,threshold) for k,d,p in [('train',train,pt_train),('validation',val,pv),('test',test,pt)]}
prior=scores(val[target],np.full(len(val),train[target].mean()))
hand,weights,posterior=hand_calculation(train,[0,0,0,1,1,1])
# 第二条独立计算路径：直接由计数计算朴素贝叶斯权重，与 sklearn 逐条互核。
def manual(x):
    w=[]
    for c in [0,1]:
        part=train[train[target]==c]; theta=(part[FEATURES].sum().to_numpy()+1)/(len(part)+2)
        w.append(len(part)/len(train)*np.prod(np.where(x==1,theta,1-theta),axis=1))
    return w[1]/(w[0]+w[1])
assert np.allclose(manual(val[FEATURES].to_numpy()),pv,rtol=1e-12,atol=1e-12)
assert np.allclose(manual(test[FEATURES].to_numpy()),pt,rtol=1e-12,atol=1e-12)
duplicates=[]; prob_dup=[]
for n in [0,1,5]:
    cols=FEATURES+['lumbar_pain']*n
    from sklearn.naive_bayes import BernoulliNB
    m=BernoulliNB(alpha=1,binarize=None).fit(train[cols].to_numpy(),train[target])
    p=m.predict_proba(val[cols].to_numpy())[:,1]; prob_dup.append(p)
    duplicates.append({'extra_copies':n,**scores(val[target],p)})
dup=pd.DataFrame(duplicates); rel=reliability(val[target],pv)
names=['发热','恶心','腰痛','尿急','排尿痛','尿道灼痛']
plt.rcParams.update({'font.sans-serif':['Microsoft YaHei','SimHei'],'axes.unicode_minus':False,'font.size':10,'savefig.dpi':200})
blue='#285879'; orange='#bd6b36'
def save(name):
    plt.tight_layout(); plt.savefig(FIG/name,bbox_inches='tight',facecolor='white');plt.close()
fig,ax=plt.subplots(figsize=(8,2.7));ax.axis('off')
for x,text in zip([.04,.37,.70],['训练集 72 条\n估计先验与症状概率\n固定 α = 1','验证集 24 条\n比较预定阈值\n选择 t = 0.5','测试集 24 条\n沿用原训练模型\n一次性复算留出表现']):
    ax.add_patch(Rectangle((x,.22),.26,.65,fc='#edf2f5',ec=blue,lw=1.4,transform=ax.transAxes)); ax.text(x+.13,.55,text,ha='center',va='center',transform=ax.transAxes)
for x in [.30,.63]:ax.annotate('',xy=(x+.05,.55),xytext=(x,.55),xycoords='axes fraction',arrowprops={'arrowstyle':'->','color':blue,'lw':2})
save('01_workflow.png')
fig,ax=plt.subplots(figsize=(7.5,3.3));x=np.arange(6);a=np.exp(model.feature_log_prob_)
ax.bar(x-.18,a[0],.36,label='阴性类',color=blue);ax.bar(x+.18,a[1],.36,label='阳性类',color=orange);ax.set_xticks(x,names);ax.set_ylim(0,1.15);ax.set_ylabel('平滑后症状出现概率');ax.legend(ncol=2);save('02_parameters.png')
fig,axes=plt.subplots(1,2,figsize=(7.5,3))
for ax,key,title in zip(axes,['validation','test'],['验证集 n = 24','测试集 n = 24']):
    s=metrics[key]; arr=np.array([[s['tn'],s['fp']],[s['fn'],s['tp']]])
    ax.imshow(arr,cmap='Blues',vmin=0,vmax=12)
    for (i,j),v in np.ndenumerate(arr):ax.text(j,i,str(v),ha='center',va='center',color='white' if v>6 else 'black',fontsize=18)
    ax.set_xticks([0,1],['预测阴性','预测阳性']);ax.set_yticks([0,1],['实际阴性','实际阳性']);ax.set_title(title)
save('03_confusion.png')
fig,ax=plt.subplots(figsize=(7.5,3));ax.plot(tt.threshold,tt.sensitivity,'o-',color=orange,label='灵敏度');ax.plot(tt.threshold,tt.specificity,'s-',color=blue,label='特异度');ax.axvline(threshold,ls='--',c='gray');ax.set(xlabel='决策阈值',ylabel='指标值',ylim=(-.05,1.08));ax.legend();save('04_threshold.png')
fig,axes=plt.subplots(1,2,figsize=(7.5,3));axes[0].scatter(pv,prob_dup[2],c=val[target],cmap='coolwarm',s=40,alpha=.8);axes[0].plot([0,1],[0,1],'--',c='gray');axes[0].set(xlabel='原模型阳性概率',ylabel='额外复制五次后的概率',xlim=(0,1),ylim=(0,1));axes[1].bar(['0 次','1 次','5 次'],dup.log_loss,color=[blue,'#7296ad',orange]);axes[1].set(xlabel='额外复制腰痛列次数',ylabel='验证集平均 log loss');save('05_duplicate.png')
fig,ax=plt.subplots(figsize=(7,3));r=rel[rel.n>0];ax.plot([0,1],[0,1],'--',c='gray',label='理想一致线');ax.plot(r.mean_probability,r.observed_fraction,'o-',c=blue,label='验证集分箱')
for _,row in r.iterrows():ax.annotate('n='+str(int(row.n)),(row.mean_probability,row.observed_fraction),xytext=(6,7),textcoords='offset points')
ax.set(xlabel='箱内平均预测阳性概率',ylabel='箱内实际阳性比例',xlim=(-.04,1.04),ylim=(-.06,1.15));ax.legend();save('06_reliability.png')
pd.DataFrame(metrics).T.to_csv(OUT/'metrics.csv');tt.to_csv(OUT/'thresholds.csv',index=False);dup.to_csv(OUT/'duplicate_audit.csv',index=False);hand.to_csv(OUT/'hand_calculation.csv',index=False)
summary={'python':platform.python_version(),'numpy':np.__version__,'pandas':pd.__version__,'sklearn':sklearn.__version__,'matplotlib':matplotlib.__version__,'threshold':threshold,'metrics':metrics,'manual_max_error':float(np.max(np.abs(manual(val[FEATURES].to_numpy())-pv))),'hashes':{str(p.relative_to(BASE)):hashlib.sha256(p.read_bytes()).hexdigest() for p in [BASE/'data/student/train.csv',BASE/'data/student/validation.csv',BASE/'data/student/locked-test-features.csv',BASE/'data/teacher/locked-test-labels.csv']}}
(OUT/'reproducibility.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')

ref=ROOT.parent/'5. Hypothesis test/统计检验与假设设计报告-吴际恒.docx'
doc=Document(ref)
for child in list(doc._element.body):
    if child.tag!=qn('w:sectPr'):doc._element.body.remove(child)
for sec in doc.sections:
    for container in [sec.header,sec.footer]:
        for p in container.paragraphs:p.text=''
doc.core_properties.title='基于症状数据的朴素贝叶斯分类与评价';doc.core_properties.author='吴际恒'
md=[]
def font(r,size=11,bold=False,east='宋体'):
    r.font.name='Times New Roman';r.font.size=Pt(size);r.bold=bold;r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),east)
def p(text):
    a=doc.add_paragraph();a.paragraph_format.first_line_indent=Cm(.74);a.paragraph_format.line_spacing=1.25;a.paragraph_format.space_after=Pt(5);font(a.add_run(text));md.append(text+'\n')
def h(text):
    a=doc.add_paragraph();a.paragraph_format.keep_with_next=False;a.paragraph_format.space_before=Pt(8);font(a.add_run(text),15,True,'黑体');md.append('## '+text+'\n')
def table(headers,rows):
    t=doc.add_table(rows=1,cols=len(headers));t.autofit=False
    width=doc.sections[0].page_width-doc.sections[0].left_margin-doc.sections[0].right_margin
    for col in t.columns:col.width=int(width/len(headers))
    borders=OxmlElement('w:tblBorders')
    for key in ['top','bottom','left','right','insideH','insideV']:
        b=OxmlElement('w:'+key);b.set(qn('w:val'),'single' if key in ['top','bottom'] else 'nil');b.set(qn('w:sz'),'8');borders.append(b)
    t._tbl.tblPr.append(borders)
    for i,row in enumerate([headers]+rows):
        cells=t.rows[0].cells if i==0 else t.add_row().cells
        tr=cells[0]._tc.getparent().get_or_add_trPr();tr.append(OxmlElement('w:cantSplit'))
        if i==0:tr.append(OxmlElement('w:tblHeader'))
        for cell,text in zip(cells,row):
            a=cell.paragraphs[0];a.alignment=1;a.paragraph_format.space_after=Pt(3);font(a.add_run(str(text)),9,i==0)
    md.append('| '+' | '.join(headers)+' |\n| '+' | '.join(['---']*len(headers))+' |\n'+'\n'.join('| '+' | '.join(map(str,row))+' |' for row in rows)+'\n')
def figure(name,caption,width=14):
    a=doc.add_paragraph(style='Normal');a.alignment=1;a.paragraph_format.keep_with_next=False
    # The retained reference document contains list formatting in its default
    # paragraph inheritance; make figure paragraphs explicitly ordinary so no
    # stray bullet appears beside an image in Word.
    ppr=a._p.get_or_add_pPr()
    for child in list(ppr):
        if child.tag in (qn('w:numPr'), qn('w:pStyle')): ppr.remove(child)
    st=OxmlElement('w:pStyle');st.set(qn('w:val'),'Normal');ppr.insert(0,st)
    a.add_run().add_picture(str(FIG/name),width=Cm(width))
    a=doc.add_paragraph();a.alignment=1;font(a.add_run(caption),9);md.append('!['+caption+'](figures/'+name+')\n')
def page():doc.add_page_break()
title=doc.add_paragraph(style='Title');title.alignment=1;font(title.add_run('基于症状数据的\n朴素贝叶斯分类与评价'),22,True,'黑体');md.append('# 基于症状数据的朴素贝叶斯分类与评价\n')
a=doc.add_paragraph();a.alignment=1;font(a.add_run('吴际恒    2026 年 11 月 13 日'),11)
h('摘要')
p('本研究使用 UCI Acute Inflammations 数据中的 120 条专家构造记录，以六项二元症状预测膀胱炎标签。保持 72 条训练、24 条验证和 24 条测试的固定划分，采用 Laplace 平滑系数为 1 的 Bernoulli 朴素贝叶斯模型。参数仅由训练集估计，阈值依据验证集的灵敏度与特异度确定。选定阈值为 0.5，验证集和测试集均得到 24/24 的正确分类，平均对数损失分别为 0.087214 和 0.074492。重复输入实验表明，未增加新信息的症状副本仍可改变概率和类别。由于全部数据仅有 9 种症状组合，且测试组合均在训练中出现，结果支持同源数据中的规则再现，尚不足以证明真实临床人群中的泛化能力。')
p('关键词：朴素贝叶斯；条件独立；分类评价；对数损失；数据划分')
h('一 研究问题与分析对象')
p('分析问题是：如何由已知症状与标签估计预测关系，以及如何判断这一关系在保留记录上是否有效。一条样本是数据表中的一条专家构造记录；特征是记录中可用于预测的症状；标签是是否存在膀胱炎。记录不等同于独立患者，预测概率也不能直接解释为已获临床验证的患病风险。')
page();h('二 数据处理与评价设计')
p('六项输入依次为体温不低于 38°C、恶心、腰痛、尿急、排尿痛及尿道灼痛，出现编码为 1，未出现为 0。原始连续温度只用于生成发热指标；样本 ID 用于关联与审计。ID、原始温度和另一疾病标签 nephritis 均不进入模型，以避免改变既定特征定义或引入标签信息。')
table(['数据部分','记录数','阴性','阳性','用途'],[['训练集',72,37,35,'估计参数'],['验证集',24,12,12,'评价与选阈值'],['测试集',24,12,12,'留出复算']])
figure('01_workflow.png','图 1 参数估计 阈值选择与留出评价的数据流程')
profiles=lambda d:set(map(tuple,d[FEATURES].to_numpy()))
seen_v=sum(tuple(x) in profiles(train) for x in val[FEATURES].to_numpy());seen_t=sum(tuple(x) in profiles(train) for x in test[FEATURES].to_numpy())
p(f'三个集合的 ID 两两不重叠。逐列核查显示，六输入与主标签不存在缺失，输入取值均为 0 或 1。验证集中 {seen_v}/24 条、测试集中 {seen_t}/24 条记录的六输入组合在训练集中出现过。新的一行记录与未见过的症状组合并非同一概念，ID 不重叠不能保证输入模式独立。')
p('原始数据共有 120 条记录；六输入二元化后仅有 9 种组合。按除 ID 外全部保留字段比较，有 21 条额外完全重复记录。分析保留固定划分而不重新去重，因此评价范围限定为该构造数据的同源留出表现。')
p('本报告按既定协议复算现有数据，阈值不使用测试标签选择；并未据此声称已经完成教师释放程序、预注册或前瞻性盲评。测试后不更换输入、调整平滑系数或合并验证集重训。')
page();h('三 从训练计数到后验概率')
p('数据模型可写为 p=fθ(x)，其中 x 为六项症状，θ 为由训练数据估计的参数，p 为阳性概率。训练改变参数，预测只将新输入代入固定模型。以单症状教学数据为例，无排尿痛的 3 条中有 1 条阳性，有排尿痛的 5 条中有 4 条阳性，故经验概率分别为 1/3 和 4/5。增加一条有排尿痛但阴性的记录并重新训练，后者变为 4/6；只调用预测则不会更新它。')
p('同一教学例子中，P(Y=1)=5/8，P(x=1|Y=1)=4/5，P(x=1|Y=0)=1/3。Bayes 公式将阳性权重 (5/8)×(4/5)=1/2 与阴性权重 (3/8)×(1/3)=1/8 归一化，得到 P(Y=1|x=1)=0.8。类内症状概率与给定症状后的疾病概率具有不同分母，不能互换。')
p('六个二元特征共有 2⁶=64 种可能组合。朴素贝叶斯以给定类别后特征条件独立为近似，将联合似然拆成六个单特征因子；该假设并非要求总体症状彼此独立。训练中先验为 πc=nc/n，平滑后症状出现概率为 θcj=(kcj+1)/(nc+2)。')
p('类别权重 wc=πc∏j[θcj^xj(1−θcj)^(1−xj)]，阳性后验为 p=w1/(w0+w1)。因此无症状也提供证据，xj=0 时必须使用补概率 1−θcj。平滑避免零计数导致乘积归零，但不能消除特征依赖或数据偏倚。')
figure('02_parameters.png','图 2 仅由训练集估计的类内症状出现概率')
page();h('四 六症状预测的手工核验')
p('固定输入 x=(0,0,0,1,1,1)，表示无发热、无恶心、无腰痛，同时出现尿急、排尿痛与尿道灼痛。训练集阴性 37 条、阳性 35 条，故 π0=37/72、π1=35/72。下表列出用于权重相乘的实际似然项。')
rows=[]
for j,name in enumerate(names):
    r0=hand[(hand.label==0)&(hand.feature==FEATURES[j])].iloc[0];r1=hand[(hand.label==1)&(hand.feature==FEATURES[j])].iloc[0]
    rows.append([name,int(r0.x),int(r0.count_one),int(r1.count_one),f'{r0.likelihood_term:.6f}',f'{r1.likelihood_term:.6f}'])
table(['特征','输入','阴性计数¹','阳性计数¹','阴性似然','阳性似然'],rows)
p('¹ 计数指该类别中症状为 1 的记录数；当当前输入为 0 时，似然项使用平滑后症状概率的补数。')
p(f'将六项似然分别与类别先验相乘，得到 w0={weights[0]:.10f}，w1={weights[1]:.10f}；归一化得到 p={posterior:.10f}。在阈值 0.5 下预测为阳性。该概率是模型在给定参数与条件独立近似下的输出，不是个体诊断的确定性结论。')
p('进一步采用独立的计数乘积路径，对验证集和测试集逐条计算后验，再与 sklearn 的 predict_proba 输出比较，结果在数值精度内一致。这一核验同时检查了先验、无症状项、平滑分母和归一化步骤。')
h('五 概率误差与类别误差')
p('二分类对数损失为 ℓ(y,p)=−[y ln p+(1−y)ln(1−p)]，数据集损失为各记录损失的平均，数值越小越好。对真实标签为 1、0 的两个样本，模型 A 给出 0.6、0.4，模型 B 给出 0.9、0.1，两者 accuracy 均为 1，但平均损失分别为 0.510826 与 0.105361。B 给真实类别分配的概率更高。')
p('当 y=0 时，p=0.6 与 p=0.99 都分类错误，损失分别为 0.916291 与 4.605170。后者几乎排除了真实类别，因此受到更大惩罚。accuracy 只保留跨阈值后的类别，无法表达这种差别。本模型通过训练计数估参，不通过梯度下降最小化验证损失。')
page();h('六 分类指标与验证集阈值选择')
p('混淆矩阵以真实类别为行、预测类别为列。TP 为检出的真实阳性，FN 为漏检阳性，FP 为误报阳性，TN 为正确识别的阴性。灵敏度 TP/(TP+FN) 的分母是真实阳性；精确率 TP/(TP+FP) 的分母是预测阳性，二者不能混淆。')
table(['指标','计算式','含义'],[['准确率','(TP+TN)/n','全部记录中分类正确比例'],['灵敏度','TP/(TP+FN)','真实阳性中检出比例'],['特异度','TN/(TN+FP)','真实阴性中排除比例'],['精确率','TP/(TP+FP)','预测阳性中实际阳性比例'],['F1','2TP/(2TP+FP+FN)','精确率与召回率的调和平均']])
p('以八条独立教学预测为例，在阈值 0.5 下 TN=3、FP=1、FN=2、TP=2，准确率为 5/8，灵敏度为 2/4，特异度为 3/4，精确率为 2/3，F1 为 4/7。将阈值降至 0.3 后，TP=3、FN=1、FP=3、TN=1，灵敏度升至 0.75，特异度降至 0.25。概率与参数未改变，改变的是概率到类别的决策规则。')
table(['阈值','FP','FN','灵敏度','特异度'],[[f'{t:.1f}',int(tt[tt.threshold==t].iloc[0].fp),int(tt[tt.threshold==t].iloc[0].fn),f'{tt[tt.threshold==t].iloc[0].sensitivity:.4f}',f'{tt[tt.threshold==t].iloc[0].specificity:.4f}'] for t in [.1,.5,.9]])
figure('04_threshold.png','图 3 验证集阈值变化与两类识别能力',13)
p('选择规则为先满足灵敏度不低于 0.90，再最大化特异度；并列时优先最接近 0.5 的阈值，仍并列取较小者。网格比较得到 t=0.5。低阈值增加检出倾向，但也可能带来误报；该规则用于本次分析，不代表临床最优决策。')
page();h('七 训练 验证与留出测试结果')
table(['模型与数据','accuracy','log loss','TN / FP / FN / TP'],[['NB 训练',f"{metrics['train']['accuracy']:.4f}",f"{metrics['train']['log_loss']:.6f}",'37 / 0 / 0 / 35'],['NB 验证','1.0000',f"{metrics['validation']['log_loss']:.6f}",'12 / 0 / 0 / 12'],['NB 测试','1.0000',f"{metrics['test']['log_loss']:.6f}",'12 / 0 / 0 / 12'],['先验基线 验证','0.5000',f"{prior['log_loss']:.6f}",'12 / 0 / 12 / 0']])
p('先验基线对所有输入输出训练集阳性比例 35/72≈0.4861；在阈值 0.5 下全部预测阴性，因此灵敏度为 0，特异度为 1，精确率因没有预测阳性而未定义。NB 在验证集上的灵敏度、特异度、精确率和 F1 均为 1，表明症状信息在该数据内部具有辨别作用。')
figure('03_confusion.png','图 4 固定阈值 0.5 下的验证集与测试集混淆矩阵')
p('测试准确率为 24/24，灵敏度与特异度分别为 12/12。验证与测试的类别指标相同，但 log loss 不同，说明相同的类别正确率可以对应不同的概率预测质量。训练损失略高于验证损失并不矛盾：两组输入构成不同，有限样本上的损失没有固定大小关系。')
ci=wilson(24,24);cs=wilson(12,12)
p(f'以二项独立抽样近似计算，24/24 准确率的 Wilson 95% 区间约为 [{ci[0]:.3f}, {ci[1]:.3f}]，12/12 灵敏度或特异度对应区间约为 [{cs[0]:.3f}, {cs[1]:.3f}]。这些区间说明即使全对，小样本仍存在估计不确定性；但专家构造与重复模式不满足临床随机独立抽样语境，不能将区间直接解释为临床总体性能保证。')
page();h('八 条件独立假设与概率可靠性')
table(['额外腰痛副本数','验证准确率','验证 log loss'],[[int(r.extra_copies),f'{r.accuracy:.4f}',f'{r.log_loss:.6f}'] for _,r in dup.iterrows()])
figure('05_duplicate.png','图 5 重复输入引起的验证概率变化与平均损失变化')
idx=int(np.argmax(np.abs(pv-prob_dup[2])))
p(f'复制腰痛列没有增加新的测量，却使同一证据的似然被重复相乘。例如验证记录 {val.iloc[idx].sample_id} 的阳性概率由 {pv[idx]:.6f} 变为 {prob_dup[2][idx]:.6f}。额外复制五次后准确率降至 0.75，log loss 升至约 0.385377；复制一次在本数据中反而降低平均损失。因此，不能推断任何重复都会使所有指标恶化，能够确定的是模型对重复计入证据并不保持不变。')
figure('06_reliability.png','图 6 验证集预测概率分箱与实际阳性比例',12)
p('可靠性图比较同一概率箱内的平均预测概率与实际阳性比例。低概率箱有 12 条，平均概率约 0.089002，实际阳性比例为 0；高概率箱有 11 条，平均概率约 0.950279，实际阳性比例为 1。其余非空箱只有 1 条，不能据此作稳定校准判断；空箱不提供信息。分类全对并不意味着每个概率数值都已得到充分验证。')
page();h('九 讨论与结论')
p('只看 accuracy 会掩盖类别比例与错误类型。例如 1,000 条记录中仅 10 条阳性，全部预测阴性即可获得 99% 准确率，但灵敏度为 0，阳性检出数为 0，精确率未定义。筛查场景通常需要关注漏检，但误报也可能增加检查负担；具体权衡应由使用目的、患病率和后续措施共同决定。')
p('本研究的训练、预测和评价具有不同职责。训练用已知标签统计参数；预测以固定参数将症状映射为概率；验证集用于选阈值；测试集描述未用于估参和选阈值的保留记录表现。若根据测试成绩调整特征或阈值，测试集就参与了开发，原成绩不能继续充当独立最终评价。')
p('从模型结构看，先验表达训练中的类别比例，似然表达给定类别时观察症状的可能性，后验通过 Bayes 归一化得到。条件独立使稀疏组合问题转化为单症状计数问题，也带来重复证据放大的风险。平滑解决零频率问题，log loss 补充类别指标未保留的概率信息，两者不能替代对数据来源与依赖结构的检查。')
p('主要限制来自数据构造和评价范围。120 条记录由专家创建用于专家系统测试，并非真实临床队列；只有 9 种二元输入模式，测试模式均在训练中出现。因此，内部全对更接近已有规则的再现，未检验全新症状模式、不同医院或不同人群的适用性。数据中的预测关联也不能解释为症状对疾病的因果作用。')
p('综合而言，固定的六输入 Bernoulli 朴素贝叶斯在本次同源留出数据中取得完全正确的分类，且优于只输出先验比例的基线。该结果说明完整的估参、预测与评价流程可以在小型数据上清楚复现；其科学含义应限定于所观察的构造记录，临床用途仍需独立、具有代表性的数据验证。')
h('复现说明与资料来源')
p('分析脚本 build_naive_bayes_report.py 以自身路径定位数据，保留原始 CSV，输出指标、阈值表、手算表、图形及 reproducibility.json。该记录包含依赖版本、输入文件 SHA-256 和两种计算路径的核验误差。运行命令为 python build_naive_bayes_report.py。')
p('数据来源：Jacek Czerniak，Acute Inflammations，UCI Machine Learning Repository，DOI 10.24432/C5V59S，CC BY 4.0。数据处理包括温度二元化、0/1 编码与固定划分。算法和指标定义见随附 scripts/lab_utils.py 及 sklearn.naive_bayes.BernoulliNB。')
p('AI 使用说明：使用 AI 辅助分析脚本编写、图表生成和文字组织；关键后验概率由训练计数公式与统计库输出交叉核验。')
doc.save(OUT/'朴素贝叶斯分类实验报告_图框修正版.docx')
(OUT/'朴素贝叶斯分类实验报告.md').write_text('\n'.join(md),encoding='utf-8')
print(json.dumps({'out':str(OUT),'metrics':metrics,'threshold':threshold,'duplicates':duplicates},ensure_ascii=False,indent=2))
