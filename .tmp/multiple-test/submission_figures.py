from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

ROOT=Path(__file__).resolve().parents[2]
OUT=ROOT/'.tmp/multiple-test/submission-figures'
OUT.mkdir(exist_ok=True)
plt.rcParams.update({'font.family':'Microsoft YaHei','axes.unicode_minus':False,'font.size':11,'axes.spines.top':False,'axes.spines.right':False,'savefig.dpi':220})
colors=['#777777','#356C9B','#238A83']
fig,ax=plt.subplots(figsize=(7.2,3.1),layout='constrained')
vals=[1694,26,535]
bars=ax.bar(['未校正','Bonferroni','BH'],vals,color=colors,width=.55)
ax.bar_label(bars,labels=['1,694','26','535'],padding=5,fontsize=12)
ax.set_ylabel('显著探针数');ax.set_ylim(0,1900);ax.grid(axis='y',alpha=.2);ax.set_axisbelow(True)
fig.savefig(OUT/'discovery.png');plt.close(fig)

df=pd.read_csv(ROOT/'data/leukemia.csv')
probe='U22376_cds2_s_at'
fig,axs=plt.subplots(1,2,figsize=(7.2,3.3),layout='constrained',sharey=True)
for ax,split,title in zip(axs,['train','test'],['发现样本（38份）','留出检验样本（34份）']):
    data=df[df.original_split.eq(split)]
    for j,(group,col) in enumerate(zip(['ALL','AML'],['#356C9B','#238A83'])):
        y=data.loc[data.diagnosis.eq(group),probe].to_numpy()
        # Deterministic offset separates individual observations without changing values.
        ax.scatter(j+np.linspace(-.1,.1,len(y)),y,color=col,alpha=.7,s=19)
        ax.hlines(y.mean(),j-.2,j+.2,color='black',linewidth=2)
    ax.set_xticks([0,1],['ALL','AML']);ax.set_title(title,fontsize=12);ax.set_xlim(-.5,1.5);ax.grid(axis='y',alpha=.2)
axs[0].set_ylabel('处理后的表达测量值')
fig.savefig(OUT/'probe_expression.png');plt.close(fig)

sim=pd.read_csv(ROOT/'.tmp/multiple-test/student-results/simulation_summary.csv')
fig,axs=plt.subplots(1,2,figsize=(7.2,3.2),layout='constrained')
for ax,metric,title in zip(axs,['FDR','power'],['错误发现率 FDR','检出力']):
    for method,label,col in zip(['raw','Bonferroni','BH'],['未校正','Bonferroni','BH'],colors):
        part=sim[(sim.signals.eq(100))&(sim.method.eq(method))].sort_values('delta')
        ax.plot(part.delta,part[metric]*100,marker='o',color=col,label=label)
    if metric=='FDR':ax.axhline(5,color='#A85845',linestyle='--',linewidth=1,label='5%目标水平')
    ax.set_xticks([.5,1.],['0.5','1.0']);ax.set_xlabel('真实均值差 δ');ax.set_ylabel(title+'（%）');ax.set_ylim(bottom=0);ax.grid(alpha=.2)
handles,labels=axs[0].get_legend_handles_labels()
fig.legend(handles,labels,loc='outside lower center',ncol=4,frameon=False,fontsize=9)
fig.savefig(OUT/'simulation.png');plt.close(fig)

v=pd.read_csv(ROOT/'.tmp/multiple-test/student-results/validation.csv')
fig,ax=plt.subplots(figsize=(7.2,3.2),layout='constrained')
for j,r in v.iterrows():
    col='#238A83' if r.passes else '#999999'
    ax.scatter(r.discovery_delta,r.validation_delta,color=col,s=38)
    ax.annotate(str(j+1),(r.discovery_delta,r.validation_delta),xytext=(4,4),textcoords='offset points',fontsize=9)
ax.axhline(0,color='#999999',linewidth=.7);ax.axvline(0,color='#999999',linewidth=.7)
ax.set_xlabel('发现样本均值差 AML−ALL');ax.set_ylabel('留出样本均值差 AML−ALL');ax.grid(alpha=.15)
fig.savefig(OUT/'validation.png');plt.close(fig)
print('Created 4 figures',OUT)
