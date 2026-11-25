from pathlib import Path
from zipfile import ZipFile
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.naive_bayes import BernoulliNB

root = Path(__file__).resolve().parent
base = root / '7. Naive-Bayes'
out = root / '报告成果'
sys.path.insert(0, str(base / 'scripts'))
from lab_utils import FEATURES, fit_model

train = pd.read_csv(base / 'data/student/train.csv')
val = pd.read_csv(base / 'data/student/validation.csv')
y = val['bladder_inflammation']
original = fit_model(train).predict_proba(val[FEATURES])[:, 1]
cols = FEATURES + ['lumbar_pain'] * 5
model = BernoulliNB(alpha=1, binarize=None).fit(train[cols].to_numpy(), train['bladder_inflammation'])
copied = model.predict_proba(val[cols].to_numpy())[:, 1]
audit = pd.read_csv(out / 'duplicate_audit.csv')
plt.rcParams.update({'font.sans-serif': ['Microsoft YaHei', 'SimHei'], 'axes.unicode_minus': False, 'font.size': 9, 'savefig.dpi': 200})
fig, axes = plt.subplots(1, 2, figsize=(8.3, 3.5))
for label, color, name in [(0, '#3b4cc0', '真实阴性'), (1, '#b40426', '真实阳性')]:
    mask = y == label
    axes[0].scatter(original[mask], copied[mask], color=color, s=32, alpha=.8, label=name)
axes[0].plot([0, 1], [0, 1], '--', color='gray', label='复制前后相等')
axes[0].set(xlabel='原模型阳性概率', ylabel='额外复制腰痛列 5 次后的阳性概率', xlim=(-.02, 1.02), ylim=(-.02, 1.02))
axes[0].legend(loc='upper center', bbox_to_anchor=(.5, 1.24), ncol=3, fontsize=8, frameon=False, handletextpad=.3, columnspacing=.8)
axes[1].bar(['0 次', '1 次', '5 次'], audit['log_loss'], color=['#285879', '#7296ad', '#bd6b36'])
axes[1].set(xlabel='额外复制腰痛列次数', ylabel='验证集平均 log loss')
fig.tight_layout()
image = out / 'figures/05_duplicate_图例补充.png'
fig.savefig(image, bbox_inches='tight', facecolor='white')
plt.close(fig)

source = out / '朴素贝叶斯分类实验报告.docx'
destination = out / '朴素贝叶斯分类实验报告_图例补充版.docx'
with ZipFile(source) as src, ZipFile(destination, 'w') as dst:
    for item in src.infolist():
        dst.writestr(item, image.read_bytes() if item.filename == 'word/media/image5.png' else src.read(item.filename))
print(destination)
print(image)
