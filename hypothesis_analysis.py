"""四中心心脏病假设检验。运行本文件生成 results/ 下的表格、图和摘要。

主分析遵循课程指定的终点与双侧方向。其余模型均为探索性敏感性分析。
依赖：numpy、pandas、scipy、matplotlib；不需要下载任何数据。
"""
from pathlib import Path
import hashlib
import json
import platform
import sys
import numpy as np
import pandas as pd
import scipy
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parent
DATA = BASE / 'V3/data/heart-disease-hypothesis.csv'
OUT = BASE / 'results'
OUT.mkdir(exist_ok=True)
ALPHA = 0.05


def welch(x, y):
    """差值方向始终为 x 减 y；手算公式与 SciPy 相互核对。"""
    x, y = np.asarray(x, dtype=float), np.asarray(y, dtype=float)
    n1, n0 = len(x), len(y)
    a, b = x.var(ddof=1) / n1, y.var(ddof=1) / n0
    delta = x.mean() - y.mean()
    se = np.sqrt(a + b)
    df = (a + b)**2 / (a*a/(n1-1) + b*b/(n0-1))
    t_value = delta / se
    p = 2 * stats.t.sf(abs(t_value), df)
    margin = stats.t.ppf(1-ALPHA/2, df) * se
    check = stats.ttest_ind(x, y, equal_var=False, alternative='two-sided')
    ci_check = check.confidence_interval(1-ALPHA)
    assert np.allclose([t_value, p, df, delta-margin, delta+margin],
                       [check.statistic, check.pvalue, check.df, ci_check.low, ci_check.high])
    return dict(n1=n1, n0=n0, delta=delta, se=se, t=t_value, df=df,
                p=p, ci_low=delta-margin, ci_high=delta+margin)


def ols_hc3(frame, covariates, with_site=False):
    """OLS 点估计及 HC3 异方差稳健标准误，使用渐近正态区间。

    HC3 不处理未测量混杂或残余中心内相关；不能据此作因果解释。
    """
    features = frame[covariates].astype(float).copy()
    if with_site:
        sites = pd.get_dummies(frame['site'], prefix='site', drop_first=True, dtype=float)
        features = pd.concat([features, sites], axis=1)
    features.insert(0, 'Intercept', 1.0)
    X, y = features.to_numpy(), frame['thalach'].to_numpy()
    assert np.linalg.matrix_rank(X) == X.shape[1]
    beta = np.linalg.lstsq(X, y, rcond=None)[0]
    residual = y - X @ beta
    bread = np.linalg.inv(X.T @ X)
    leverage = np.einsum('ij,jk,ik->i', X, bread, X)
    adjusted_resid = residual / (1-leverage)
    meat = X.T @ (X * adjusted_resid[:, None]**2)
    covariance = bread @ meat @ bread
    se = np.sqrt(np.diag(covariance))
    rows = []
    for name, estimate, stderr in zip(features.columns, beta, se):
        z = estimate / stderr
        rows.append(dict(term=name, estimate=estimate, se_hc3=stderr,
                         ci_low=estimate-stats.norm.ppf(.975)*stderr,
                         ci_high=estimate+stats.norm.ppf(.975)*stderr,
                         p=2*stats.norm.sf(abs(z)), n=len(y)))
    return pd.DataFrame(rows)


def main():
    df = pd.read_csv(DATA)  # CSV 空白默认保留为 NaN，不能填成 0。
    assert df['record_id'].is_unique
    assert df['disease'].isin([0, 1]).all()
    assert np.array_equal(df['disease'], (df['num'] > 0).astype(int))
    assert np.isfinite(df[['age', 'sex', 'thalach']].dropna().to_numpy()).all()
    complete = df.dropna(subset=['disease', 'thalach']).copy()
    x = complete.loc[complete['disease'] == 1, 'thalach'].to_numpy()
    y = complete.loc[complete['disease'] == 0, 'thalach'].to_numpy()
    primary = welch(x, y)
    descriptives = []
    for group in [0, 1]:
        raw = df.loc[df['disease'] == group]
        values = raw['thalach'].dropna()
        descriptives.append(dict(disease=group, original_n=len(raw), n=len(values),
            missing_n=raw['thalach'].isna().sum(), missing_pct=100*raw['thalach'].isna().mean(),
            mean=values.mean(), sd=values.std(ddof=1), median=values.median(),
            q1=values.quantile(.25), q3=values.quantile(.75), minimum=values.min(), maximum=values.max()))
    pd.DataFrame(descriptives).to_csv(OUT/'group_summary.csv', index=False, encoding='utf-8-sig')

    missing_profile = []
    for missing, sub in df.groupby(df['thalach'].isna()):
        row = dict(thalach_missing=missing, n=len(sub), age_mean=sub['age'].mean(),
                   age_sd=sub['age'].std(), male_pct=100*sub['sex'].mean(),
                   disease_pct=100*sub['disease'].mean())
        for site in sorted(df['site'].unique()):
            row[site+'_n'] = int((sub['site'] == site).sum())
        missing_profile.append(row)
    pd.DataFrame(missing_profile).to_csv(OUT/'missing_profile.csv', index=False, encoding='utf-8-sig')
    site_missing = df.groupby('site')['thalach'].agg(total='size', complete='count')
    site_missing['missing'] = site_missing['total'] - site_missing['complete']
    site_missing.to_csv(OUT/'site_missing.csv', encoding='utf-8-sig')

    models = []
    for label, covars, site in [('unadjusted', ['disease'], False),
                                ('age_sex', ['disease','age','sex'], False),
                                ('age_sex_site', ['disease','age','sex'], True)]:
        frame = complete.dropna(subset=covars+(['site'] if site else []))
        fit = ols_hc3(frame, covars, site)
        fit.to_csv(OUT/f'ols_{label}_hc3.csv', index=False, encoding='utf-8-sig')
        row = fit.loc[fit['term']=='disease'].iloc[0].to_dict()
        row['model'] = label
        models.append(row)
    assert np.isclose(models[0]['estimate'], primary['delta'])
    sites = []
    for site, frame in complete.groupby('site'):
        a = frame.loc[frame['disease']==1, 'thalach']
        b = frame.loc[frame['disease']==0, 'thalach']
        sites.append(dict(site=site, **welch(a,b)))
    pd.DataFrame(sites).to_csv(OUT/'site_welch.csv', index=False, encoding='utf-8-sig')

    u = stats.mannwhitneyu(x, y, alternative='two-sided', method='asymptotic')
    tab = pd.crosstab(df['sex'], df['disease'])
    chi = stats.chi2_contingency(tab, correction=False)
    tab.to_csv(OUT/'sex_disease_table.csv', encoding='utf-8-sig')
    corr = stats.pearsonr(complete['age'], complete['thalach'])
    corr_ci = corr.confidence_interval(.95)
    # 前瞻性规划示例：独立两组、近似正态、等分配、共同SD=24。
    # delta 是未来希望检出的目标差异，不使用当前观察差异计算“事后power”。
    power_rows = []
    for target in [5, 10, 15]:
        zsum = stats.norm.ppf(.975) + stats.norm.ppf(.8)
        n_per_group = int(np.ceil(2 * (zsum*24/target)**2))
        power_rows.append(dict(target_delta=target, assumed_sd=24, alpha=.05,
                               power=.8, n_per_group_normal_approx=n_per_group))
    pd.DataFrame(power_rows).to_csv(OUT/'sample_size_plan.csv', index=False, encoding='utf-8-sig')

    plt.rcParams.update({'font.sans-serif':['Microsoft YaHei','SimHei','DejaVu Sans'],
                         'axes.unicode_minus':False, 'font.size':10})
    colors = ['#397293','#B56A3F']
    fig, axes = plt.subplots(1,2,figsize=(9.2,3.4), layout='constrained')
    bins = np.arange(50,221,10)
    for values, label, color in [(y,'无疾病证据（0）',colors[0]),(x,'有疾病证据（1）',colors[1])]:
        axes[0].hist(values, bins=bins, alpha=.5, density=True, color=color, label=label)
    axes[0].set(xlabel='最大运动心率（次/分钟）', ylabel='概率密度')
    axes[0].legend(fontsize=8)
    boxes=axes[1].boxplot([y,x], tick_labels=['无疾病证据','有疾病证据'], patch_artist=True,
                          medianprops={'color':'black'}, widths=.5)
    for patch,color in zip(boxes['boxes'],colors): patch.set_facecolor(color); patch.set_alpha(.55)
    axes[1].set(ylabel='最大运动心率（次/分钟）')
    fig.savefig(OUT/'thalach_distribution.png', dpi=200); plt.close(fig)

    fig,ax=plt.subplots(figsize=(7.7,3.4),layout='constrained')
    labels=['总体：Welch','调整年龄、性别、中心']+[r['site'] for r in sites]
    adjusted=models[-1]
    estimates=[primary['delta'],adjusted['estimate']]+[r['delta'] for r in sites]
    low=[primary['ci_low'],adjusted['ci_low']]+[r['ci_low'] for r in sites]
    high=[primary['ci_high'],adjusted['ci_high']]+[r['ci_high'] for r in sites]
    for i,(m,l,h) in enumerate(zip(estimates,low,high)):
        ax.errorbar(m,i,xerr=[[m-l],[h-m]],fmt='o',color=colors[0] if i==0 else colors[1],capsize=3)
    ax.set_yticks(range(len(labels)),labels); ax.invert_yaxis(); ax.axvline(0,color='gray',linestyle='--')
    ax.set(xlabel='疾病组减无疾病组：心率差与95% CI（次/分钟）')
    fig.savefig(OUT/'sensitivity_forest.png',dpi=200); plt.close(fig)
    summary=dict(primary=primary, descriptives=descriptives, missing_profile=missing_profile,
        adjusted_models=models, sites=sites,
        mannwhitney=dict(u=float(u.statistic),p=float(u.pvalue),probability_superiority=float(u.statistic/(len(x)*len(y)))),
        chi_square=dict(statistic=float(chi.statistic),p=float(chi.pvalue),min_expected=float(chi.expected_freq.min())),
        pearson=dict(r=float(corr.statistic),p=float(corr.pvalue),ci_low=float(corr_ci.low),ci_high=float(corr_ci.high)),
        sample_size_plan=power_rows, family_false_positive=1-.95**20,
        sha256=hashlib.sha256(DATA.read_bytes()).hexdigest(),
        versions=dict(python=platform.python_version(),numpy=np.__version__,pandas=pd.__version__,
                      scipy=scipy.__version__,matplotlib=matplotlib.__version__))
    (OUT/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,default=lambda x:x.item()),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2,default=lambda x:x.item()))


if __name__ == '__main__':
    main()
