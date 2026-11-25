from pathlib import Path
import json, hashlib, sys
import numpy as np
import pandas as pd
import scipy
from scipy import stats

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / '.tmp/multiple-test/student-results'
OUT.mkdir(parents=True, exist_ok=True)
df = pd.read_csv(ROOT / 'data/leukemia.csv')
meta = ['sample_id','original_split','diagnosis']
cols = df.columns.drop(meta)
train = df[df.original_split.eq('train')]
test = df[df.original_split.eq('test')]
a = train.loc[train.diagnosis.eq('AML'),cols].to_numpy(float)
b = train.loc[train.diagnosis.eq('ALL'),cols].to_numpy(float)
m = len(cols)
t, p = stats.ttest_ind(a,b,axis=0,equal_var=False)
order = np.argsort(p,kind='stable')
scaled = m*p[order]/np.arange(1,m+1)
bh_manual = np.empty(m)
bh_manual[order] = np.minimum(1,np.minimum.accumulate(scaled[::-1])[::-1])
bh = stats.false_discovery_control(p,method='bh')
assert np.allclose(bh,bh_manual,rtol=1e-13,atol=1e-16)
bon = np.minimum(1,m*p)
delta = a.mean(0)-b.mean(0)
v1,v0 = a.var(0,ddof=1),b.var(0,ddof=1)
se = np.sqrt(v1/len(a)+v0/len(b))
dof = (v1/len(a)+v0/len(b))**2/((v1/len(a))**2/(len(a)-1)+(v0/len(b))**2/(len(b)-1))
manual_p = 2*stats.t.sf(np.abs(delta/se),dof)
assert np.allclose(p,manual_p,rtol=1e-12,atol=1e-15)
sp = np.sqrt(((len(a)-1)*v1+(len(b)-1)*v0)/(len(a)+len(b)-2))
half = stats.t.ppf(.975,dof)*se
res = pd.DataFrame({'probe':cols,'mean_AML':a.mean(0),'mean_ALL':b.mean(0),'sd_AML':np.sqrt(v1),'sd_ALL':np.sqrt(v0),'delta':delta,'ci_low':delta-half,'ci_high':delta+half,'pooled_sd':sp,'cohen_d':delta/sp,'raw_p':p,'bonferroni':bon,'bh':bh})
res = res.sort_values(['raw_p','probe']).reset_index(drop=True)
res.insert(1,'rank',np.arange(1,m+1))
res.to_csv(OUT/'discovery_all.csv',index=False,encoding='utf-8-sig')
locked = res.loc[res.bh.le(.05)].head(10).copy()
locked.to_csv(OUT/'locked_candidates.csv',index=False,encoding='utf-8-sig')
# Save the discovery list before accessing held-out values. This rerun is not a new blind validation.
ac = test.loc[test.diagnosis.eq('AML'),locked.probe].to_numpy(float)
bc = test.loc[test.diagnosis.eq('ALL'),locked.probe].to_numpy(float)
tp = stats.ttest_ind(ac,bc,axis=0,equal_var=False).pvalue
vd = ac.mean(0)-bc.mean(0)
val = pd.DataFrame({'probe':locked.probe,'discovery_delta':locked.delta,'validation_delta':vd,'raw_p':tp,'bonferroni':np.minimum(len(locked)*tp,1)})
val['same_direction'] = np.sign(val.discovery_delta)==np.sign(vd)
val['passes'] = val.bonferroni.le(.05)&val.same_direction
val.to_csv(OUT/'validation.csv',index=False,encoding='utf-8-sig')
simrows=[]
for delta_sim, signals in [(0.,0),(1.,100),(.5,100)]:
    rng=np.random.default_rng(42)
    records=[]
    for run in range(1000):
        aa=rng.normal(size=(11,1000)); bb=rng.normal(size=(27,1000))
        aa[:,:signals]+=delta_sim
        pv=stats.ttest_ind(aa,bb,axis=0,equal_var=True).pvalue
        for method, ap in [('raw',pv),('Bonferroni',np.minimum(1000*pv,1)),('BH',stats.false_discovery_control(pv))]:
            rej=ap<=.05; R=int(rej.sum()); S=int(rej[:signals].sum()); V=R-S
            records.append({'run':run,'method':method,'R':R,'V':V,'FDP':V/max(R,1),'any_false':int(V>0),'power':S/signals if signals else np.nan})
    runs=pd.DataFrame(records)
    for method,g in runs.groupby('method',sort=False):
        simrows.append({'delta':delta_sim,'signals':signals,'method':method,'first_R':int(g.iloc[0].R),'mean_R':g.R.mean(),'sd_R':g.R.std(ddof=1),'q05':g.R.quantile(.05),'q95':g.R.quantile(.95),'mean_V':g.V.mean(),'FWER':g.any_false.mean(),'FDR':g.FDP.mean(),'power':g.power.mean()})
sim=pd.DataFrame(simrows)
sim.to_csv(OUT/'simulation_summary.csv',index=False,encoding='utf-8-sig')
i=cols.get_loc('U22376_cds2_s_at')
argmin=int(np.argmin(scaled))
summary={'data_sha256':hashlib.sha256((ROOT/'data/leukemia.csv').read_bytes()).hexdigest(),'shape':list(df.shape),'n_probes':m,'missing_values':int(df.isna().sum().sum()),'duplicate_ids':int(df.sample_id.duplicated().sum()),'duplicate_expression_rows':int(df[cols].duplicated().sum()),'constant_probes':int((df[cols].nunique()==1).sum()),'nonfinite_values':int((~np.isfinite(df[cols].to_numpy())).sum()),'expression_min':float(df[cols].min().min()),'expression_max':float(df[cols].max().max()),'counts':{'raw':int((p<=.05).sum()),'Bonferroni':int((bon<=.05).sum()),'BH':int((bh<=.05).sum())},'example':res.loc[res.probe.eq('U22376_cds2_s_at')].iloc[0].to_dict(),'bh_min_rank':argmin+1,'bh_min_value':float(scaled[argmin]),'bh_rank2_value':float(scaled[1]),'bh_boundary_p':float(p[order[534]]),'bon_raw_threshold':.05/m,'bh_boundary_threshold':535*.05/m,'validation_passes':int(val.passes.sum()),'versions':{'Python':sys.version.split()[0],'numpy':np.__version__,'pandas':pd.__version__,'scipy':scipy.__version__}}
(OUT/'analysis_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(summary,ensure_ascii=False,indent=2))
print(sim.to_string(index=False))
