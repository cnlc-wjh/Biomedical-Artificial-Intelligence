from pathlib import Path
import numpy as np
import pandas as pd
import scipy
from scipy import stats

ROOT = Path(__file__).resolve().parents[2]
df = pd.read_csv(ROOT / "data" / "leukemia.csv")
meta = ["sample_id", "original_split", "diagnosis"]
train = df[df.original_split.eq("train")]
test = df[df.original_split.eq("test")]
probes = [c for c in df.columns if c not in meta]
a = train.loc[train.diagnosis.eq("AML"), probes].to_numpy()
b = train.loc[train.diagnosis.eq("ALL"), probes].to_numpy()
p = stats.ttest_ind(a, b, axis=0, equal_var=False).pvalue
bh = stats.false_discovery_control(p, method="bh")
bon = np.minimum(len(p) * p, 1)
print("versions", np.__version__, pd.__version__, scipy.__version__)
print("data", df.shape, len(probes), train.diagnosis.value_counts().to_dict(), test.diagnosis.value_counts().to_dict())
print("counts", int((p <= .05).sum()), int((bon <= .05).sum()), int((bh <= .05).sum()))

def adjust(p):
    order = np.argsort(p, kind="stable")
    ranks = np.arange(1, len(p) + 1)
    sorted_adjusted = np.minimum.accumulate((p[order] * len(p) / ranks)[::-1])[::-1]
    adjusted = np.empty(len(p))
    adjusted[order] = np.minimum(sorted_adjusted, 1)
    return {"raw": p, "bonferroni": np.minimum(len(p) * p, 1), "bh": adjusted}


def simulate(delta, signals):
    rng = np.random.default_rng(42)
    m, repeats, n_aml, n_all = 1000, 1000, 11, 27
    truth = np.arange(m) < signals
    output = []
    for run in range(repeats):
        aa = rng.normal(size=(n_aml, m))
        bb = rng.normal(size=(n_all, m))
        aa[:, truth] += delta
        pv = stats.ttest_ind(aa, bb, axis=0, equal_var=True).pvalue
        adj = adjust(pv)
        for method, values in adj.items():
            rejected = values <= .05
            r = int(rejected.sum())
            v = int((rejected & ~truth).sum())
            s = int((rejected & truth).sum())
            output.append((run, method, r, v, s, v / max(r, 1), int(v > 0), s / signals if signals else np.nan))
    runs = pd.DataFrame(output, columns=["run", "method", "R", "V", "S", "FDP", "any_false", "power"])
    return runs.groupby("method", sort=False).agg(mean_R=("R", "mean"), mean_V=("V", "mean"), FWER=("any_false", "mean"), FDR=("FDP", "mean"), power=("power", "mean"))


for d, s in [(0, 0), (1.0, 100), (.5, 100)]:
    print("simulation", d, s)
    print(simulate(d, s).round(6).to_string())
