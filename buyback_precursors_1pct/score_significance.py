"""Значимость балла из 5 условий: 4–5 условий против 0 и рост доли выкупа с каждым условием.

Запуск: python3 score_significance.py   (после analyze.py; берёт dataset.csv и условия из analyze.py)
Выход:  score_significance.xlsx — тесты с учётом того, что годы одной компании не независимы.

Тесты: точный тест Фишера (как если бы годы были независимы); логит с ошибками, кластеризованными
по компаниям; GEE с обменной корреляцией внутри компании; бутстреп по компаниям (5000 повторов).
Проверка на отбор: пять условий заново выбираются только по 2017–2021 гг. и проверяются на 2022–2025 гг.
"""
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats

HERE = Path(__file__).absolute().parent
sys.path.insert(0, str(HERE))
import analyze as an  # noqa: E402

warnings.filterwarnings("ignore")
B = 5000
RNG = np.random.default_rng(20261024)


def points(d, items):
    return sum(f(d).fillna(False).astype(int) for f in items.values())


def cluster_logit(y, x, groups):
    m = sm.Logit(y.astype(float), sm.add_constant(x.astype(float))).fit(
        disp=0, cov_type="cluster", cov_kwds={"groups": pd.factorize(groups)[0]})
    b, se = m.params.iloc[1], m.bse.iloc[1]
    return np.exp(b), np.exp(b - 1.96 * se), np.exp(b + 1.96 * se), m.pvalues.iloc[1]


def gee_p(y, x, groups):
    m = sm.GEE(y.astype(float), sm.add_constant(x.astype(float)), groups=pd.factorize(groups)[0],
               family=sm.families.Binomial(), cov_struct=sm.cov_struct.Exchangeable()).fit()
    return m.pvalues.iloc[1]


def boot_diff(x):
    """Бутстреп по компаниям: разница и отношение долей «4–5 условий» и «0 условий»."""
    codes, uniq = pd.factorize(x.company)
    members = [np.flatnonzero(codes == k) for k in range(len(uniq))]
    hi, y = (x.grp == 4).values, x.bb.values
    diffs, ratios = [], []
    for _ in range(B):
        idx = np.concatenate([members[k] for k in RNG.integers(0, len(uniq), len(uniq))])
        h, yy = hi[idx], y[idx]
        if h.sum() == 0 or (~h).sum() == 0:
            continue
        p4, p0 = yy[h].mean(), yy[~h].mean()
        diffs.append(p4 - p0)
        ratios.append(p4 / p0 if p0 > 0 else np.nan)
    diffs, ratios = np.array(diffs), np.array(ratios)
    p_two = min(1.0, 2 * min((diffs <= 0).mean(), (diffs >= 0).mean()))
    return (np.percentile(diffs, 2.5), np.percentile(diffs, 97.5), np.nanpercentile(ratios, 2.5),
            np.nanpercentile(ratios, 97.5), max(p_two, 1 / len(diffs)))


def test_sample(name, d):
    d = d[d.pts.notna()]
    x = d[d.grp.isin([0, 4])]
    lo, hi = x[x.grp == 0], x[x.grp == 4]
    k0, n0, k4, n4 = int(lo.bb.sum()), len(lo), int(hi.bb.sum()), len(hi)
    fisher = stats.fisher_exact([[k4, n4 - k4], [k0, n0 - k0]])[1]
    or2, or2_lo, or2_hi, p2 = cluster_logit(x.bb, (x.grp == 4), x.company)
    ort, ort_lo, ort_hi, pt = cluster_logit(d.bb, d.pts, d.company)
    dl, dh, rl, rh, pb = boot_diff(x)
    return {"Выборка": name, "0 условий: лет": n0, "0 условий: с выкупом": k0, "0 условий: доля": k0 / n0,
            "4–5 условий: лет": n4, "4–5 условий: с выкупом": k4, "4–5 условий: доля": k4 / n4,
            "4–5 условий: компаний": hi.company.nunique(),
            "4–5 условий: компаний, выкупавших хотя бы раз": hi[hi.bb == 1].company.nunique(),
            "Разница долей": k4 / n4 - k0 / n0, "Разница: 95% интервал (бутстреп по компаниям), от": dl,
            "Разница: 95% интервал, до": dh, "Во сколько раз чаще": (k4 / n4) / (k0 / n0) if k0 else np.nan,
            "Во сколько раз: 95% интервал, от": rl, "Во сколько раз: 95% интервал, до": rh,
            "p: точный тест Фишера (годы как независимые)": fisher,
            "p: логит 4–5 против 0, ошибки по компаниям": p2,
            "p: GEE 4–5 против 0 (корреляция внутри компании)": gee_p(x.bb, (x.grp == 4), x.company),
            "p: бутстреп по компаниям": pb,
            "Отношение шансов на каждое условие (все уровни 0–5)": ort,
            "Отношение шансов на условие: 95% интервал, от": ort_lo, "Отношение шансов на условие: до": ort_hi,
            "p: рост с каждым условием, ошибки по компаниям": pt, "Компаний в выборке": d.company.nunique()}


def reselect_on_train(d, k=5, min_cover=0.7):
    """Пять условий без истории выкупов, выбранные только по 2017–2021 гг. (наибольшее |AUC − 0,5|)."""
    train = d[d.year <= 2021]
    rows = []
    for c in an.FEATS:
        if c in an.HISTORY or c not in d:
            continue
        x = train[[c, "bb", "year"]].dropna()
        if len(x) < min_cover * len(train) or x[c].nunique() < 2:
            continue
        a = an.auc(x.bb.values, an.score(x, c).values)
        rows.append((abs(a - 0.5), c, a))
    picked = sorted(rows, reverse=True)[:k]
    items = {}
    for _, c, a in picked:
        if c in an.BINARY:
            val = 1.0 if a >= 0.5 else 0.0
            items[f"{an.NAMES[c]}: {'да' if val else 'нет'} (AUC {a:.2f})"] = (lambda dd, c=c, v=val: an.score(dd, c) == v)
        elif a >= 0.5:
            items[f"{an.NAMES[c]}: верхняя треть года (AUC {a:.2f})"] = (lambda dd, c=c: an.score(dd, c) >= 2 / 3)
        else:
            items[f"{an.NAMES[c]}: нижняя треть года (AUC {a:.2f})"] = (lambda dd, c=c: an.score(dd, c) <= 1 / 3)
    return items


def main():
    raw = pd.read_csv(HERE / "dataset.csv")
    d = raw[raw.bb.notna()].reset_index(drop=True)
    d["pts"] = points(d, an.SCORE_ITEMS)
    d["grp"] = d.pts.clip(upper=4)
    special = d.bb_type.fillna("") != ""

    res = [test_sample("Все годы, 2017–2025", d),
           test_sample("2017–2021", d[d.year <= 2021]),
           test_sample("2022–2025", d[d.year >= 2022]),
           test_sample("Начало выкупа (в прошлом году не было)", d[d.f_bb_prev == 0]),
           test_sample("Без особых выкупов", d[~special])]

    picked = {}
    for cover in (0.5, 0.7):
        items = reselect_on_train(d, min_cover=cover)
        d2 = d.copy()
        d2["pts"] = points(d2, items)
        d2["grp"] = d2.pts.clip(upper=4)
        name = f"Условия выбраны по 2017–2021 (признаки с данными не меньше чем в {cover:.0%} лет), проверка на 2022–2025"
        res.append(test_sample(name, d2[d2.year >= 2022]))
        picked[name] = list(items)
    out = pd.DataFrame(res)
    picked = pd.DataFrame(picked)

    with pd.ExcelWriter(HERE / "score_significance.xlsx") as xw:
        out.to_excel(xw, sheet_name="Значимость балла", index=False)
        picked.T.to_excel(xw, sheet_name="Условия по 2017–2021", header=[f"Условие {i + 1}" for i in range(picked.shape[0])], index_label="Вариант")
    pd.set_option("display.width", 250)
    print(out.T.to_string())
    print(picked.to_string())


if __name__ == "__main__":
    main()
