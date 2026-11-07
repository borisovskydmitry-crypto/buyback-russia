"""Что предшествует погашению акций: те же пять условий и все признаки, но событие — погашение.

Запуск: python3 cancel_precursors.py   (после build_dataset.py и analyze.py)
Выход:  cancel_precursors.xlsx (русские листы) и cancel_summary.json.

Погашение в году t — число акций на конец t меньше, чем на конец t−1 (файл Анализ_акций_Мосбиржи.xlsx,
столбец cancel в dataset.csv). Выборка — годы, когда у компании есть капитализация на Мосбирже (иначе в файле
нет числа акций). Ранг показателя — место среди компаний того же года с известным выкупом (та же база, что
в анализе выкупов); год, не входящий в эту базу, ставится на то место, которое занял бы, если его добавить.
Событий мало (19 у 10 компаний, у Банка «Санкт-Петербург» 5, у Лукойла 3), поэтому интервалы и p-значения —
бутстреп по компаниям; точный тест Фишера считает годы независимыми и показан для сравнения.
"""
import json
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

HERE = Path(__file__).absolute().parent
sys.path.insert(0, str(HERE))
import analyze as an  # noqa: E402

warnings.filterwarnings("ignore")
B = 5000
RNG = np.random.default_rng(20261024)
NAME = {"bb_now": "Выкуп ≥ 0,1% в том же году", "f_treasury_pos": "Казначейские акции на балансе на конец t−1"}


_CACHE = {}


def rel_rank(d, ref, c):
    key = (c, len(d), int(ref.sum()))
    if key not in _CACHE:
        _CACHE[key] = _rel_rank(d, ref, c)
    return _CACHE[key]


def _rel_rank(d, ref, c):
    """Ранг внутри года среди строк ref; строка вне ref ставится так, как если бы её добавили."""
    out = pd.Series(np.nan, index=d.index)
    for y, g in d.groupby("year"):
        base = np.sort(g.loc[ref.loc[g.index] & g[c].notna(), c].values)
        n = len(base)
        if n == 0:
            continue
        for i, v in g[c].items():
            if pd.isna(v):
                continue
            less = np.searchsorted(base, v, "left")
            eq = np.searchsorted(base, v, "right") - less
            out[i] = (less + (eq + 1) / 2) / n if ref[i] else (less + (eq + 2) / 2) / (n + 1)
    return out


def conditions(d):
    ref = d.bb.notna()
    rk = {c: rel_rank(d, ref, c) for c in ["f_roe", "f_pbv", "f_cash_assets", "f_cash_growth"]}
    return pd.DataFrame({
        "Частная компания (не госкомпания)": d.state == 0,
        "ROE в верхней трети года": rk["f_roe"] >= 2 / 3,
        "P/BV в верхней трети года": rk["f_pbv"] >= 2 / 3,
        "Денежные средства / активы в верхней трети года": rk["f_cash_assets"] >= 2 / 3,
        "Прирост денежных средств в верхней трети года": rk["f_cash_growth"] >= 2 / 3,
    }, index=d.index), rk


def boot(d, stat, b=B):
    """Бутстреп по компаниям: 95% интервал и двусторонний p против нуля."""
    codes, uniq = pd.factorize(d.company)
    members = [np.flatnonzero(codes == k) for k in range(len(uniq))]
    vals = []
    for _ in range(b):
        idx = np.concatenate([members[k] for k in RNG.integers(0, len(uniq), len(uniq))])
        v = stat(d.iloc[idx])
        if np.isfinite(v):
            vals.append(v)
    vals = np.array(vals)
    p = min(1.0, 2 * min((vals <= 0).mean(), (vals >= 0).mean()))
    return np.percentile(vals, 2.5), np.percentile(vals, 97.5), max(p, 1 / len(vals))


def auc_vs(d, c, ev, other):
    x = d[(ev | other) & d[c].notna()]
    return an.auc(ev[x.index].values, x[c].values)


def score_table(d, pts, ev, label):
    rows = []
    for k, m in [("0", pts == 0), ("1", pts == 1), ("2", pts == 2), ("3", pts == 3), ("4–5", pts >= 4)]:
        rows.append({"Выборка": label, "Балл": k, "Лет": int(m.sum()), "Из них с погашением": int(ev[m].sum()),
                     "Доля лет с погашением": ev[m].mean() if m.any() else np.nan,
                     "Компаний с погашением в этих годах": d[m & ev].company.nunique()})
    return rows


def main():
    raw = pd.read_csv(HERE / "dataset.csv")
    d = raw[raw.cap.notna()].reset_index(drop=True)
    cond, rk = conditions(d)
    d["pts"] = cond.sum(axis=1)
    d["f_treasury_pos"] = (d.f_treasury_cap > 0).astype(float).where(d.f_treasury_cap.notna())
    d["bb_now"] = d.bb
    ev = d.cancel == 1
    buy_only = (d.bb == 1) & ~ev
    rest = ~ev
    cancellers = set(d[ev].company)
    inside = d.company.isin(cancellers)
    first = ev & ~d[ev].duplicated("company").reindex(d.index, fill_value=False)

    # проверка: на строках с известным выкупом условия совпадают с анализом выкупов
    k = d.bb.notna()
    ref_pts = sum(f(d[k]).fillna(False).astype(int) for f in an.SCORE_ITEMS.values())
    assert (ref_pts == d.pts[k]).all(), "ранги расходятся с analyze.py"

    # 1. балл и доля лет с погашением
    t1 = pd.DataFrame(score_table(d, d.pts, ev, "Все годы с торгами, 2017–2025")
                      + score_table(d[inside], d.pts[inside], ev[inside], "Только 10 компаний, гасивших акции"))
    t1b = []
    for lab, m in [("Все годы с торгами", d.index == d.index), ("Только гасившие компании", inside)]:
        x = d[m].assign(hi=(d.pts[m] >= 4), y=ev[m].astype(int))
        x2 = x[(x.pts >= 4) | (x.pts <= 1)]
        tab = [[int(x2[x2.hi].y.sum()), int((x2.hi & (x2.y == 0)).sum())],
               [int(x2[~x2.hi].y.sum()), int((~x2.hi & (x2.y == 0)).sum())]]
        lo, hi_, pb = boot(x2, lambda z: z[z.hi].y.mean() - z[~z.hi].y.mean())
        t1b.append({"Выборка": lab, "4–5 условий: доля лет с погашением": x2[x2.hi].y.mean(),
                    "0–1 условие: доля лет с погашением": x2[~x2.hi].y.mean(),
                    "Разница": x2[x2.hi].y.mean() - x2[~x2.hi].y.mean(), "95% интервал, от": lo, "до": hi_,
                    "p: бутстреп по компаниям": pb, "p: точный тест Фишера (годы независимы)": stats.fisher_exact(tab)[1],
                    "Средний балл в годы погашения": d.pts[m & ev].mean(), "Средний балл в остальные годы": d.pts[m & ~ev].mean()})
        x["ev"] = x.y
        lo, hi_, pb = boot(x, lambda z: z[z.ev == 1].pts.mean() - z[z.ev == 0].pts.mean())
        t1b[-1].update({"Разница средних баллов: 95% интервал, от": lo, "до ": hi_, "p (разница баллов)": pb})
    t1b = pd.DataFrame(t1b)

    # 2. условия: погашения против остальных лет и против лет с выкупом без погашения
    flags = cond.astype(float).copy()
    flags["Балл 4–5"] = (d.pts >= 4).astype(float)
    for c in ["f_treasury_pos", "f_bb_prev", "bb_now"]:
        flags[NAME.get(c, an.NAMES.get(c))] = d[c]
    t2 = []
    for name in flags:
        f = flags[name]
        row = {"Условие": name}
        for lab, m in [("годы погашения", ev), ("одно погашение на компанию (первое)", first),
                       ("остальные годы", rest), ("годы выкупа без погашения", buy_only)]:
            z = f[m].dropna()
            row[f"Доля: {lab}"] = z.mean() if len(z) else np.nan
            row[f"Лет с данными: {lab}"] = len(z)
        z = d.assign(f=f, e=ev.astype(int))[f.notna()]
        a, b_ = int(z[z.e == 1].f.sum()), int((z.e == 1).sum())
        c_, n_ = int(z[z.e == 0].f.sum()), int((z.e == 0).sum())
        row["Во сколько раз чаще, чем в остальные годы"] = (a / b_) / (c_ / n_) if c_ else np.nan
        lo, hi_, pb = boot(z, lambda q: q[q.e == 1].f.mean() - q[q.e == 0].f.mean())
        row.update({"Разница с остальными годами": a / b_ - c_ / n_, "95% интервал, от": lo, "до": hi_,
                    "p: бутстреп по компаниям": pb,
                    "p: точный тест Фишера": stats.fisher_exact([[a, b_ - a], [c_, n_ - c_]])[1]})
        t2.append(row)
    t2 = pd.DataFrame(t2)

    # 3. все признаки: средний ранг года погашения, AUC против остальных и против выкупов без погашения
    t3 = []
    for c, (nm, grp) in an.FEATS.items():
        if c not in d or c == "f_cancel_prev":
            continue
        s = d[c].astype(float) if c in an.BINARY else rel_rank(d, d.bb.notna(), c)
        z = d.assign(s=s, e=ev.astype(int))[s.notna()]
        if z.e.sum() < 5:
            continue
        a_all = an.auc(z.e.values, z.s.values)
        lo, hi_, pb = boot(z, lambda q: an.auc(q.e.values, q.s.values) - 0.5 if 0 < q.e.sum() < len(q) else np.nan, 2000)
        t3.append({"Признак": nm, "Группа": grp, "Тип": "доля (флаг)" if c in an.BINARY else "средний ранг в году",
                   "Погашения: среднее": z[z.e == 1].s.mean(), "Остальные годы: среднее": z[z.e == 0].s.mean(),
                   "Годы выкупа без погашения: среднее": s[buy_only].mean(),
                   "AUC: погашения против остальных лет": a_all, "AUC: 95% интервал, от": lo + 0.5, "до": hi_ + 0.5,
                   "p: бутстреп по компаниям": pb,
                   "p: Манна–Уитни (годы независимы)": stats.mannwhitneyu(z[z.e == 1].s, z[z.e == 0].s).pvalue,
                   "AUC: погашения против лет выкупа без погашения": auc_vs(d.assign(s=s), "s", ev, buy_only),
                   "Погашений с данными": int(z.e.sum())})
    t3 = pd.DataFrame(t3).sort_values("p: бутстреп по компаниям")
    from statsmodels.stats.multitest import multipletests
    t3["p с поправкой Бенджамини–Хохберга"] = multipletests(t3["p: бутстреп по компаниям"], method="fdr_bh")[1]

    # 4. внутри гасивших компаний: годы погашения против их же остальных лет
    t4 = []
    for c, (nm, grp) in an.FEATS.items():
        if c not in d or c in an.BINARY:
            continue
        s = rel_rank(d, d.bb.notna(), c)
        z = d[inside].assign(s=s[inside], e=ev[inside])[s[inside].notna()]
        per = z.groupby(["company", "e"]).s.mean().unstack()
        per = per.dropna()
        if len(per) < 6:
            continue
        diff = per[True] - per[False]
        t4.append({"Признак": nm, "Компаний": len(per), "Средний ранг в годы погашения": per[True].mean(),
                   "Средний ранг в их остальные годы": per[False].mean(), "Разница": diff.mean(),
                   "Компаний, где в год погашения выше": int((diff > 0).sum()),
                   "p: знаковый ранговый тест Уилкоксона по компаниям": stats.wilcoxon(diff).pvalue})
    t4 = pd.DataFrame(t4).sort_values("p: знаковый ранговый тест Уилкоксона по компаниям")

    # 5. события
    t5 = d[ev][["company", "year", "cancel_pct", "bb", "bb_share", "f_bb_prev", "f_treasury_cap",
                "f_roe", "f_pbv", "f_cash_assets", "f_cash_growth"]].copy()
    for c in ["f_roe", "f_pbv", "f_cash_assets", "f_cash_growth"]:
        t5[f"r_{c}"] = rk[c][ev]
    t5 = t5.join(cond[ev].replace({True: "да", False: "нет"})).assign(pts=d.pts[ev])
    t5 = t5.rename(columns={"company": "Тикер", "year": "Год погашения (t)", "cancel_pct": "Погашено, доля акций",
                            "bb": "Выкуп ≥ 0,1% в году t (1 — да)", "bb_share": "Выкуп в году t, доля капитализации",
                            "f_bb_prev": "Выкуп в прошлом году (1 — да)", "f_treasury_cap": "Казначейские акции / капитализация на конец t−1",
                            "f_roe": "ROE", "f_pbv": "P/BV", "f_cash_assets": "Деньги / активы",
                            "f_cash_growth": "Прирост денег / активы", "r_f_roe": "Ранг ROE", "r_f_pbv": "Ранг P/BV",
                            "r_f_cash_assets": "Ранг денег / активов", "r_f_cash_growth": "Ранг прироста денег",
                            "pts": "Балл (из 5)"})

    with pd.ExcelWriter(HERE / "cancel_precursors.xlsx") as xw:
        t1.to_excel(xw, sheet_name="Балл и погашение", index=False)
        t1b.to_excel(xw, sheet_name="Балл — значимость", index=False)
        t2.to_excel(xw, sheet_name="Условия", index=False)
        t3.to_excel(xw, sheet_name="Все признаки", index=False)
        t4.to_excel(xw, sheet_name="Внутри гасивших компаний", index=False)
        t5.to_excel(xw, sheet_name="Погашения", index=False)
    summary = {"events": int(ev.sum()), "companies": len(cancellers), "rows": len(d),
               "score": t1.to_dict("records"), "score_test": t1b.to_dict("records"), "conditions": t2.to_dict("records")}
    (HERE / "cancel_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1, default=float))
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 30)
    print(t1.to_string()); print(t1b.T.to_string()); print(t2.T.to_string())
    print(t3.head(20).to_string()); print(t4.head(12).to_string()); print(t5.to_string())


if __name__ == "__main__":
    main()
