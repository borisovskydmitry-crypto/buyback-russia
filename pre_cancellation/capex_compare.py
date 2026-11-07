"""Насколько низкие капзатраты/выручка погашавших компаний необычны для рынка.

Запуск: python3 capex_compare.py (после build_panel.py). Состав IMOEX на конец года — ISS
(imoex_constituents.csv). Банки исключены (у них нет выручки и капзатрат в том же смысле).
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).absolute().parent
sys.path.insert(0, str(HERE))
from build_panel import FILE_TICKERS, SECTOR  # noqa: E402

p = pd.read_csv(HERE / "panel.csv")
idx = pd.read_csv(HERE / "imoex_constituents.csv")
tick2co = {t: v[0] for v in FILE_TICKERS.values() for t in v}
idx["company"] = idx.ticker.map(tick2co)
idx = idx.dropna(subset=["company"])
mem = set(zip(idx.company, idx.year))
p["in_imoex"] = [(c, y) in mem for c, y in zip(p.company, p.year)]
d = p[p.capex_rev.notna() & ~p.is_bank].copy()
d["sector"] = d.company.map(SECTOR).fillna("прочие")
pre = d[d.pre_cancel]
thr = pre.capex_rev.median()
q25, q75 = pre.capex_rev.quantile([0.25, 0.75])

rows = []
for name, g in [("Все остальные компании", d[d.control]), ("В индексе Мосбиржи", d[d.control & d.in_imoex]),
                ("Вне индекса", d[d.control & ~d.in_imoex])]:
    rows.append({"Группа": name, "Компаний": g.company.nunique(), "Компаний-лет": len(g),
                 "Медиана капзатраты/выручка": g.capex_rev.median(),
                 f"Доля с капзатратами ≤ {thr:.1%} (медиана погашавших)": (g.capex_rev <= thr).mean(),
                 "Доля с капзатратами ≤ 10%": (g.capex_rev <= 0.10).mean()})
rows.append({"Группа": "Погашавшие, год до погашения", "Компаний": pre.company.nunique(), "Компаний-лет": len(pre),
             "Медиана капзатраты/выручка": thr,
             f"Доля с капзатратами ≤ {thr:.1%} (медиана погашавших)": (pre.capex_rev <= thr).mean(),
             "Доля с капзатратами ≤ 10%": (pre.capex_rev <= 0.10).mean()})
summary = pd.DataFrame(rows)

by_sector = d[d.control].groupby("sector").capex_rev.agg(["median", "count"]).rename(
    columns={"median": "Медиана капзатраты/выручка", "count": "Компаний-лет"})
by_sector["Доля ≤ медианы погашавших"] = d[d.control].groupby("sector").capex_rev.apply(lambda s: (s <= thr).mean())
pre_sector = pre.groupby("sector").capex_rev.median().rename("Медиана у погашавших (год до)")
by_sector = by_sector.join(pre_sector).reset_index().rename(columns={"sector": "Отрасль"})

co = d[d.control].groupby("company").agg(capex=("capex_rev", "median"), years=("year", "count"),
                                          imoex=("in_imoex", "max"))
co["sector"] = co.index.map(SECTOR).fillna("прочие")
co["low"] = co.capex <= thr
co = co.sort_values("capex").reset_index().rename(columns={
    "company": "Компания (тикер)", "capex": "Медиана капзатраты/выручка за годы без погашения",
    "years": "Лет с данными", "imoex": "Была в индексе", "sector": "Отрасль",
    "low": f"Не выше {thr:.1%}"})
prof = pre[["company", "year", "sector", "capex_rev", "in_imoex"]].copy()
prof["year"] += 1
prof.columns = ["Компания (тикер)", "Год погашения", "Отрасль", "Капзатраты/выручка за год до", "В индексе"]

with pd.ExcelWriter(HERE / "capex_comparison.xlsx") as xw:
    summary.to_excel(xw, sheet_name="Итоги", index=False)
    by_sector.to_excel(xw, sheet_name="По отраслям", index=False)
    co.to_excel(xw, sheet_name="По компаниям", index=False)
    prof.to_excel(xw, sheet_name="Погашавшие", index=False)
pd.set_option("display.width", 250)
print(summary.round(3).to_string()); print(by_sector.round(3).to_string()); print(prof.round(3).to_string())
print(co.round(3).to_string())
