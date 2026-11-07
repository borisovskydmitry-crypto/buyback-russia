"""Набор «компания–год» для поиска признаков-предвестников обратного выкупа акций.

Запуск: python3 extract_extra.py && python3 build_dataset.py   (после ../pre_cancellation/build_panel.py)
Вход:  ../pre_cancellation/panel.csv — показатели МСФО и капитализация на конец года;
       msfo_extra.csv — выкуп в ДДС, казначейские акции, амортизация, долг (extract_extra.py);
       market_cache/*.csv — дневные цены и обороты Мосбиржи (fetch_market.py);
       ../pre_cancellation/Анализ_акций_Мосбиржи.xlsx — число акций (погашения), файл Dmitry.
Выход: dataset.csv — одна строка на компанию и год t: был ли выкуп в году t, признаки на конец
       года t-1 (префикс f_) и отдельно то, что происходило в самом году t (префикс c_).

Как определяется выкуп в году t:
- если в ДДС компании есть строка покупки собственных акций — сумма оттока по ней (пустой год при
  наличии ДДС = 0);
- если такой строки нет, но есть казначейские акции в балансе — их рост за год;
- если в файле нет ни того, ни другого — считаем, что выкупа не было.
Существенный выкуп — не меньше 0,1% капитализации на конец года t-1 (если её нет — года t).
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).absolute().parent
PRE = HERE.parent / "pre_cancellation"
CACHE = HERE / "market_cache"
sys.path.insert(0, str(PRE))
from build_panel import (BANKS, DROP, FILE_TICKERS, USD_AVG, USD_END, USD_FILES,  # noqa: E402
                         cancellation_events, file_key, load_moex)

MATERIAL = 0.01  # 1% капитализации (строгий вариант)
# Старые тикеры тех же бумаг (до редомициляции или смены тикера) — для ценовой истории
ALIASES = {"HEAD": "HHRU", "X5": "FIVE", "RAGR": "AGRO", "LENT": "LNTA", "ENPG": "ENPL", "VKCO": "MAIL"}
# Сколько расписок старого тикера приходится на одну акцию нового (из общих знаний; сверено со скачком цены
# там, где старый и новый ряды стыкуются). Русагро не включена: после редомициляции число акций менялось.
ALIAS_RATIO = {"HEAD": 1, "X5": 1, "ENPG": 1, "VKCO": 1, "LENT": 5}
# Из общих знаний, не из файлов (проверить перед цитированием): контроль государства
# (федерального или регионального) над компанией; год — с какого года действует
STATE = {**{t: 2000 for t in ["FEES", "MRKV", "LSNG", "MSRS", "MRKZ", "MRKS", "MRKC", "MRKP", "HYDR", "IRAO",
                              "AFLT", "ROSN", "TRNFP", "ALRS", "RTKM", "DOMRF", "FLOT", "MSNG", "TGKA", "BANE",
                              "NMTP", "KMAZ", "IRKT", "TATN"]},
         "VKCO": 2022, "UGLD": 2025}
# Из общих знаний: материнская компания была зарегистрирована за рубежом (до редомициляции в Россию)
FOREIGN = {"HEAD", "VKCO", "X5", "MDMG", "ENPG", "RUAL", "RAGR", "LENT"}
# Выкупы особого рода (не обычная программа выкупа): по пояснениям в файлах МСФО и из общих знаний
SPECIAL = {
    ("URKA", 2017): "выкуп по требованию акционеров (пояснение в файле)",
    ("URKA", 2018): "выкуп по требованию акционеров (пояснение в файле)",
    ("URKA", 2019): "выкуп по требованию акционеров (пояснение в файле)",
    ("FLOT", 2020): "стабилизация после IPO, опцион андеррайтеров (пояснение в файле)",
    ("IRAO", 2018): "покупка у ФСК и РусГидро (пояснение в файле)",
    ("IRAO", 2019): "покупка у ФСК и РусГидро (пояснение в файле)",
    ("RASP", 2021): "пут-опцион миноритариев (пояснение в файле)",
    ("EUTR", 2023): "стабилизация после IPO (общие знания)",
    ("DOMRF", 2025): "вероятно стабилизация после IPO (торги начались в ноябре 2025 г.)",
    ("FEES", 2022): "выкуп по требованию акционеров после присоединения ПАО «Россети» (общие знания)",
    ("MGNT", 2023): "выкуп у нерезидентов с дисконтом (общие знания)",
    ("X5", 2025): "выкуп у нерезидентов (общие знания)",
}


def load_extra():
    x = pd.read_csv(HERE / "msfo_extra.csv")
    x = x[x.file != "Лукойл_отчетность.xlsx"]
    x["mln"] = x.value_raw * x.unit_factor
    for k, f in USD_FILES.items():
        m = x.file.str.contains(k)
        stock = x.loc[m, "item"].isin(["treasury_bs", "borrowings"])
        rate = np.where(stock, x.loc[m, "year"].map(USD_END), x.loc[m, "year"].map(USD_AVG))
        x.loc[m, "mln"] = x.loc[m, "value_raw"] * f * rate
    m = x.file.str.contains("КуйбышевАзот")
    x.loc[m, "mln"] = x.loc[m, "value_raw"]
    for k, yrs in DROP:
        x = x[~(x.file.str.contains(k) & x.year.isin(yrs))]
    x["company"] = x.file.map(lambda f: (file_key(f)[1] or [None])[0])
    x = x[x.company.notna()]
    return x.pivot_table(index=["company", "year"], columns="item", values="mln", aggfunc="first").reset_index()


def load_px(ticker):
    f = CACHE / f"{ticker}.csv"
    if not f.exists():
        return None
    d = pd.read_csv(f, parse_dates=["date"]).set_index("date").sort_index()
    d = d[d.close > 0]
    return d if len(d) else None


def year_features(d):
    """Рыночные признаки ряда цен по календарным годам (только годы, торговавшиеся до декабря)."""
    lr = np.log(d.close).diff()
    out = {}
    for y in sorted(set(d.index.year)):
        cur = d[d.index.year == y]
        end = cur.index[-1]
        if end.month < 12 or len(cur) < 20:
            continue
        prev = d[d.index.year == y - 1]
        win = d.loc[end - pd.Timedelta(days=365):end]
        half = d.loc[end - pd.Timedelta(days=182):end]
        three = d.loc[end - pd.Timedelta(days=3 * 365):end]
        full_prev = len(prev) and prev.index[-1].month == 12
        prev2 = d[d.index.year == y - 2]
        full_prev2 = len(prev2) and prev2.index[-1].month == 12
        long_hist = three.index[0] <= end - pd.Timedelta(days=int(2.5 * 365))
        out[y] = {
            "close": cur.close.iloc[-1],
            "days": len(cur),
            "ret_1y": cur.close.iloc[-1] / prev.close.iloc[-1] - 1 if full_prev else np.nan,
            "ret_2y": cur.close.iloc[-1] / prev2.close.iloc[-1] - 1 if full_prev2 else np.nan,
            "ret_6m": half.close.iloc[-1] / half.close.iloc[0] - 1 if half.index[0] <= end - pd.Timedelta(days=170) else np.nan,
            "dd_52w": cur.close.iloc[-1] / win.close.max() - 1,
            "mdd_year": (cur.close / cur.close.cummax() - 1).min(),
            "vol": lr.loc[cur.index].std() * np.sqrt(252) if len(cur) >= 60 else np.nan,
            "px_vs_3y": cur.close.iloc[-1] / three.close.mean() - 1 if long_hist else np.nan,
            "dd_3y": cur.close.iloc[-1] / three.close.max() - 1 if long_hist else np.nan,
            "value": cur.value.sum(),
        }
    return pd.DataFrame(out).T


def market_panel(moex):
    """Признаки по компаниям: основной тикер (при пропуске — старый тикер), обороты — по всем классам."""
    idx = load_px("IMOEX")
    imo = year_features(idx)
    rows = []
    for key, tickers in FILE_TICKERS.items():
        main = tickers[0]
        series = {t: load_px(t) for t in tickers + ([ALIASES[main]] if main in ALIASES else [])}
        series = {t: s for t, s in series.items() if s is not None}
        if main not in series and ALIASES.get(main) not in series:
            continue
        feats = {t: year_features(s) for t, s in series.items()}
        first = min(s.index.min() for s in series.values())
        years = sorted(set().union(*[set(f.index) for f in feats.values()]))
        for y in years:
            src = next((t for t in [main, ALIASES.get(main)] if t in feats and y in feats[t].index
                        and feats[t].loc[y, "days"] >= 100), None)
            if src is None:
                continue
            r = feats[src].loc[y].to_dict()
            r["value"] = sum(f.loc[y, "value"] for f in feats.values() if y in f.index)
            rows.append({"company": main, "year": y, "px_src": src, "first_trade": first, **r})
    mk = pd.DataFrame(rows)
    mk = mk.merge(imo[["ret_1y", "ret_2y", "ret_6m", "dd_52w", "mdd_year"]].add_prefix("imoex_"),
                  left_on="year", right_index=True, how="left")
    mk["ret_vs_imoex_mkt"] = mk.ret_1y - mk.imoex_ret_1y
    mk["ret2_vs_imoex"] = mk.ret_2y - mk.imoex_ret_2y
    mk["ret6_vs_imoex"] = mk.ret_6m - mk.imoex_ret_6m
    mk["dd_rel"] = mk.dd_52w - mk.imoex_dd_52w
    mk["mdd_rel"] = mk.mdd_year - mk.imoex_mdd_year
    mk["listing_age"] = (pd.to_datetime(mk.year.astype(int).astype(str) + "-12-31") - mk.first_trade).dt.days / 365.25
    # ряды начинаются с 2015 г.: «на бирже с 2015 г. или раньше» — возраст не определён точнее
    mk.loc[mk.first_trade <= "2015-01-31", "listing_age"] = np.nan
    mk["listed_pre2015"] = mk.first_trade <= "2015-01-31"
    return mk


def main():
    p = pd.read_csv(PRE / "panel.csv")
    base_cols = ["company", "year", "revenue", "net_income", "equity", "total_assets", "total_liabilities", "cash",
                 "ocf", "capex", "dividends_paid", "operating_profit", "cap", "sector", "is_bank"]
    p = p[base_cols]
    moex = load_moex()
    mk = market_panel(moex)
    x = load_extra()
    d = p.merge(x, on=["company", "year"], how="left").merge(mk, on=["company", "year"], how="left")
    d = d.sort_values(["company", "year"]).reset_index(drop=True)
    # Капитализации нет в файле (2016 г., годы под старым тикером): берём ближайший год из файла и
    # переносим её по цене ISS (цены ISS скорректированы на сплиты, число акций в файле — нет,
    # поэтому «цена × число акций» дало бы ошибку, например, у Полюса после сплита 2025 г.)
    d["cap_est"] = np.nan
    for c, g in d.groupby("company"):
        known = g[g.cap.notna() & g.close.notna()]
        for i, r in g[g.cap.isna() & g.close.notna()].iterrows():
            if known.empty:
                continue
            j = known.index[np.argmin(np.abs(known.year - r.year))]
            ratio = 1 if known.loc[j, "px_src"] == r.px_src else ALIAS_RATIO.get(c)
            if ratio:
                d.loc[i, "cap_est"] = known.loc[j, "cap"] * ratio * r.close / known.loc[j, "close"]
    d["cap_src"] = np.where(d.cap.notna(), "файл Мосбиржи", np.where(d.cap_est.notna(), "оценка по цене ISS", ""))
    d["cap"] = d.cap.fillna(d.cap_est)

    # ---- выкуп в году ----
    cf_cos = set(x.loc[x.buyback_cf.notna(), "company"])
    tre_cos = set(x.loc[x.treasury_bs.notna(), "company"])
    d["tre"] = (-d.treasury_bs).where(d.treasury_bs <= 0)  # казначейские акции, положительная величина
    g = d.groupby("company")
    consecutive = g.year.shift() == d.year - 1
    d["tre_inc"] = (d.tre - g.tre.shift()).where(consecutive)
    has_cf = d.ocf.notna() | d.buyback_cf.notna()
    d["bb_amt"], d["bb_src"] = np.nan, ""
    m1 = d.company.isin(cf_cos) & has_cf
    d.loc[m1, "bb_amt"] = (-d.loc[m1, "buyback_cf"]).clip(lower=0).fillna(0)
    d.loc[m1, "bb_src"] = "ДДС"
    m2 = ~m1 & d.company.isin(tre_cos) & d.tre_inc.notna()
    d.loc[m2, "bb_amt"] = d.loc[m2, "tre_inc"].clip(lower=0)
    d.loc[m2, "bb_src"] = "рост казначейских акций"
    m3 = ~m1 & ~m2 & ~d.company.isin(cf_cos | tre_cos) & has_cf
    d.loc[m3, "bb_amt"] = 0.0
    d.loc[m3, "bb_src"] = "строк выкупа в отчётности нет"
    cap_prev = g.cap.shift().where(consecutive)
    d["cap_ref"] = cap_prev.fillna(d.cap)
    d["bb_share"] = d.bb_amt / d.cap_ref
    # казначейских акций стало заметно меньше (погашение, продажа, передача сотрудникам) — покупки
    # за этот год по балансу не видны, год считаем неизвестным, а не «без выкупа»
    hidden = m2 & (d.tre_inc < -MATERIAL * d.cap_ref)
    d.loc[hidden, ["bb_amt", "bb_share"]] = np.nan
    d.loc[hidden, "bb_src"] = "казначейских акций стало меньше — покупки не видны"
    d["bb_type"] = [SPECIAL.get((c, y), "") for c, y in zip(d.company, d.year)]
    d["bb"] = (d.bb_share >= MATERIAL).astype(float).where(d.bb_share.notna())
    d["bb_any"] = (d.bb_amt > 0).astype(float).where(d.bb_amt.notna())
    for thr, name in [(0.005, "bb_05"), (0.01, "bb_1")]:
        d[name] = (d.bb_share >= thr).astype(float).where(d.bb_share.notna())
    ev = cancellation_events(moex)
    tick2co = {t: v[0] for v in FILE_TICKERS.values() for t in v}
    ev["company"] = ev.ticker.map(tick2co)
    ev = ev[ev.company.notna()].groupby(["company", "year"]).cut_pct.max()
    d["cancel"] = [float((c, y) in ev.index) for c, y in zip(d.company, d.year)]
    d["cancel_pct"] = [ev.get((c, y), 0.0) for c, y in zip(d.company, d.year)]

    # ---- показатели на конец года ----
    g = d.groupby("company")
    lag = lambda col: g[col].shift().where(consecutive)  # noqa: E731
    eq_avg = (d.equity + lag("equity")) / 2
    pos_rev = d.revenue.where(d.revenue > 0)
    d["roe"] = d.net_income / eq_avg.where(eq_avg > 0)
    d["roa"] = d.net_income / d.total_assets
    d["net_margin"] = d.net_income / pos_rev
    d["op_margin"] = d.operating_profit / pos_rev
    da_ok = d.da.where((d.da > 0) & (d.da < 0.5 * d.revenue.abs()))
    d["ebitda"] = d.operating_profit + da_ok
    d["ebitda_margin"] = d.ebitda / pos_rev
    d["roe_change"] = d.roe - lag("roe")
    d["ni_growth_pos"] = (d.net_income > lag("net_income")).astype(float).where(lag("net_income").notna())
    d["rev_growth"] = d.revenue / lag("revenue") - 1
    d["eq_growth"] = (d.equity / lag("equity") - 1).where(lag("equity") > 0)
    d["ni_growth"] = (d.net_income - lag("net_income")) / d.total_assets
    d["loss"] = (d.net_income < 0).astype(float).where(d.net_income.notna())
    d["cash_assets"] = d.cash / d.total_assets
    d["cash_cap"] = d.cash / d.cap
    d["cash_growth"] = (d.cash - lag("cash")) / d.total_assets
    debt = d.borrowings.where(~d.is_bank)
    d["debt_assets"] = debt / d.total_assets
    d["net_cash"] = (d.cash > debt).astype(float).where(debt.notna() & d.cash.notna())
    nd = debt - d.cash
    d["nd_ebitda"] = (nd / d.ebitda).where(d.ebitda > 0)
    d["ev_ebitda"] = ((d.cap + nd) / d.ebitda).where(d.ebitda > 0)
    liab = d.total_liabilities.fillna(d.total_assets - d.equity)
    d["liab_assets"] = liab / d.total_assets
    d["ocf_assets"] = d.ocf / d.total_assets
    d["fcf_yield"] = (d.ocf - d.capex.abs().fillna(0)) / d.cap
    d["capex_rev"] = d.capex.abs() / pos_rev
    d["ocf_ni"] = (d.ocf / d.net_income).where(d.net_income > 0)
    d["div_yield"] = d.dividends_paid.abs() / d.cap
    d["payout"] = (d.dividends_paid.abs() / d.net_income).where(d.net_income > 0)
    prev_div = lag("dividends_paid").abs()
    d["div_cut"] = (d.dividends_paid.abs().fillna(0) < 0.5 * prev_div).astype(float).where(prev_div > 0)
    d["no_div"] = (d.dividends_paid.abs().fillna(0) == 0).astype(float).where(d.ocf.notna())
    d["pbv"] = d.cap / d.equity.where(d.equity > 0)
    d["ps"] = d.cap / pos_rev
    d["ep"] = d.net_income / d.cap
    pe = (d.cap / d.net_income).where(d.net_income > 0)
    d["log_cap"] = np.log(d.cap)
    d["turnover"] = d.value / (d.cap * 1e6)
    d["ret_vs_imoex"] = d.ret_vs_imoex_mkt
    d["treasury_cap"] = (d.tre / d.cap).fillna(0).where(d.company.isin(tre_cos) | d.bb_amt.notna())
    d["has_pref"] = d.company.map({v[0]: float(len(v) > 1) for v in FILE_TICKERS.values()})
    d["state"] = [float(c in STATE and y >= STATE[c]) for c, y in zip(d.company, d.year)]
    d["foreign"] = d.company.isin(FOREIGN).astype(float)
    d["bank"] = d.is_bank.astype(float)
    for c, lo, hi in [("pbv", 0, 100), ("ps", 0, 100), ("roe", -5, 5), ("net_margin", -5, 5), ("op_margin", -5, 5),
                      ("ebitda_margin", -5, 5), ("div_yield", 0, 1), ("fcf_yield", -3, 3), ("cash_cap", 0, 20),
                      ("payout", 0, 10), ("rev_growth", -1, 10), ("nd_ebitda", -50, 50), ("ev_ebitda", -50, 200),
                      ("ocf_ni", -20, 20), ("turnover", 0, 50), ("treasury_cap", 0, 1), ("eq_growth", -1, 10),
                      ("ni_growth", -2, 2)]:
        d.loc[(d[c] < lo) | (d[c] > hi), c] = np.nan
    pe = pe.where((pe > 0) & (pe < 500))
    d["pbv_vs_own"] = d.pbv / d.groupby("company").pbv.transform("median")
    d["pe_vs_own"] = pe / pe.groupby(d.company).transform("median")
    # то же без заглядывания вперёд: к медиане трёх ПРЕДЫДУЩИХ лет (нужно не меньше двух)
    past_med = lambda v: v.groupby(d.company).transform(  # noqa: E731
        lambda s: s.shift().rolling(3, min_periods=2).median())
    d["pbv_vs_past"] = d.pbv / past_med(d.pbv)
    d["pe_vs_past"] = pe / past_med(pe)

    # ---- история выкупов (по годам, известным до t) ----
    g = d.groupby("company")
    d["bb_prev"] = g.bb.shift().where(consecutive)
    seen = g.bb.transform(lambda s: s.fillna(0).cumsum().shift())
    obs = g.bb.transform(lambda s: s.notna().cumsum().shift())
    d["bb_ever"] = (seen > 0).astype(float).where(obs > 0)
    d["cancel_prev"] = g.cancel.shift().where(consecutive)

    pre_feats = ["roe", "roa", "net_margin", "op_margin", "ebitda_margin", "roe_change", "ni_growth_pos", "rev_growth",
                 "eq_growth", "ni_growth",
                 "loss", "cash_assets", "cash_cap", "cash_growth", "debt_assets", "net_cash", "nd_ebitda",
                 "liab_assets", "ocf_assets", "fcf_yield", "capex_rev", "ocf_ni", "div_yield", "payout", "div_cut",
                 "no_div", "pbv", "ps", "ep", "ev_ebitda", "pbv_vs_own", "pe_vs_own", "pbv_vs_past", "pe_vs_past", "log_cap",
                 "ret_vs_imoex", "ret2_vs_imoex", "ret6_vs_imoex", "dd_52w", "dd_rel", "dd_3y", "mdd_year", "mdd_rel",
                 "vol", "turnover", "px_vs_3y",
                 "listing_age", "treasury_cap", "bb_prev", "bb_ever", "cancel_prev"]
    same_year = ["ret_vs_imoex", "dd_52w", "dd_rel", "mdd_year", "mdd_rel", "imoex_ret_1y", "imoex_mdd_year"]
    lagged = d[["company", "year"] + pre_feats].copy()
    lagged["year"] += 1
    lagged = lagged.rename(columns={c: f"f_{c}" for c in pre_feats})
    # bb_prev/bb_ever/cancel_prev уже относятся к году до t — берём их без сдвига
    hist = d[["company", "year", "bb_prev", "bb_ever", "cancel_prev"]].rename(
        columns={"bb_prev": "f_bb_prev", "bb_ever": "f_bb_ever", "cancel_prev": "f_cancel_prev"})
    lagged = lagged.drop(columns=["f_bb_prev", "f_bb_ever", "f_cancel_prev"])
    out = d[["company", "year", "sector", "bank", "state", "foreign", "has_pref", "cap", "cap_ref", "cap_src",
             "bb_amt", "bb_share", "bb_src", "bb_type", "bb", "bb_any", "bb_05", "bb_1", "tre", "tre_inc", "cancel", "cancel_pct",
             "first_trade", "listed_pre2015", "px_src"] + same_year].copy()
    out = out.rename(columns={c: f"c_{c}" for c in same_year})
    out = out.merge(lagged, on=["company", "year"], how="left").merge(hist, on=["company", "year"], how="left")
    first = pd.to_datetime(out.first_trade)
    new = first > "2015-01-31"  # начало торгов видно в данных (ряды ISS загружены с 2015 г.)
    out["c_ipo_year"] = (new & (first.dt.year == out.year)).astype(float).where(first.notna())
    out["f_recent_ipo"] = (new & (out.year - first.dt.year).between(1, 2)).astype(float).where(first.notna())
    out = out[(out.year >= 2017) & (out.year <= 2025)]
    out.to_csv(HERE / "dataset.csv", index=False)
    d.to_csv(HERE / "panel_full.csv", index=False)

    s = out[out.bb.notna()]
    print("строк с известным выкупом:", len(s), "компаний:", s.company.nunique(),
          "выкуп ≥0,1%:", int(s.bb.sum()), "любой выкуп:", int(s.bb_any.sum()))
    print(s.groupby("year").agg(n=("bb", "size"), bb=("bb", "sum"), share=("bb", "mean")).round(2).T.to_string())
    print(s.bb_src.value_counts().to_string())
    print("нет данных о выкупе:", out.bb.isna().sum(), sorted(out[out.bb.isna()].company.unique()))


if __name__ == "__main__":
    main()
