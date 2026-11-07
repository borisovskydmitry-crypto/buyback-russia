"""Панель «компания–год» с мультипликаторами и сравнение компаний перед погашением акций.

Запуск: python3 extract_msfo.py && python3 build_panel.py
Вход:  msfo_long.csv (строки МСФО), Анализ_акций_Мосбиржи.xlsx (число акций и капитализация
       на конец года по ~340 тикерам, файл Dmitry), ../buyback_cancellation/price_cache/IMOEX.csv.
Выход: panel.csv и pre_cancellation_results.xlsx (заголовки на русском).

Логика:
- Событие = год t, в котором число акций в обращении сократилось (по файлу Мосбиржи).
  Показатели берём на конец года t-1 (состояние ДО погашения); выкуп обычно идёт в t-1..t.
- Для каждого показателя и каждого года считаем процентильный ранг компании среди всех
  компаний того же года (0 = минимум, 1 = максимум). Ранг снимает разницу между годами
  (2022 и 2021 — разные рынки). Средний ранг погашавших = 0,5 означает «ничем не выделялись».
- Контрольная группа: все остальные компании-годы, где не было погашения в t, t+1, t-1.
  Отдельно — компании, которые выкупали акции (отток в ДДС), но не гасили их.
"""
import re
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

HERE = Path(__file__).absolute().parent
MOEX_FILE = HERE / "Анализ_акций_Мосбиржи.xlsx"
IMOEX = HERE.parent / "buyback_cancellation" / "price_cache" / "IMOEX.csv"

# Файл МСФО -> тикеры (все классы акций компании; капитализация суммируется)
FILE_TICKERS = {
    "ALROSA": ["ALRS"], "Abrau": ["ABRD"], "DOM_RF": ["DOMRF"], "EnPlus": ["ENPG"], "HeadHunter": ["HEAD"],
    "KAMAZ": ["KMAZ"], "LSR": ["LSRG"], "Mosenergo": ["MSNG"], "NMTP": ["NMTP"], "PIK": ["PIKK"],
    "RUSAL": ["RUAL"], "RusHydro": ["HYDR"], "TGK14": ["TGKN"], "TMK": ["TRMK"], "Uralkuz": ["URKZ"],
    "VK_": ["VKCO"], "VUSH": ["WUSH"], "novabev": ["BELU"], "АПРИ": ["APRI"], "АФК": ["AFKS"],
    "Акрон": ["AKRN"], "Аптечная": ["APTK"], "Аэрофлот": ["AFLT"], "Банк_Кузнецкий": ["KUZB"],
    "Банк_Санкт": ["BSPB", "BSPBP"], "Башнефть": ["BANE", "BANEP"], "ЕвроТранс": ["EUTR"],
    "Европейская_Электро": ["EELT"], "ИНАРКТИКА": ["AQUA"], "ИСКЧ": ["ABIO"], "Интер_РАО": ["IRAO"],
    "Каршеринг": ["DELI"], "КуйбышевАзот": ["KAZT", "KAZTP"], "Левенгук": ["LVHK"], "Лента": ["LENT"],
    "Лукойл.xlsx": ["LKOH"], "МГКЛ": ["MGKL"], "МД_Медикал": ["MDMG"], "МТС-Банк": ["MBNK"],
    "МТС.xlsx": ["MTSS"], "Магнит": ["MGNT"], "Московская_биржа": ["MOEX"], "Московский_кредитный": ["CBOM"],
    "НКНХ": ["NKNC", "NKNCP"], "НОВАТЭК": ["NVTK"], "Норникель": ["GMKN"], "Сегежа": ["SGZH"],
    "ИКС 5": ["X5"], "Позитив": ["POSI"], "Полюс": ["PLZL"], "Распадская": ["RASP"], "Ренессанс": ["RENI"],
    "Росинтер": ["ROST"], "Роснефть": ["ROSN"], "Россети_Волга": ["MRKV"], "Россети_Ленэнерго": ["LSNG", "LSNGP"],
    "Россети_МР": ["MSRS"], "Россети_СЗ": ["MRKZ"], "Россети_Сибирь": ["MRKS"], "Россети_ФСК": ["FEES"],
    "Россети_Центр_": ["MRKC"], "Россети_ЦиП": ["MRKP"], "Ростелеком": ["RTKM", "RTKMP"], "Русагро": ["RAGR"],
    "Северсталь": ["CHMF"], "Селигдар": ["SELG"], "Совкомбанк": ["SVCB"], "Совкомфлот": ["FLOT"],
    "Соллерс": ["SVAV"], "Софтлайн": ["SOFL"], "ТГК-1": ["TGKA"], "Татнефть": ["TATN", "TATNP"],
    "Транснефть": ["TRNFP"], "Уралкалий": ["URKA"], "Уралсиб": ["USBN"], "ФосАгро": ["PHOR"],
    "Черкизово": ["GCHE"], "ЮГК": ["UGLD"], "Яковлев": ["IRKT"],
}
BANKS = {"BSPB", "KUZB", "MBNK", "CBOM", "SVCB", "USBN", "DOMRF", "RENI", "MOEX"}  # финсектор
SECTOR = {  # укрупнённые отрасли для проверки устойчивости
    **{t: "финансы" for t in BANKS},
    **{t: "нефть и газ" for t in ["LKOH", "ROSN", "BANE", "TATN", "NVTK", "TRNFP"]},
    **{t: "металлы и добыча" for t in ["GMKN", "RASP", "PLZL", "ALRS", "CHMF", "RUAL", "ENPG", "TRMK", "SELG", "UGLD", "URKZ", "MGKL"]},
    **{t: "химия и удобрения" for t in ["AKRN", "KAZT", "PHOR", "URKA", "NKNC"]},
    **{t: "энергетика" for t in ["IRAO", "HYDR", "MSNG", "TGKA", "TGKN", "FEES", "MRKV", "LSNG", "MSRS", "MRKZ", "MRKS", "MRKC", "MRKP"]},
    **{t: "потреб. сектор" for t in ["BELU", "GCHE", "ABRD", "MGNT", "X5", "LENT", "RAGR", "APTK", "AQUA", "ROST", "MDMG", "LVHK"]},
}
# Единицы, которые в файлах не указаны или указаны неверно (проверено вручную по величинам)
UNIT_OVERRIDE = {"АФК_Система": 1.0, "Селигдар": 0.001}
USD_FILES = {"EnPlus": 1.0, "RUSAL": 1.0}  # млн долл. США
DROP = [("Магнит", [2016, 2017])]  # в файле Магнита 2016–2017 в тыс. долл., остальные годы в тыс. руб.
# Курс ЦБ (приблизительно): на конец года — для балансовых статей, средний — для потоков
USD_END = {2016: 60.66, 2017: 57.60, 2018: 69.47, 2019: 61.91, 2020: 73.88, 2021: 74.29,
           2022: 70.34, 2023: 89.69, 2024: 101.68, 2025: 78.2}
USD_AVG = {2016: 67.03, 2017: 58.35, 2018: 62.71, 2019: 64.74, 2020: 72.15, 2021: 73.65,
           2022: 68.55, 2023: 85.25, 2024: 92.57, 2025: 83.0}
STOCKS = {"equity", "total_assets", "total_liabilities", "cash", "debt_lt", "debt_st"}

METRICS = {  # код -> (русское название, чем «лучше» — для пояснения)
    "roe": "ROE (прибыль / капитал)",
    "roa": "ROA (прибыль / активы)",
    "net_margin": "Чистая маржа",
    "op_margin": "Операционная маржа",
    "pbv": "P/BV",
    "ps": "P/S",
    "ep": "E/P (доходность прибыли = 1/PE)",
    "pe": "P/E (только прибыльные)",
    "liab_assets": "Обязательства / активы",
    "cash_assets": "Денежные средства / активы",
    "cash_cap": "Денежные средства / капитализация",
    "ocf_assets": "Операционный поток / активы",
    "fcf_yield": "FCF / капитализация",
    "capex_rev": "Капзатраты / выручка",
    "div_yield": "Дивиденды выплаченные / капитализация",
    "payout": "Дивиденды / прибыль",
    "rev_growth": "Рост выручки за год",
    "ni_growth_pos": "Прибыль выросла (1/0)",
    "log_cap": "Размер: ln(капитализация)",
    "ret_vs_imoex": "Доходность акции за год минус IMOEX",
    "ret_vs_imoex_2y": "Доходность акции за 2 года минус IMOEX",
    "buyback_cap": "Выкуп в ДДС / капитализация",
    "pbv_vs_own": "P/BV к собственной медиане компании",
    "pe_vs_own": "P/E к собственной медиане компании",
    "roe_change": "Изменение ROE за год (п.п.)",
}


def file_key(name):
    for k, v in FILE_TICKERS.items():
        if k in name:
            return k, v
    return None, None


def load_msfo():
    d = pd.read_csv(HERE / "msfo_long.csv")
    d = d[d.file != "Лукойл_отчетность.xlsx"]  # квартальный дубль, используем Лукойл.xlsx
    for k, f in UNIT_OVERRIDE.items():
        m = d.file.str.contains(k)
        d.loc[m, "value_mln"] = d.loc[m, "value_raw"] * f
    for k, f in USD_FILES.items():
        m = d.file.str.contains(k)
        rate = np.where(d.loc[m, "item"].isin(STOCKS), d.loc[m, "year"].map(USD_END), d.loc[m, "year"].map(USD_AVG))
        d.loc[m, "value_mln"] = d.loc[m, "value_raw"] * f * rate
    m = d.file.str.contains("КуйбышевАзот")
    d.loc[m, "value_mln"] = d.loc[m, "value_raw"]  # первые колонки — млн руб.
    for k, yrs in DROP:
        d = d[~(d.file.str.contains(k) & d.year.isin(yrs))]
    d["key"], d["tickers"] = zip(*d.file.map(file_key))
    d = d[d.key.notna()]
    d["company"] = d.tickers.map(lambda t: t[0])
    wide = d.pivot_table(index=["company", "year"], columns="item", values="value_mln", aggfunc="first")
    return wide.reset_index(), d


def load_moex():
    df = pd.read_excel(MOEX_FILE, header=1)
    df.columns = [str(c).replace("\n", " ") for c in df.columns]
    df = df[df["Тикер"].astype(str).str.fullmatch(r"[A-Z0-9\-]+")]
    rows = []
    for _, r in df.iterrows():
        for y in range(2017, 2026):
            rows.append({"ticker": r["Тикер"], "name": r["Наименование"], "year": y,
                         "shares": r.get(f"Объём {y}, шт."), "cap_bln": r.get(f"Кап. {y}, млрд руб.")})
    m = pd.DataFrame(rows)
    m["price"] = m.cap_bln * 1e9 / m.shares
    return m


def cancellation_events(m):
    m = m.sort_values(["ticker", "year"])
    m["prev_shares"] = m.groupby("ticker").shares.shift()
    ev = m[(m.shares < m.prev_shares * 0.999)].copy()
    ev["cut_pct"] = 1 - ev.shares / ev.prev_shares
    return ev[["ticker", "name", "year", "cut_pct"]]


def pct_rank_by_year(panel, col):
    return panel.groupby("year")[col].rank(pct=True)


def main():
    wide, long = load_msfo()
    moex = load_moex()
    imoex = pd.read_csv(IMOEX, index_col=0, parse_dates=True).iloc[:, 0]
    imoex_y = imoex.groupby(imoex.index.year).last()

    # рыночные данные компании: сумма капитализаций классов акций, цена обыкновенной акции
    rows = []
    for key, tickers in FILE_TICKERS.items():
        mm = moex[moex.ticker.isin(tickers)]
        if mm.empty:
            continue
        cap = mm.groupby("year").cap_bln.sum(min_count=1) * 1000  # млн руб.
        price = mm[mm.ticker == tickers[0]].set_index("year").price
        for y in cap.index:
            rows.append({"company": tickers[0], "year": y, "cap": cap[y], "price": price.get(y)})
    mk = pd.DataFrame(rows)
    p = wide.merge(mk, on=["company", "year"], how="left").sort_values(["company", "year"])
    if "net_income_tci" in p:
        p["net_income"] = p.net_income.fillna(p.net_income_tci)
    g = p.groupby("company")
    eq_avg = (p.equity + g.equity.shift()) / 2
    p["roe"] = p.net_income / eq_avg.where(eq_avg > 0)
    p["roa"] = p.net_income / p.total_assets
    p["net_margin"] = p.net_income / p.revenue.where(p.revenue > 0)
    p["op_margin"] = p.operating_profit / p.revenue.where(p.revenue > 0)
    p["pbv"] = p.cap / p.equity.where(p.equity > 0)
    p["ps"] = p.cap / p.revenue.where(p.revenue > 0)
    p["ep"] = p.net_income / p.cap
    p["pe"] = (p.cap / p.net_income).where(p.net_income > 0)
    liab = p.total_liabilities.fillna(p.total_assets - p.equity)
    p["liab_assets"] = liab / p.total_assets
    p["cash_assets"] = p.cash / p.total_assets
    p["cash_cap"] = p.cash / p.cap
    p["ocf_assets"] = p.ocf / p.total_assets
    p["fcf_yield"] = (p.ocf - p.capex.abs().fillna(0)) / p.cap
    p["capex_rev"] = p.capex.abs() / p.revenue.where(p.revenue > 0)
    p["div_yield"] = p.dividends_paid.abs() / p.cap
    p["payout"] = (p.dividends_paid.abs() / p.net_income).where(p.net_income > 0)
    p["rev_growth"] = p.revenue / g.revenue.shift() - 1
    p["ni_growth_pos"] = (p.net_income > g.net_income.shift()).astype(float).where(g.net_income.shift().notna())
    p["log_cap"] = np.log(p.cap)
    ret = p.price / g.price.shift() - 1
    p["ret_vs_imoex"] = ret - (p.year.map(imoex_y) / (p.year - 1).map(imoex_y) - 1)
    ret2 = p.price / g.price.shift(2) - 1
    p["ret_vs_imoex_2y"] = ret2 - (p.year.map(imoex_y) / (p.year - 2).map(imoex_y) - 1)
    p["buyback_cap"] = p.buyback.abs() / p.cap
    p["roe_change"] = p.roe - p.groupby("company").roe.shift()
    # явные ошибки единиц/данных — убираем крайние значения мультипликаторов
    for c, lo, hi in [("pbv", 0, 100), ("ps", 0, 100), ("roe", -5, 5), ("net_margin", -5, 5),
                      ("op_margin", -5, 5), ("div_yield", 0, 1), ("fcf_yield", -3, 3), ("pe", 0, 500),
                      ("cash_cap", 0, 20), ("payout", 0, 10), ("rev_growth", -1, 10), ("buyback_cap", 0, 1)]:
        p.loc[(p[c] < lo) | (p[c] > hi), c] = np.nan
    # дёшево ли стоила компания относительно СВОЕЙ истории (а не рынка)
    p["pbv_vs_own"] = p.pbv / p.groupby("company").pbv.transform("median")
    p["pe_vs_own"] = p.pe / p.groupby("company").pe.transform("median")

    # события погашения (по файлу Мосбиржи) для компаний с МСФО
    ev = cancellation_events(moex)
    tick2co = {t: v[0] for v in FILE_TICKERS.values() for t in v}
    ev["company"] = ev.ticker.map(tick2co)
    ev = ev[ev.company.notna()].groupby(["company", "year"]).cut_pct.max().reset_index()
    ev = ev[ev.company != "VTBR"]
    ev_set = set(zip(ev.company, ev.year))
    p["is_bank"] = p.company.isin(BANKS)
    p["sector"] = p.company.map(SECTOR).fillna("прочие")
    # группа на конец года y: «до погашения» если погашение в y+1
    p["pre_cancel"] = [(c, y + 1) in ev_set for c, y in zip(p.company, p.year)]
    near = [any((c, y + k) in ev_set for k in (-1, 0, 1, 2)) for c, y in zip(p.company, p.year)]
    p["control"] = ~np.array(near)
    p["bought_back_no_cancel"] = p.control & (p.buyback.abs() > 0)

    metrics = list(METRICS)
    for c in metrics:
        p[f"rank_{c}"] = pct_rank_by_year(p, c)
        # ранг внутри отрасли и года (только если в группе не меньше 4 компаний)
        grp = p.groupby(["year", "sector"])[c]
        p[f"srank_{c}"] = grp.rank(pct=True).where(grp.transform("count") >= 4)
    p.to_csv(HERE / "panel.csv", index=False)

    pre = p[p.pre_cancel]
    ctrl = p[p.control]
    bb = p[p.bought_back_no_cancel]

    def compare(pre, ctrl, bb, label):
        out = []
        for c in metrics:
            x, y, z = pre[c].dropna(), ctrl[c].dropna(), bb[c].dropna()
            r = pre[f"rank_{c}"].dropna()
            if len(x) < 3:
                continue
            mw = stats.mannwhitneyu(x, y) if len(y) > 3 else None
            wr = stats.wilcoxon(r - 0.5) if len(r) >= 6 else None
            out.append({"Выборка": label, "Показатель": METRICS[c], "код": c,
                        "N погашавших": len(x), "Медиана: погашавшие (год до)": x.median(),
                        "Медиана: остальные компании": y.median(),
                        "Медиана: выкупали, но не гасили": z.median() if len(z) else np.nan,
                        "Средний ранг погашавших в своём году (0,5 = как все)": r.mean(),
                        "Доля погашавших выше медианы года": (r > 0.5).mean(),
                        "p-значение (Манна–Уитни, против остальных)": mw.pvalue if mw else np.nan,
                        "p-значение (Манна–Уитни, против выкупавших без погашения)":
                            stats.mannwhitneyu(x, z).pvalue if len(z) > 3 else np.nan,
                        "N выкупавших без погашения (компании-годы)": len(z),
                        "p-значение (ранг ≠ 0,5, Уилкоксон)": wr.pvalue if wr else np.nan})
        return out

    res = compare(pre, ctrl, bb, "все события")
    res += compare(pre[~pre.is_bank], ctrl[~ctrl.is_bank], bb[~bb.is_bank], "без банков")
    # каждая компания один раз: средние ранги по компании
    by_co = pre.groupby("company")[[f"rank_{c}" for c in metrics] + metrics].mean().reset_index()
    res += compare(by_co, ctrl, bb, "каждая компания один раз (среднее по её событиям)")
    # отраслевая поправка: ранги внутри отрасли и года вместо рангов по всему рынку
    pre_s = pre.copy()
    for c in metrics:
        pre_s[f"rank_{c}"] = pre_s[f"srank_{c}"]
    res += compare(pre_s, ctrl, bb, "ранг внутри своей отрасли и года")
    res = pd.DataFrame(res)

    # многомерная логит-модель: вероятность погашения в следующем году
    import statsmodels.api as sm
    xs = ["rank_roe", "rank_ep", "rank_cash_assets", "rank_ret_vs_imoex", "rank_log_cap"]
    data = p[p.pre_cancel | p.control][xs + ["pre_cancel"]].dropna()
    logit_rows = []
    try:
        fit = sm.Logit(data.pre_cancel.astype(float), sm.add_constant(data[xs])).fit(disp=0)
        for k in fit.params.index:
            logit_rows.append({"Переменная": "константа" if k == "const" else
                               "Ранг: " + METRICS[k.removeprefix("rank_")],
                               "Коэффициент": fit.params[k], "Ст. ошибка": fit.bse[k], "p-значение": fit.pvalues[k]})
        logit_rows.append({"Переменная": f"Наблюдений: {int(fit.nobs)}, погашений: {int(data.pre_cancel.sum())}, "
                                         f"псевдо-R² Макфаддена: {fit.prsquared:.3f}"})
    except Exception as e:
        logit_rows.append({"Переменная": f"модель не оценена: {e}"})
    logit = pd.DataFrame(logit_rows)

    # профиль каждого события
    cols = ["company", "year", "sector"] + metrics
    prof = pre[cols].copy()
    prof["year"] = prof.year + 1
    prof = prof.merge(ev, on=["company", "year"], how="left")
    rank_prof = pre[["company", "year"] + [f"rank_{c}" for c in metrics]].copy()
    rank_prof["year"] += 1

    ren = {"company": "Компания (тикер)", "year": "Год погашения", "sector": "Отрасль", "cut_pct": "Доля погашенных акций",
           **METRICS, **{f"rank_{c}": f"Ранг: {n}" for c, n in METRICS.items()}}
    coverage = pd.DataFrame({"Показатель": [METRICS[c] for c in metrics],
                             "Компаний-лет с данными": [p[c].notna().sum() for c in metrics],
                             "из них погашавших (год до)": [pre[c].notna().sum() for c in metrics]})
    out = HERE / "pre_cancellation_results.xlsx"
    with pd.ExcelWriter(out) as xw:
        res.drop(columns="код").to_excel(xw, sheet_name="Сравнение", index=False)
        prof.rename(columns=ren).to_excel(xw, sheet_name="Профиль событий", index=False)
        rank_prof.rename(columns=ren).to_excel(xw, sheet_name="Ранги событий", index=False)
        logit.to_excel(xw, sheet_name="Логит-модель", index=False)
        coverage.to_excel(xw, sheet_name="Покрытие данных", index=False)
        ev.rename(columns=ren).to_excel(xw, sheet_name="События", index=False)
    pd.set_option("display.width", 250)
    print(res[res["Выборка"] == "все события"].drop(columns=["Выборка", "Показатель"]).round(3).to_string())
    print(res[res["Выборка"] != "все события"][["Выборка", "код", "N погашавших",
          "Средний ранг погашавших в своём году (0,5 = как все)", "p-значение (ранг ≠ 0,5, Уилкоксон)"]].round(3).to_string())
    print(ev.to_string())
    print(logit.round(3).to_string())


if __name__ == "__main__":
    main()
