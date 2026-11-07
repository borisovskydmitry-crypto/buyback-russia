"""Поиск признаков-предвестников обратного выкупа: расчёты и Excel с русскими заголовками.

Запуск: python3 analyze.py  (после build_dataset.py)
Вход:  dataset.csv. Выход: precursors_results.xlsx и summary.json (цифры для отчёта).

Наблюдение — компания и год t. «Выкуп» — выкуп в году t не меньше 0,1% капитализации.
Признаки с префиксом f_ измерены на конец года t-1, то есть до года выкупа (отчётность за t-1
выходит в феврале–апреле года t). Признаки с префиксом c_ относятся к самому году выкупа.

1. Каждый признак отдельно. Значение переводится в процентильный ранг среди компаний того же года
   (так 2022 г. не смешивается с 2021 г.). AUC — вероятность того, что у случайного года с выкупом
   ранг выше, чем у случайного года без выкупа (0,5 — признак ничего не различает, 1 — различает
   идеально, меньше 0,5 — у выкупающих признак ниже). Бинарная версия — верхняя или нижняя треть
   года (или сам флаг): как часто встречается у выкупавших и у не выкупавших, во сколько раз чаще.
   Интервалы — бутстреп по компаниям (компания целиком попадает или не попадает в выборку),
   p-значения — логит с ошибками, кластеризованными по компаниям, и поправка Бенджамини–Хохберга.
2. «Какие компании» против «когда»: условный логит (сравниваются только годы одной компании) и
   корреляция средних по компаниям.
3. Модели из нескольких признаков: проверка на компаниях, не участвовавших в оценке, и на будущих
   годах (оценка на 2017–2021, проверка на 2022–2025).
4. Правила из 1–3 условий: отбор на 2017–2021, проверка на 2022–2025.
5. Устойчивость к определению выкупа и составу выборки.
"""
import itertools
import json
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import GroupKFold
from statsmodels.discrete.conditional_models import ConditionalLogit
from statsmodels.stats.multitest import multipletests

warnings.filterwarnings("ignore")
HERE = Path(__file__).absolute().parent
RNG = np.random.default_rng(20261024)
B = 1000  # повторов бутстрепа

FEATS = {  # код -> (название, группа)
    "f_bb_prev": ("Выкуп в прошлом году", "История выкупов"),
    "f_bb_ever": ("Выкупала раньше (любой год с 2016 г.)", "История выкупов"),
    "f_treasury_cap": ("Казначейские акции на балансе / капитализация", "История выкупов"),
    "f_cancel_prev": ("Погашение акций в прошлом году", "История выкупов"),
    "state": ("Госкомпания (из общих знаний)", "Собственник и устройство"),
    "foreign": ("Иностранный холдинг до редомициляции (из общих знаний)", "Собственник и устройство"),
    "bank": ("Финансовый сектор", "Собственник и устройство"),
    "has_pref": ("Есть привилегированные акции", "Собственник и устройство"),
    "f_recent_ipo": ("IPO 1–2 года назад", "Собственник и устройство"),
    "f_log_cap": ("Размер (капитализация)", "Собственник и устройство"),
    "f_roe": ("ROE", "Прибыльность и рост"),
    "f_roa": ("ROA", "Прибыльность и рост"),
    "f_net_margin": ("Чистая маржа", "Прибыльность и рост"),
    "f_op_margin": ("Операционная маржа", "Прибыльность и рост"),
    "f_ebitda_margin": ("Маржа EBITDA", "Прибыльность и рост"),
    "f_roe_change": ("Изменение ROE за год", "Прибыльность и рост"),
    "f_ni_growth_pos": ("Прибыль выросла за год", "Прибыльность и рост"),
    "f_loss": ("Убыток", "Прибыльность и рост"),
    "f_rev_growth": ("Рост выручки", "Прибыльность и рост"),
    "f_ni_growth": ("Изменение прибыли за год / активы", "Прибыльность и рост"),
    "f_eq_growth": ("Рост собственного капитала за год", "Деньги и долг"),
    "f_cash_assets": ("Денежные средства / активы", "Деньги и долг"),
    "f_cash_cap": ("Денежные средства / капитализация", "Деньги и долг"),
    "f_cash_growth": ("Прирост денежных средств за год / активы", "Деньги и долг"),
    "f_debt_assets": ("Кредиты и займы / активы", "Деньги и долг"),
    "f_net_cash": ("Денег больше, чем долга", "Деньги и долг"),
    "f_nd_ebitda": ("Чистый долг / EBITDA", "Деньги и долг"),
    "f_liab_assets": ("Обязательства / активы", "Деньги и долг"),
    "f_ocf_assets": ("Операционный поток / активы", "Денежный поток и инвестиции"),
    "f_fcf_yield": ("Свободный поток / капитализация", "Денежный поток и инвестиции"),
    "f_capex_rev": ("Капзатраты / выручка", "Денежный поток и инвестиции"),
    "f_ocf_ni": ("Операционный поток / прибыль", "Денежный поток и инвестиции"),
    "f_div_yield": ("Дивидендная доходность (выплаченные)", "Дивиденды"),
    "f_payout": ("Доля прибыли на дивиденды", "Дивиденды"),
    "f_div_cut": ("Дивиденды сокращены вдвое и больше", "Дивиденды"),
    "f_no_div": ("Дивиденды не платились", "Дивиденды"),
    "f_pbv": ("P/BV", "Оценка акций"),
    "f_ps": ("P/S", "Оценка акций"),
    "f_ep": ("E/P (обратный P/E)", "Оценка акций"),
    "f_ev_ebitda": ("EV/EBITDA", "Оценка акций"),
    "f_pbv_vs_past": ("P/BV к медиане трёх прошлых лет", "Оценка акций"),
    "f_pe_vs_past": ("P/E к медиане трёх прошлых лет", "Оценка акций"),
    "f_ret_vs_imoex": ("Доходность акции за год минус IMOEX", "Цена и торговля"),
    "f_ret2_vs_imoex": ("Доходность за 2 года минус IMOEX", "Цена и торговля"),
    "f_ret6_vs_imoex": ("Доходность за последние 6 месяцев минус IMOEX", "Цена и торговля"),
    "f_dd_52w": ("Цена к максимуму за 52 недели", "Цена и торговля"),
    "f_dd_3y": ("Цена к максимуму за 3 года", "Цена и торговля"),
    "f_px_vs_3y": ("Цена к средней за 3 года", "Цена и торговля"),
    "f_mdd_year": ("Максимальная просадка за год", "Цена и торговля"),
    "f_vol": ("Волатильность", "Цена и торговля"),
    "f_turnover": ("Оборот за год / капитализация", "Цена и торговля"),
}
CONCURRENT = {
    "c_ret_vs_imoex": ("Доходность акции минус IMOEX в год выкупа", "В год выкупа"),
    "c_dd_52w": ("Цена к максимуму за 52 недели на конец года выкупа", "В год выкупа"),
    "c_mdd_year": ("Максимальная просадка в год выкупа", "В год выкупа"),
    "c_ipo_year": ("IPO в год выкупа", "В год выкупа"),
}
BINARY = {"f_bb_prev", "f_bb_ever", "f_cancel_prev", "state", "foreign", "bank", "has_pref", "f_recent_ipo",
          "f_ni_growth_pos", "f_loss", "f_net_cash", "f_div_cut", "f_no_div", "c_ipo_year"}
NAMES = {k: v[0] for k, v in {**FEATS, **CONCURRENT}.items()}
HISTORY = {"f_bb_prev", "f_bb_ever", "f_treasury_cap", "f_cancel_prev"}


def auc(y, s):
    """AUC через ранги (доля пар «выкуп — без выкупа», где у выкупа значение выше; ничьи — пополам)."""
    y = np.asarray(y, bool)
    n1, n0 = y.sum(), (~y).sum()
    if n1 == 0 or n0 == 0:
        return np.nan
    r = stats.rankdata(s)
    return (r[y].sum() - n1 * (n1 + 1) / 2) / (n1 * n0)


def score(x, c):
    """Ранг внутри года (для бинарных — сам флаг)."""
    return x[c].astype(float) if c in BINARY else x.groupby("year")[c].rank(pct=True)


def cluster_logit_p(y, s, groups):
    try:
        m = sm.Logit(y, sm.add_constant(np.asarray(s, float))).fit(
            disp=0, cov_type="cluster", cov_kwds={"groups": pd.factorize(groups)[0]})
        return float(m.pvalues[1])
    except Exception:  # noqa: BLE001
        return np.nan


def boot_indices(companies):
    """Индексы бутстреп-выборок по компаниям."""
    codes, uniq = pd.factorize(companies)
    members = [np.flatnonzero(codes == k) for k in range(len(uniq))]
    for _ in range(B):
        pick = RNG.integers(0, len(uniq), len(uniq))
        yield np.concatenate([members[k] for k in pick])


def describe_feature(df, c, target="bb", boot=True):
    x = df[[c, target, "year", "company"]].dropna()
    if x[c].nunique() < 2 or x[target].sum() < 5 or (x[target] == 0).sum() < 5:
        return None
    s = score(x, c).values
    y = x[target].values.astype(bool)
    a = auc(y, s)
    binary = c in BINARY
    if binary:
        flag = s == 1 if a >= 0.5 else s == 0
        side = "да" if a >= 0.5 else "нет"
    else:
        flag = s >= 2 / 3 if a >= 0.5 else s <= 1 / 3
        side = "верхняя треть года" if a >= 0.5 else "нижняя треть года"
    sh1, sh0 = flag[y].mean(), flag[~y].mean()
    p_with, p_without = y[flag].mean() if flag.any() else np.nan, y[~flag].mean() if (~flag).any() else np.nan
    row = {"код": c, "Признак": NAMES[c], "Группа": {**FEATS, **CONCURRENT}[c][1],
           "Лет с выкупом (есть данные)": int(y.sum()), "Лет без выкупа (есть данные)": int((~y).sum()),
           "Компаний": x.company.nunique(),
           "Медиана: годы с выкупом": x.loc[y, c].median(), "Медиана: годы без выкупа": x.loc[~y, c].median(),
           "AUC": a, "Вариант признака": side,
           "Доля среди лет с выкупом": sh1, "Доля среди лет без выкупа": sh0,
           "Во сколько раз чаще у выкупавших": sh1 / sh0 if sh0 > 0 else np.nan,
           "Вероятность выкупа, если признак есть": p_with, "Вероятность выкупа, если признака нет": p_without,
           "p-значение (логит, кластеры по компаниям)": cluster_logit_p(y.astype(float), s, x.company.values)}
    if boot:
        aucs, ratios = [], []
        for idx in boot_indices(x.company.values):
            yb, sb, fb = y[idx], s[idx], flag[idx]
            aucs.append(auc(yb, sb))
            r0 = fb[~yb].mean() if (~yb).any() else np.nan
            ratios.append(fb[yb].mean() / r0 if r0 > 0 else np.nan)
        row["AUC: 95% интервал, нижняя граница"], row["AUC: 95% интервал, верхняя граница"] = np.nanpercentile(aucs, [2.5, 97.5])
        row["Во сколько раз чаще: 95% интервал, нижняя"], row["Во сколько раз чаще: 95% интервал, верхняя"] = \
            np.nanpercentile(ratios, [2.5, 97.5])
    return row


def univariate(df, feats, target="bb", boot=True):
    rows = [r for c in feats if (r := describe_feature(df, c, target, boot)) is not None]
    out = pd.DataFrame(rows)
    ok = out["p-значение (логит, кластеры по компаниям)"].notna()
    out.loc[ok, "q-значение (поправка на множественные сравнения)"] = multipletests(
        out.loc[ok, "p-значение (логит, кластеры по компаниям)"], method="fdr_bh")[1]
    out["Сила различия |AUC−0,5|"] = (out.AUC - 0.5).abs()
    return out.sort_values("Сила различия |AUC−0,5|", ascending=False)


def within_between(df, feats, target="bb"):
    rows = []
    for c in feats:
        x = df[[c, target, "year", "company"]].dropna()
        if x[c].nunique() < 2:
            continue
        x = x.assign(s=score(x, c))
        var = x.groupby("company")[target].transform("nunique") > 1
        xx = x[var]
        row = {"код": c, "Признак": NAMES[c], "Компаний, где были годы и с выкупом, и без": xx.company.nunique(),
               "Лет у этих компаний": len(xx)}
        if xx.company.nunique() >= 5 and xx.s.nunique() > 1:
            try:
                m = ConditionalLogit(xx[target], xx[["s"]], groups=xx.company).fit(disp=0)
                row["Внутри компании: коэффициент"] = m.params["s"]
                row["Внутри компании: отношение шансов (низ → верх года)"] = np.exp(m.params["s"])
                row["Внутри компании: p-значение"] = m.pvalues["s"]
            except Exception:  # noqa: BLE001
                pass
            diff = xx.groupby("company").apply(lambda g: g.loc[g[target] == 1, "s"].mean() - g.loc[g[target] == 0, "s"].mean())
            diff = diff.dropna()
            row["Внутри компании: средняя разница ранга (год выкупа − год без)"] = diff.mean()
            row["Доля компаний, у которых в годы выкупа выше"] = (diff > 0).mean()
            if (diff != 0).sum() >= 5:
                row["p-значение (знаковый тест Уилкоксона по компаниям)"] = stats.wilcoxon(diff).pvalue
        co = x.groupby("company").agg(s=("s", "mean"), f=(target, "mean"), n=("s", "size"))
        co = co[co.n >= 3]
        if len(co) >= 10:
            rho, p = stats.spearmanr(co.s, co.f)
            row["Между компаниями: корреляция среднего значения с частотой выкупов"] = rho
            row["Между компаниями: p-значение"] = p
            row["Компаний (не меньше 3 лет данных)"] = len(co)
        rows.append(row)
    return pd.DataFrame(rows)


def design(df, cols):
    """Матрица ранговых признаков; пропуск = середина распределения (0,5) или 0 для флагов."""
    X = pd.DataFrame(index=df.index)
    for c in cols:
        s = score(df, c)
        X[c] = s.fillna(0.0 if c in BINARY else 0.5)
    return X


def fit_predict(Xtr, ytr, Xte):
    m = LogisticRegression(C=1.0, max_iter=2000)
    m.fit(Xtr, ytr)
    return m.predict_proba(Xte)[:, 1]


def evaluate_models(df, models, target="bb"):
    rows = []
    y = df[target].values.astype(int)
    groups = df.company.values
    for name, cols in models.items():
        X = design(df, cols)
        p_in = fit_predict(X, y, X)
        cv = []
        for rep in range(20):  # повторы с разным разбиением компаний на 5 групп
            perm = {c: RNG.random() for c in np.unique(groups)}
            g = np.array([perm[c] for c in groups])
            pred = np.zeros(len(y))
            for tr, te in GroupKFold(5).split(X, y, pd.qcut(g, 5, labels=False)):
                pred[te] = fit_predict(X.iloc[tr], y[tr], X.iloc[te])
            cv.append(auc(y, pred))
        tr, te = df.year <= 2021, df.year >= 2022
        pt = fit_predict(X[tr], y[tr], X[te])
        yt = y[te.values]
        top = pt >= np.quantile(pt, 0.8)
        rows.append({"Модель": name, "Признаки": ", ".join(NAMES[c] for c in cols),
                     "AUC на тех же данных": auc(y, p_in),
                     "AUC на новых компаниях (среднее 20 разбиений)": np.mean(cv),
                     "AUC на новых компаниях: разброс (ст. откл.)": np.std(cv),
                     "AUC на будущих годах (обучение 2017–2021, проверка 2022–2025)": auc(yt, pt),
                     "Проверка 2022–2025: доля лет с выкупом всего": yt.mean(),
                     "Проверка 2022–2025: доля лет с выкупом среди 20% с наибольшей оценкой": yt[top].mean(),
                     "Во сколько раз выше среднего (топ-20%)": yt[top].mean() / yt.mean()})
    return pd.DataFrame(rows)


def coefficients(df, cols, target="bb"):
    X = sm.add_constant(design(df, cols))
    m = sm.Logit(df[target].astype(float), X).fit(disp=0, cov_type="cluster",
                                                     cov_kwds={"groups": pd.factorize(df.company)[0]})
    out = pd.DataFrame({"Признак": ["константа"] + [NAMES[c] for c in cols], "Коэффициент": m.params.values,
                        "Отношение шансов (низ → верх года / нет → да)": np.exp(m.params.values),
                        "Ст. ошибка (кластеры по компаниям)": m.bse.values, "p-значение": m.pvalues.values})
    return out, m


def candidate_conditions(df, feats, train):
    """Условия «верхняя/нижняя треть года» или флаг; направление выбирается только по обучающим годам."""
    conds = {}
    for c in feats:
        x = df[[c, "bb", "year"]].dropna()
        s = score(x, c).reindex(df.index)
        tr = train & s.notna()
        if tr.sum() < 30:
            continue
        a = auc(df.loc[tr, "bb"].values, s[tr].values)
        if c in BINARY:
            val = 1.0 if a >= 0.5 else 0.0
            conds[f"{NAMES[c]}: {'да' if val else 'нет'}"] = (c, (s == val).values)
        else:
            if a >= 0.5:
                conds[f"{NAMES[c]}: верхняя треть года"] = (c, (s >= 2 / 3).values)
            else:
                conds[f"{NAMES[c]}: нижняя треть года"] = (c, (s <= 1 / 3).values)
    return conds


def rule_stats(flag, y):
    n = flag.sum()
    prec = y[flag].mean() if n else np.nan
    rec = flag[y].mean()
    fpr = flag[~y].mean()
    return n, prec, rec, fpr


def wilson_low(k, n, z=1.96):
    if n == 0:
        return 0.0
    p = k / n
    return (p + z * z / (2 * n) - z * np.sqrt(p * (1 - p) / n + z * z / (4 * n * n))) / (1 + z * z / n)


def search_rules(df, feats, exclude_history=False, max_len=3, min_support=15):
    train = (df.year <= 2021).values
    test = ~train
    y = df.bb.values.astype(bool)
    conds = candidate_conditions(df, [f for f in feats if not (exclude_history and f in HISTORY)], pd.Series(train, df.index))
    keys = list(conds)
    base_tr, base_te = y[train].mean(), y[test].mean()
    found = []
    for k in range(1, max_len + 1):
        for combo in itertools.combinations(keys, k):
            if len({conds[c][0] for c in combo}) < k:
                continue
            flag = np.logical_and.reduce([conds[c][1] for c in combo])
            n, prec, rec, fpr = rule_stats(flag[train], y[train])
            if n < min_support:
                continue
            found.append({"Правило": " И ".join(combo), "Условий": k, "n": n, "prec": prec, "rec": rec, "fpr": fpr,
                          "low": wilson_low(prec * n, n), "flag": flag})
    found = sorted(found, key=lambda r: -r["low"])
    rows = []
    for r in found[:15]:
        nte, pte, rte, fte = rule_stats(r["flag"][test], y[test])
        rows.append({"Правило": r["Правило"], "Условий": r["Условий"],
                     "2017–2021: лет под правилом": r["n"], "2017–2021: доля с выкупом": r["prec"],
                     "2017–2021: во сколько раз выше средней доли": r["prec"] / base_tr,
                     "2017–2021: охват лет с выкупом": r["rec"], "2017–2021: доля лет без выкупа под правилом": r["fpr"],
                     "2022–2025: лет под правилом": nte, "2022–2025: доля с выкупом": pte,
                     "2022–2025: во сколько раз выше средней доли": pte / base_te if nte else np.nan,
                     "2022–2025: охват лет с выкупом": rte, "2022–2025: доля лет без выкупа под правилом": fte})
    return pd.DataFrame(rows), base_tr, base_te


def robustness(d, feats):
    """AUC ключевых признаков при разных определениях выкупа и составе выборки."""
    variants = {
        "Основной вариант (выкуп от 0,1% капитализации)": (d, "bb"),
        "Любой выкуп (даже меньше 0,1%)": (d, "bb_any"),
        "Выкуп от 0,5% капитализации": (d, "bb_05"),
        "Выкуп от 1% капитализации": (d, "bb_1"),
        "Без особых выкупов (по требованию акционеров, после IPO, у нерезидентов)":
            (d[d.bb_type.fillna("") == ""], "bb"),
        "Только частные компании": (d[d.state == 0], "bb"),
        "Без финансового сектора": (d[d.bank == 0], "bb"),
        "2017–2021": (d[d.year <= 2021], "bb"),
        "2022–2025": (d[d.year >= 2022], "bb"),
        "Первый выкуп (раньше не выкупала)": (d[d.f_bb_ever == 0], "bb"),
        "Только годы, где выкуп виден в ДДС": (d[d.bb_src == "ДДС"], "bb"),
        "Крупный выкуп (от 5% капитализации) против лет без выкупа": (d[(d.bb_share >= 0.05) | (d.bb_any == 0)], "bb"),
        "Небольшой выкуп (0,1–5%) против лет без выкупа": (d[(d.bb_share < 0.05) | (d.bb_any == 0)], "bb"),
        "Погашение акций в году t (по файлу Мосбиржи)": (d, "cancel"),
    }
    rows = []
    for name, (x, target) in variants.items():
        row = {"Вариант": name, "Лет": len(x), "Лет с выкупом": int(x[target].sum())}
        for c in feats:
            xx = x[[c, target, "year", "company"]].dropna()
            if xx[target].sum() >= 5 and (xx[target] == 0).sum() >= 5 and xx[c].nunique() > 1:
                row[NAMES[c]] = auc(xx[target].values, score(xx, c).values)
        rows.append(row)
    return pd.DataFrame(rows)


SCORE_ITEMS = {  # простой балл: каждое условие на конец года t-1 даёт 1 очко
    "Частная компания (не госкомпания)": lambda d: d.state == 0,
    "ROE в верхней трети года": lambda d: score(d, "f_roe") >= 2 / 3,
    "Денежные средства / активы в верхней трети года": lambda d: score(d, "f_cash_assets") >= 2 / 3,
    "Прирост денежных средств в верхней трети года": lambda d: score(d, "f_cash_growth") >= 2 / 3,
    "P/BV в верхней трети года": lambda d: score(d, "f_pbv") >= 2 / 3,
}


def score_table(d):
    """Доля лет с выкупом в зависимости от числа выполненных условий — по периодам."""
    pts = sum(f(d).fillna(False).astype(int) for f in SCORE_ITEMS.values())
    rows = []
    for name, sub, cap in [("Все годы", d.index == d.index, 4), ("Без выкупа в прошлом году (начало выкупа)", d.f_bb_prev == 0, 4)]:
        for per, m in [("2017–2025", d.year > 0), ("2017–2021", d.year <= 2021), ("2022–2025", d.year >= 2022)]:
            x = d[sub & m]
            p = pts[sub & m].clip(upper=cap)
            row = {"Выборка": name, "Период": per, "AUC балла": auc(x.bb.values, p.values)}
            for k in range(cap + 1):
                g = x.bb[p == k]
                lab = f"{k}+" if k == cap else str(k)
                row[f"Условий {lab}: лет"] = len(g)
                row[f"Условий {lab}: доля с выкупом"] = g.mean() if len(g) else np.nan
            rows.append(row)
    return pd.DataFrame(rows), pts


def price_weakness(d):
    """Составной показатель слабости цены за год t-1: среднее «обратных» рангов трёх ценовых признаков."""
    parts = [1 - d.groupby("year")[c].rank(pct=True) for c in ["f_ret_vs_imoex", "f_ret6_vs_imoex", "f_dd_52w"]]
    return pd.concat(parts, axis=1).mean(axis=1, skipna=False)


def clogit(x, cols, target="bb"):
    x = x[cols + [target, "company"]].dropna()
    x = x[x.groupby("company")[target].transform("nunique") > 1]
    m = ConditionalLogit(x[target], x[cols], groups=x.company).fit(disp=0)
    return m, x.company.nunique(), len(x)


def sector_adjusted(d, feats):
    rows = []
    for c in feats:
        x = d[[c, "bb", "year", "sector", "company"]].dropna()
        r = x.groupby(["year", "sector"])[c].rank(pct=True)
        ok = x.groupby(["year", "sector"])[c].transform("count") >= 4
        rows.append({"Признак": NAMES[c], "AUC: ранг среди всех компаний года": auc(x.bb.values, score(x, c).values),
                     "AUC: ранг внутри своей отрасли и года (отраслей-лет от 4 компаний)": auc(x.bb[ok].values, r[ok].values),
                     "Лет в расчёте по отраслям": int(ok.sum())})
    return pd.DataFrame(rows)


def main():
    raw = pd.read_csv(HERE / "dataset.csv")
    d = raw[raw.bb.notna()].reset_index(drop=True)
    summary = {"лет": len(d), "компаний": int(d.company.nunique()), "лет_с_выкупом": int(d.bb.sum()),
               "доля": float(d.bb.mean())}

    uni = univariate(d, FEATS)
    uni_private = univariate(d[d.state == 0].reset_index(drop=True), [f for f in FEATS if f != "state"], boot=False)
    conc = univariate(d, CONCURRENT)
    wb = within_between(d, [f for f in FEATS if f not in {"state", "foreign", "bank", "has_pref", "f_bb_ever"}] + list(CONCURRENT))

    models = {
        "1. Только «выкуп в прошлом году»": ["f_bb_prev"],
        "2. История выкупов и собственник": ["f_bb_prev", "f_treasury_cap", "state"],
        "3. Финансы и рынок без истории выкупов": ["state", "bank", "f_roe", "f_cash_assets", "f_cash_growth",
                                                    "f_debt_assets", "f_pbv", "f_pbv_vs_past", "f_ret_vs_imoex",
                                                    "f_turnover", "f_log_cap", "f_div_yield"],
        "4. Всё вместе": ["f_bb_prev", "f_treasury_cap", "state", "bank", "f_roe", "f_cash_assets", "f_cash_growth",
                          "f_debt_assets", "f_pbv", "f_pbv_vs_past", "f_ret_vs_imoex", "f_turnover", "f_log_cap",
                          "f_div_yield"],
    }
    ev = evaluate_models(d, models)
    coef_all, _ = coefficients(d, models["4. Всё вместе"])
    coef_nohist, _ = coefficients(d, models["3. Финансы и рынок без истории выкупов"])

    rule_feats = [f for f in FEATS if f not in {"f_pbv_vs_own"}]
    rules_all, base_tr, base_te = search_rules(d, rule_feats)
    rules_nohist, _, _ = search_rules(d, rule_feats, exclude_history=True)

    key = list(uni.head(15)["код"]) + [c for c in ["f_pbv_vs_past", "f_ret_vs_imoex", "f_cash_growth", "f_roe_change"]
                                       if c not in set(uni.head(15)["код"])]
    rob = robustness(d, key)
    rob_within = []
    for name, x in {"Основной вариант": d, "Без особых выкупов": d[d.bb_type.fillna("") == ""],
                    "Только частные компании": d[d.state == 0], "Только годы, где выкуп виден в ДДС": d[d.bb_src == "ДДС"],
                    "Выкуп от 0,5% капитализации": d.assign(bb=d.bb_05),
                    "Любой выкуп": d.assign(bb=d.bb_any)}.items():
        w = within_between(x.reset_index(drop=True), ["f_pbv", "f_pbv_vs_past", "f_pe_vs_past", "f_ret_vs_imoex",
                                                      "f_ret2_vs_imoex", "f_px_vs_3y", "f_dd_3y", "f_cash_growth",
                                                      "f_eq_growth", "f_roe_change", "f_roe", "f_cash_assets", "f_turnover"])
        w.insert(0, "Вариант", name)
        rob_within.append(w)
    rob_within = pd.concat(rob_within)

    # привычка: начало и продолжение выкупа
    trans = pd.crosstab(d.f_bb_prev.map({0.0: "в прошлом году выкупа не было", 1.0: "в прошлом году выкупала"}),
                        d.bb.map({0.0: "в году t не выкупала", 1.0: "в году t выкупала"}))
    trans["Доля выкупавших в году t"] = trans["в году t выкупала"] / trans.sum(axis=1)
    trans = trans.reset_index().rename(columns={"f_bb_prev": "Прошлый год"})
    start = d[d.f_bb_prev == 0].reset_index(drop=True)
    uni_start = univariate(start, [f for f in FEATS if f not in {"f_bb_prev"}])
    cont = d[d.f_bb_prev == 1].reset_index(drop=True)
    uni_cont = univariate(cont, [f for f in FEATS if f not in {"f_bb_prev", "f_bb_ever"}], boot=False)
    scores, pts = score_table(d)
    sect = sector_adjusted(d, ["f_pbv", "f_roe", "f_cash_assets", "f_cash_growth", "f_debt_assets", "f_treasury_cap",
                               "f_turnover", "f_rev_growth", "f_ps"])
    # «когда»: составная слабость цены и совместная условная модель
    d["weak"] = price_weakness(d)
    NAMES["weak"] = "Слабость цены за год (составной: доходность за год и за 6 мес. к IMOEX, просадка от максимума)"
    wk = []
    for name, x in {"Основной вариант": d, "Без особых выкупов": d[d.bb_type.fillna("") == ""],
                    "Только частные компании": d[d.state == 0], "Начало выкупа (в прошлом году не было)": d[d.f_bb_prev == 0],
                    "Только годы, где выкуп виден в ДДС": d[d.bb_src == "ДДС"]}.items():
        for cols in (["weak"], ["f_cash_growth"], ["f_eq_growth"], ["f_roe_change"], ["f_ni_growth"],
                     ["f_cash_growth", "f_eq_growth", "f_ni_growth"], ["f_cash_growth", "f_roe_change", "weak"]):
            xr = x.copy()
            for c in cols:
                if c != "weak":
                    xr[c] = score(xr, c)
            try:
                m, nco, nobs = clogit(xr, cols)
            except Exception:  # noqa: BLE001
                continue
            for c in cols:
                wk.append({"Вариант": name, "Модель": " + ".join(NAMES[k] for k in cols), "Признак": NAMES[c],
                           "Отношение шансов (низ → верх года)": np.exp(m.params[c]), "p-значение": m.pvalues[c],
                           "Компаний": nco, "Лет": nobs})
    wk = pd.DataFrame(wk)
    comp = d.groupby("company").agg(sector=("sector", "first"), state=("state", "max"), n=("bb", "size"),
                                    k=("bb", "sum"), share=("bb", "mean"), first=("year", "min"), last=("year", "max"))
    comp = comp.sort_values(["share", "k"], ascending=False).reset_index()
    comp.columns = ["Компания (тикер)", "Отрасль", "Госкомпания", "Лет с данными", "Лет с выкупом", "Доля лет с выкупом",
                    "Первый год", "Последний год"]

    years = d.groupby("year").agg(n=("bb", "size"), k=("bb", "sum"), share=("bb", "mean"),
                                  imoex=("c_imoex_ret_1y", "first"), mdd=("c_imoex_mdd_year", "first")).reset_index()
    years.columns = ["Год", "Компаний с данными", "Из них выкупали", "Доля выкупавших",
                     "Доходность IMOEX за год", "Максимальная просадка IMOEX за год"]
    ev_list = d[d.bb == 1][["company", "year", "sector", "bb_amt", "bb_share", "bb_src", "bb_type", "cancel",
                            "f_bb_prev", "state"]].copy()
    ev_list["size"] = np.where(ev_list.bb_share >= 0.05, "крупный (от 5%)", "обычный (0,1–5%)")
    ev_list = ev_list.rename(columns={"company": "Компания (тикер)", "year": "Год", "sector": "Отрасль",
                                      "bb_amt": "Сумма выкупа, млн руб.", "bb_share": "Доля капитализации",
                                      "bb_src": "Как определён", "bb_type": "Особый случай", "cancel": "Погашение в том же году",
                                      "f_bb_prev": "Выкуп в прошлом году", "state": "Госкомпания", "size": "Размер"})

    # краткие цифры для отчёта
    summary["по_годам"] = years.round(4).to_dict("records")
    summary["топ_признаки"] = uni.head(20)[["код", "Признак", "AUC", "Вариант признака", "Доля среди лет с выкупом",
                                             "Доля среди лет без выкупа", "Во сколько раз чаще у выкупавших",
                                             "q-значение (поправка на множественные сравнения)"]].round(4).to_dict("records")
    summary["модели"] = ev.round(4).to_dict("records")
    summary["правила_база"] = [base_tr, base_te]
    summary["привычка"] = trans.round(4).to_dict("records")
    summary["начало"] = uni_start.head(15)[["Признак", "AUC", "Вариант признака", "Доля среди лет с выкупом",
                                             "Доля среди лет без выкупа", "Во сколько раз чаще у выкупавших",
                                             "Вероятность выкупа, если признак есть", "Вероятность выкупа, если признака нет",
                                             "q-значение (поправка на множественные сравнения)"]].round(4).to_dict("records")
    summary["балл"] = scores.round(4).to_dict("records")
    summary["когда"] = wk.round(4).to_dict("records")
    summary["отрасль"] = sect.round(4).to_dict("records")
    (HERE / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1, default=float))

    legend = pd.DataFrame({"Пояснение": [
        "Наблюдение — компания и год t. Выкуп — покупка собственных акций в году t на сумму не меньше 0,1% капитализации "
        "на конец года t-1 (по ДДС; если в ДДС нет такой строки — по росту казначейских акций в балансе).",
        "Признаки (кроме листа «В год выкупа») измерены на конец года t-1, то есть до года выкупа.",
        "AUC: 0,5 — признак не отличает годы с выкупом от лет без выкупа; 0,6 — слабое различие; 0,7 — заметное; "
        "меньше 0,5 — у выкупавших признак, наоборот, ниже. Сравнение — по рангу внутри года.",
        "«Во сколько раз чаще»: доля лет с выкупом, где признак был, делённая на долю лет без выкупа, где он был. "
        "Для непрерывных показателей признак = попадание в верхнюю или нижнюю треть компаний своего года.",
        "95% интервалы — бутстреп по компаниям (1000 повторов). q-значение — p-значение с поправкой на то, что "
        "проверено около 50 признаков (метод Бенджамини–Хохберга); q < 0,05 — различие вряд ли случайное.",
        "«Внутри компании» — сравниваются только годы одной и той же компании (условная логит-модель): "
        "отвечает на вопрос «когда компания выкупает», а не «какие компании выкупают».",
        "Госкомпании и иностранные холдинги помечены из общих знаний (не из файлов) — перед цитированием проверить.",
        f"Выборка: {len(d)} компаний-лет, {d.company.nunique()} компаний, {int(d.bb.sum())} лет с выкупом "
        f"({d.bb.mean():.0%}).",
    ]})
    rob_cols = ["Вариант", "Лет", "Лет с выкупом"]
    with pd.ExcelWriter(HERE / "precursors_results.xlsx") as xw:
        legend.to_excel(xw, sheet_name="Как читать", index=False)
        uni.drop(columns=["код", "Сила различия |AUC−0,5|"]).to_excel(xw, sheet_name="Признаки по одному", index=False)
        uni_private.drop(columns=["код", "Сила различия |AUC−0,5|"]).to_excel(xw, sheet_name="Только частные", index=False)
        wb.drop(columns=["код"]).to_excel(xw, sheet_name="Внутри и между компаниями", index=False)
        ev.to_excel(xw, sheet_name="Модели", index=False)
        coef_all.to_excel(xw, sheet_name="Модель 4 — коэффициенты", index=False)
        coef_nohist.to_excel(xw, sheet_name="Модель 3 — коэффициенты", index=False)
        rules_all.to_excel(xw, sheet_name="Правила", index=False)
        rules_nohist.to_excel(xw, sheet_name="Правила без истории", index=False)
        rob[rob_cols + [c for c in rob.columns if c not in rob_cols]].to_excel(xw, sheet_name="Устойчивость (AUC)", index=False)
        rob_within.drop(columns=["код"]).to_excel(xw, sheet_name="Устойчивость внутри компании", index=False)
        conc.drop(columns=["код", "Сила различия |AUC−0,5|"]).to_excel(xw, sheet_name="В год выкупа", index=False)
        trans.to_excel(xw, sheet_name="Привычка", index=False)
        uni_start.drop(columns=["код", "Сила различия |AUC−0,5|"]).to_excel(xw, sheet_name="Начало выкупа", index=False)
        uni_cont.drop(columns=["код", "Сила различия |AUC−0,5|"]).to_excel(xw, sheet_name="Продолжение выкупа", index=False)
        scores.to_excel(xw, sheet_name="Балл из 5 условий", index=False)
        wk.to_excel(xw, sheet_name="Когда — внутри компании", index=False)
        sect.to_excel(xw, sheet_name="Поправка на отрасль", index=False)
        comp.to_excel(xw, sheet_name="Компании", index=False)
        years.to_excel(xw, sheet_name="По годам", index=False)
        ev_list.to_excel(xw, sheet_name="Годы с выкупом", index=False)

    pd.set_option("display.width", 250)
    pd.set_option("display.max_rows", 200)
    show = ["Признак", "AUC", "AUC: 95% интервал, нижняя граница", "AUC: 95% интервал, верхняя граница", "Вариант признака",
            "Доля среди лет с выкупом", "Доля среди лет без выкупа", "Во сколько раз чаще у выкупавших",
            "q-значение (поправка на множественные сравнения)"]
    print(uni[show].round(3).to_string())
    print(uni_private[["Признак", "AUC", "Во сколько раз чаще у выкупавших", "q-значение (поправка на множественные сравнения)"]].head(15).round(3).to_string())
    print(wb.round(3).to_string())
    print(ev.round(3).to_string())
    print(coef_all.round(3).to_string())
    print(rules_all.round(3).to_string())
    print(rules_nohist.round(3).to_string())
    print(rob.round(3).T.to_string())
    print(rob_within[["Вариант", "Признак", "Компаний, где были годы и с выкупом, и без", "Внутри компании: отношение шансов (низ → верх года)",
                      "Внутри компании: p-значение", "Доля компаний, у которых в годы выкупа выше"]].round(3).to_string())
    print(conc[show].round(3).to_string())
    print(years.round(3).to_string())
    print(trans.round(3).to_string())
    print(uni_start[show].head(20).round(3).to_string())
    print(scores.round(3).T.to_string())
    print(wk.round(3).to_string())
    print(sect.round(3).to_string())
    print(comp.round(2).to_string())


if __name__ == "__main__":
    main()
