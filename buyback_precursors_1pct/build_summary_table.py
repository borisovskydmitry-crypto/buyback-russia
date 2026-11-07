"""Сводная таблица «компания × год»: выкуп и пять условий балла (частная; ROE, P/BV, денежные средства
к активам и их прирост в верхней трети года).

Запуск: python3 build_summary_table.py   (после build_dataset.py; берёт dataset.csv и panel_full.csv)
Выход:  Сводная_таблица_по_компаниям.xlsx. Ранги, флаги условий, балл и итоги по компаниям — формулы Excel,
поэтому таблица пересчитывается, если поменять исходные значения. Строки 2026 г. — показатели на конец 2025 г.
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).absolute().parent
sys.path.insert(0, str(HERE))
import analyze as an  # noqa: E402

OUT = HERE / "Сводная_таблица_по_компаниям.xlsx"
NEXT = 2026  # год, для которого выкуп ещё неизвестен

NAME_RU = {  # подписи для таблицы
    "ABIO": "ИСКЧ (Артген)", "ABRD": "Абрау-Дюрсо", "AFKS": "АФК «Система»", "AFLT": "Аэрофлот", "AKRN": "Акрон",
    "ALRS": "АЛРОСА", "APRI": "АПРИ", "APTK": "Аптечная сеть 36,6", "AQUA": "Инарктика", "BANE": "Башнефть",
    "BELU": "Новабев (Белуга)", "BSPB": "Банк «Санкт-Петербург»", "CBOM": "Московский кредитный банк",
    "CHMF": "Северсталь", "DELI": "Делимобиль (каршеринг)", "DOMRF": "ДОМ.РФ", "EELT": "Европейская электротехника",
    "ENPG": "ЭН+", "EUTR": "ЕвроТранс", "FEES": "Россети (ФСК)", "FLOT": "Совкомфлот", "GCHE": "Черкизово",
    "GMKN": "Норникель", "HEAD": "Хэдхантер", "HYDR": "РусГидро", "IRAO": "Интер РАО", "IRKT": "Яковлев (Иркут)",
    "KAZT": "КуйбышевАзот", "KMAZ": "КАМАЗ", "KUZB": "Банк «Кузнецкий»", "LENT": "Лента", "LKOH": "Лукойл",
    "LSNG": "Россети Ленэнерго", "LSRG": "Группа ЛСР", "LVHK": "Левенгук", "MBNK": "МТС-Банк", "MDMG": "МД Медикал",
    "MGKL": "МГКЛ", "MGNT": "Магнит", "MOEX": "Московская биржа", "MRKC": "Россети Центр",
    "MRKP": "Россети Центр и Приволжье", "MRKS": "Россети Сибирь", "MRKV": "Россети Волга",
    "MRKZ": "Россети Северо-Запад", "MSNG": "Мосэнерго", "MSRS": "Россети Московский регион", "MTSS": "МТС",
    "NKNC": "Нижнекамскнефтехим", "NMTP": "НМТП", "NVTK": "НОВАТЭК", "PHOR": "ФосАгро", "PIKK": "ПИК",
    "PLZL": "Полюс", "POSI": "Позитив", "RAGR": "Русагро", "RASP": "Распадская", "RENI": "Ренессанс Страхование",
    "ROSN": "Роснефть", "ROST": "Росинтер", "RTKM": "Ростелеком", "RUAL": "РУСАЛ", "SELG": "Селигдар",
    "SGZH": "Сегежа", "SOFL": "Софтлайн", "SVAV": "Соллерс", "SVCB": "Совкомбанк", "TATN": "Татнефть",
    "TGKA": "ТГК-1", "TGKN": "ТГК-14", "TRMK": "ТМК", "TRNFP": "Транснефть", "UGLD": "ЮГК", "URKA": "Уралкалий",
    "URKZ": "Уральская кузница", "USBN": "Уралсиб", "VKCO": "ВК", "WUSH": "Whoosh", "X5": "ИКС 5",
}

# (заголовок, ширина) — порядок столбцов листа «Компании по годам»
COLS = [
    ("Компания", 26), ("Тикер", 8), ("Отрасль", 17), ("Госкомпания", 11), ("Год выкупа (t)", 9),
    ("Показатели на конец года", 11), ("Выкуп в году t (≥ 0,1% капитализации)", 13), ("Почему неизвестно", 30),
    ("Сумма выкупа, млн руб.", 12), ("Выкуп, доля капитализации на конец t−1", 13), ("Как определён выкуп", 22),
    ("Особый выкуп", 26), ("Погашение акций в году t (файл Мосбиржи)", 12), ("Погашено, доля акций", 10),
    ("ROE", 9), ("P/BV", 8), ("Денежные средства / активы", 11), ("Прирост денежных средств / активы", 11),
    ("Выкуп в прошлом году", 10), ("Входит в базу рангов года", 10),
    ("Ранг ROE в году", 9), ("Ранг P/BV в году", 9), ("Ранг денег / активов в году", 9), ("Ранг прироста денег в году", 9),
    ("Условие 1: частная компания", 10), ("Условие 2: ROE в верхней трети", 10), ("Условие 3: P/BV в верхней трети", 10),
    ("Условие 4: деньги / активы в верхней трети", 11), ("Условие 5: прирост денег в верхней трети", 11),
    ("Балл (сколько из 5 условий)", 10), ("Ключ (тикер|год)", 12),
]
L = {h: get_column_letter(i + 1) for i, (h, _) in enumerate(COLS)}

FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
CALC_FILL = PatternFill("solid", fgColor="F2F2F2")
NEXT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
WRAP = Alignment(wrap_text=True, vertical="top")


def yes_no(v):
    return "" if pd.isna(v) else ("да" if v else "нет")


def r6(v):
    return None if pd.isna(v) or not np.isfinite(v) else round(float(v), 6)


def unknown_reason(r):
    if r.bb_src == "казначейских акций стало меньше — покупки не видны":
        return "казначейских акций стало меньше (погашение или продажа), покупки за год не видны"
    if pd.notna(r.bb_amt) and pd.isna(r.cap_ref):
        return "нет рыночной капитализации (акции не торговались)"
    return "нет данных о выкупе: ни строки в ДДС, ни казначейских акций за два года подряд"


def cancel_status(r):
    if r.year == NEXT:
        return "ещё неизвестно"
    if pd.isna(r.cap):
        return "нет данных"  # в файле Мосбиржи нет числа акций, если акции не торговались
    return "да" if r.cancel == 1 else "нет"


def load_rows():
    d = pd.read_csv(HERE / "dataset.csv")
    p = pd.read_csv(HERE / "panel_full.csv")
    last = p[p.year == NEXT - 1].copy()
    nxt = pd.DataFrame({"company": last.company, "year": NEXT, "sector": last.sector, "state": last.state,
                        "f_roe": last.roe, "f_pbv": last.pbv, "f_cash_assets": last.cash_assets,
                        "f_cash_growth": last.cash_growth, "f_bb_prev": last.bb})
    rows = pd.concat([d, nxt], ignore_index=True).sort_values(["company", "year"]).reset_index(drop=True)
    for c in ["f_roe", "f_pbv", "f_cash_assets", "f_cash_growth"]:
        rows[c] = rows[c].map(r6)
    return d, rows


def style_header(ws, row, ncol, height=60):
    for j in range(1, ncol + 1):
        c = ws.cell(row, j)
        c.font, c.fill, c.alignment = BOLD, HEAD_FILL, Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[row].height = height


RAW_SHEET = "Исходные данные"
SRC_SHEET = "Строки отчётности"
# исходные показатели: код -> заголовок (млн руб.)
RAW_ITEMS = [("revenue", "Выручка"), ("operating_profit", "Операционная прибыль"), ("net_income", "Чистая прибыль"),
             ("equity", "Итого капитал"), ("total_assets", "Итого активы"), ("total_liabilities", "Итого обязательства"),
             ("cash", "Денежные средства"), ("borrowings", "Кредиты и займы"), ("ocf", "Операционный денежный поток"),
             ("capex", "Капитальные затраты (ДДС)"), ("dividends_paid", "Дивиденды выплаченные (ДДС)"),
             ("buyback_cf", "Выкуп собственных акций (ДДС)"), ("treasury_bs", "Казначейские акции (баланс)"),
             ("da", "Амортизация")]
ITEM_RU = dict(RAW_ITEMS) | {"net_income_parent": "Чистая прибыль акционеров материнской компании",
                             "net_income_tci": "Итого совокупный доход", "interest_expense": "Процентные расходы",
                             "debt_st": "Краткосрочные кредиты", "debt_lt": "Долгосрочные кредиты", "buyback": "Выкуп собственных акций (ДДС)"}
RAW_COLS = ([("Компания", 24), ("Тикер", 8), ("Год (на 31 декабря)", 9), ("Ключ (тикер|год)", 12), ("Файл МСФО", 34),
             ("Пересчёт в млн руб.", 26)]
            + [(f"{h}, млн руб.", 13) for _, h in RAW_ITEMS]
            + [("Чистая прибыль: какая строка", 20),
               ("Акций обыкн. на конец года, шт.", 15), ("Капитализация обыкн., млн руб.", 13),
               ("Акций прив. на конец года, шт.", 15), ("Капитализация прив., млн руб.", 13),
               ("Капитализация всего, млн руб.", 13), ("Капитализация: источник", 20), ("Цена обыкн. акции, руб.", 11),
               ("Капитал на конец прошлого года, млн руб.", 13), ("Средний капитал, млн руб.", 13),
               ("ROE = прибыль / средний капитал", 10), ("P/BV = капитализация / капитал", 10),
               ("Денежные средства / активы", 10), ("Прирост денежных средств / активы", 10),
               ("Капзатраты / выручка", 10), ("P/S = капитализация / выручка", 10), ("Чистая маржа", 10),
               ("Дивиденды / капитализация", 10), ("Выкуп по ДДС / капитализация прошлого года", 11)])
R = {h: get_column_letter(i + 1) for i, (h, _) in enumerate(RAW_COLS)}
RAW_FILL = PatternFill("solid", fgColor="FFFFFF")
EST_FILL = PatternFill("solid", fgColor="FCE4D6")  # оценка или подстановка
BLUE = Font(name="Arial", size=10, color="0000FF")


def load_raw():
    """Исходные значения ровно в том виде, в каком они вошли в panel_full.csv, и строки файлов-источников."""
    sys.path.insert(0, str(HERE.parent / "pre_cancellation"))
    import build_panel as bp
    import build_dataset as bd
    _, long = bp.load_msfo()
    x = pd.read_csv(HERE / "msfo_extra.csv")
    x = x[x.file != "Лукойл_отчетность.xlsx"]
    xm = bd.load_extra()
    xr = pd.read_csv(HERE / "msfo_extra_rows.csv")
    xr = xr[xr.file.isin(set(x.file))]
    p = pd.read_csv(HERE / "panel_full.csv")
    moex = bp.load_moex()
    files = long.groupby("company").file.first()
    xfile = xr.copy()
    xfile["company"] = xfile.file.map(lambda f: (bp.file_key(f)[1] or [None])[0])
    files = files.combine_first(xfile.dropna(subset=["company"]).groupby("company").file.first())
    # множитель к млн руб. для каждой ячейки = итог / значение в файле
    long["mult"] = long.value_mln / long.value_raw.where(long.value_raw != 0)
    xr["company"] = xr.file.map(lambda f: (bp.file_key(f)[1] or [None])[0])
    agg = xr.groupby(["company", "item", "year"]).value_raw.sum()
    xm_long = xm.melt(id_vars=["company", "year"], var_name="item", value_name="mln").dropna()
    mult = (xm_long.set_index(["company", "item", "year"]).mln / agg.where(agg != 0)).rename("mult")
    xr = xr.join(mult, on=["company", "item", "year"])
    return p, long, xr, moex, files


def conv_note(file):
    if any(k in file for k in ("EnPlus", "RUSAL")):
        return "млн долл. × курс ЦБ (баланс — на конец года, потоки — средний)"
    if "КуйбышевАзот" in file:
        return "взяты значения в млн руб. из первых столбцов"
    if "АФК_Система" in file or "Селигдар" in file:
        return "единицы в файле указаны неверно — исправлено вручную по величинам"
    if "Магнит" in file:
        return "2016–2017 гг. в тыс. долл. — исключены"
    return ""


def raw_sheet(wb):
    p, long, xr, moex, files = load_raw()
    ws = wb.create_sheet(RAW_SHEET)
    ws.append([h for h, _ in RAW_COLS])
    m = moex.set_index(["ticker", "year"])
    tickers = {v[0]: v for v in __import__("build_panel").FILE_TICKERS.values()}
    tci_rows = set()
    if "net_income_tci" in set(long.item):
        have = set(zip(long[long.item == "net_income"].company, long[long.item == "net_income"].year))
        tci = long[long.item == "net_income_tci"]
        tci_rows = {(c, y) for c, y in zip(tci.company, tci.year) if (c, y) not in have}
    p = p.sort_values(["company", "year"]).reset_index(drop=True)
    n = len(p) + 1
    rng = lambda h: f"${R[h]}$2:${R[h]}${n}"  # noqa: E731
    est_cells = []
    for i, r in enumerate(p.itertuples(), start=2):
        f = files.get(r.company, "")
        cls = tickers.get(r.company, [r.company])
        get = lambda t, c: m[c].get((t, r.year), np.nan) if t in m.index.get_level_values(0) else np.nan  # noqa: E731
        row = {"Компания": NAME_RU.get(r.company, r.company), "Тикер": r.company, "Год (на 31 декабря)": int(r.year),
               "Файл МСФО": f, "Пересчёт в млн руб.": conv_note(f)}
        for code, h in RAW_ITEMS:
            v = getattr(r, code)
            row[f"{h}, млн руб."] = None if pd.isna(v) else round(float(v), 3)
        row["Чистая прибыль: какая строка"] = ("итого совокупный доход (строки чистой прибыли нет)"
                                               if (r.company, r.year) in tci_rows else ("" if pd.isna(r.net_income) else "чистая прибыль"))
        if (r.company, r.year) in tci_rows:
            est_cells.append((i, "Чистая прибыль, млн руб."))
        sh1, cp1 = get(cls[0], "shares"), get(cls[0], "cap_bln")
        row["Акций обыкн. на конец года, шт."] = None if pd.isna(sh1) else int(sh1)
        row["Капитализация обыкн., млн руб."] = None if pd.isna(cp1) else round(cp1 * 1000, 3)
        if len(cls) > 1:
            sh2, cp2 = get(cls[1], "shares"), get(cls[1], "cap_bln")
            row["Акций прив. на конец года, шт."] = None if pd.isna(sh2) else int(sh2)
            row["Капитализация прив., млн руб."] = None if pd.isna(cp2) else round(cp2 * 1000, 3)
        src = "" if pd.isna(r.cap_src) else r.cap_src
        row["Капитализация: источник"] = src
        k_o, k_p = R["Капитализация обыкн., млн руб."], R["Капитализация прив., млн руб."]
        if src == "файл Мосбиржи":
            row["Капитализация всего, млн руб."] = f"=SUM({k_o}{i},{k_p}{i})"
        elif src:
            row["Капитализация всего, млн руб."] = round(float(r.cap), 3)
            est_cells.append((i, "Капитализация всего, млн руб."))
        sh = R["Акций обыкн. на конец года, шт."]
        row["Цена обыкн. акции, руб."] = f'=IF(OR({sh}{i}="",{k_o}{i}=""),"",{k_o}{i}*1000000/{sh}{i})'
        for h, v in row.items():
            ws[f"{R[h]}{i}"] = v
        c = {k: f"{R[k]}{i}" for k in R}
        col = lambda name: c[f"{name}, млн руб."]  # noqa: E731
        look = lambda name: f"INDEX({rng(name)},MATCH({c['Тикер']}&\"|\"&({c['Год (на 31 декабря)']}-1),{rng('Ключ (тикер|год)')},0))"  # noqa: E731
        prev = lambda name: f'IF({look(name)}="","",{look(name)})'  # noqa: E731
        cap, eq, ta, cash, rev, ni = (c["Капитализация всего, млн руб."], col("Итого капитал"), col("Итого активы"),
                                      col("Денежные средства"), col("Выручка"), col("Чистая прибыль"))
        lim = lambda expr, lo, hi: f'IF(OR({expr}<{lo},{expr}>{hi}),"",ROUND({expr},10))'  # noqa: E731
        f_ = {
            "Ключ (тикер|год)": f'={c["Тикер"]}&"|"&{c["Год (на 31 декабря)"]}',
            "Капитал на конец прошлого года, млн руб.": f'=IFERROR({prev("Итого капитал, млн руб.")},"")',
            "Средний капитал, млн руб.": (f'=IF(OR({eq}="",{c["Капитал на конец прошлого года, млн руб."]}=""),"",'
                                          f'({eq}+{c["Капитал на конец прошлого года, млн руб."]})/2)'),
            "ROE = прибыль / средний капитал": (f'=IF(OR({ni}="",{c["Средний капитал, млн руб."]}=""),"",'
                                                f'IF({c["Средний капитал, млн руб."]}<=0,"",'
                                                + lim(f'{ni}/{c["Средний капитал, млн руб."]}', -5, 5) + "))"),
            "P/BV = капитализация / капитал": f'=IF(OR({cap}="",{eq}=""),"",IF({eq}<=0,"",' + lim(f"{cap}/{eq}", 0, 100) + "))",
            "Денежные средства / активы": f'=IF(OR({cash}="",{ta}=""),"",ROUND({cash}/{ta},10))',
            "Прирост денежных средств / активы": (f'=IFERROR(IF(OR({cash}="",{ta}=""),"",ROUND(({cash}-{prev("Денежные средства, млн руб.")})/{ta},10)),"")'),
            "Капзатраты / выручка": f'=IF(OR({col("Капитальные затраты (ДДС)")}="",{rev}=""),"",IF({rev}<=0,"",ABS({col("Капитальные затраты (ДДС)")})/{rev}))',
            "P/S = капитализация / выручка": f'=IF(OR({cap}="",{rev}=""),"",IF({rev}<=0,"",' + lim(f"{cap}/{rev}", 0, 100) + "))",
            "Чистая маржа": f'=IF(OR({ni}="",{rev}=""),"",IF({rev}<=0,"",' + lim(f"{ni}/{rev}", -5, 5) + "))",
            "Дивиденды / капитализация": f'=IF(OR({col("Дивиденды выплаченные (ДДС)")}="",{cap}=""),"",' + lim(f'ABS({col("Дивиденды выплаченные (ДДС)")})/{cap}', 0, 1) + ")",
            "Выкуп по ДДС / капитализация прошлого года": (f'=IFERROR(IF({col("Выкуп собственных акций (ДДС)")}="","",'
                                                           f'MAX(0,-{col("Выкуп собственных акций (ДДС)")})/{prev("Капитализация всего, млн руб.")}),"")'),
        }
        for h, v in f_.items():
            ws[f"{R[h]}{i}"] = v
    start = [h for h, _ in RAW_COLS].index("Капитал на конец прошлого года, млн руб.")
    fmls = {h for h, _ in RAW_COLS[start:]} | {"Ключ (тикер|год)", "Цена обыкн. акции, руб."}
    pct = {"ROE = прибыль / средний капитал", "Денежные средства / активы", "Прирост денежных средств / активы",
           "Капзатраты / выручка", "Чистая маржа", "Дивиденды / капитализация", "Выкуп по ДДС / капитализация прошлого года"}
    for j, (h, w) in enumerate(RAW_COLS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for row in ws.iter_rows(min_row=2, max_row=n):
        for j, cell in enumerate(row):
            h = RAW_COLS[j][0]
            cell.font = FONT
            if h.endswith("млн руб."):
                cell.number_format = "#,##0.0"
            if h.endswith("шт."):
                cell.number_format = "#,##0"
            if h in pct:
                cell.number_format = "0.0%"
            if h in ("P/BV = капитализация / капитал", "P/S = капитализация / выручка", "Цена обыкн. акции, руб."):
                cell.number_format = "0.00"
            if h == "Год (на 31 декабря)":
                cell.number_format = "0"
            if h in fmls or h == "Капитализация всего, млн руб.":
                cell.fill = CALC_FILL
            elif h.endswith("млн руб.") or h.endswith("шт."):
                cell.font = BLUE
    for i, h in est_cells:
        cell = ws[f"{R[h]}{i}"]
        cell.fill, cell.font = EST_FILL, BLUE
    style_header(ws, 1, len(RAW_COLS), 90)
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(RAW_COLS))}{n}"

    # строки файлов-источников
    ws2 = wb.create_sheet(SRC_SHEET)
    heads = [("Тикер", 8), ("Год", 7), ("Показатель", 28), ("Файл МСФО", 36), ("Лист", 14), ("Номер строки в файле", 9),
             ("Название строки в файле", 50), ("Значение в файле", 14), ("Множитель к млн руб.", 11), ("Значение, млн руб.", 14)]
    ws2.append([h for h, _ in heads])
    src = pd.concat([
        long.assign(show=long.value_mln)[["company", "year", "item", "file", "sheet", "row_idx", "source_row", "value_raw", "mult", "show"]],
        xr.assign(show=xr.value_raw * xr.mult)[["company", "year", "item", "file", "sheet", "row_idx", "source_row", "value_raw", "mult", "show"]],
    ]).dropna(subset=["company"])
    src = src[src.item.isin(set(dict(RAW_ITEMS)) | {"net_income_tci"})]
    src = src.sort_values(["company", "year", "item"])
    for r in src.itertuples():
        ws2.append([r.company, int(r.year), ITEM_RU.get(r.item, r.item), r.file, r.sheet, int(r.row_idx) + 1,
                    str(r.source_row).strip(), r.value_raw, None if pd.isna(r.mult) else round(float(r.mult), 6),
                    None if pd.isna(r.show) else round(float(r.show), 3)])
    for j, (h, w) in enumerate(heads, start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w
    for row in ws2.iter_rows(min_row=2, max_row=len(src) + 1):
        for j, cell in enumerate(row):
            cell.font = FONT
            if j in (7, 9):
                cell.number_format = "#,##0.0"
    style_header(ws2, 1, len(heads), 45)
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{len(src) + 1}"
    return n


def main_sheet(wb, rows, nraw):
    ws = wb.active
    ws.title = "Компании по годам"
    ws.append([h for h, _ in COLS])
    n = len(rows) + 1
    rng = lambda h: f"${L[h]}$2:${L[h]}${n}"  # noqa: E731
    yr, inr = rng("Год выкупа (t)"), rng("Входит в базу рангов года")
    feats = [("ROE", "Ранг ROE в году", "Условие 2: ROE в верхней трети"),
             ("P/BV", "Ранг P/BV в году", "Условие 3: P/BV в верхней трети"),
             ("Денежные средства / активы", "Ранг денег / активов в году", "Условие 4: деньги / активы в верхней трети"),
             ("Прирост денежных средств / активы", "Ранг прироста денег в году", "Условие 5: прирост денег в верхней трети")]
    for i, r in enumerate(rows.itertuples(), start=2):
        nxt = r.year == NEXT
        status = "ещё неизвестно" if nxt else ("неизвестно" if pd.isna(r.bb) else yes_no(r.bb))
        vals = {
            "Компания": NAME_RU.get(r.company, r.company), "Тикер": r.company, "Отрасль": r.sector,
            "Госкомпания": yes_no(r.state), "Год выкупа (t)": int(r.year), "Показатели на конец года": int(r.year) - 1,
            "Выкуп в году t (≥ 0,1% капитализации)": status,
            "Почему неизвестно": unknown_reason(r) if status == "неизвестно" else "",
            "Сумма выкупа, млн руб.": None if nxt or pd.isna(r.bb_amt) else round(float(r.bb_amt), 1),
            "Выкуп, доля капитализации на конец t−1": None if nxt else r6(r.bb_share),
            "Как определён выкуп": "" if nxt or pd.isna(r.bb_src) else r.bb_src,
            "Особый выкуп": "" if pd.isna(r.bb_type) else r.bb_type,
            "Погашение акций в году t (файл Мосбиржи)": cancel_status(r),
            "Погашено, доля акций": r6(r.cancel_pct) if r.cancel == 1 else None,
            "ROE": r.f_roe, "P/BV": r.f_pbv, "Денежные средства / активы": r.f_cash_assets,
            "Прирост денежных средств / активы": r.f_cash_growth, "Выкуп в прошлом году": yes_no(r.f_bb_prev),
            "Входит в базу рангов года": "да" if nxt or pd.notna(r.bb) else "нет",
        }
        for h, v in vals.items():
            ws[f"{L[h]}{i}"] = v
        key = f'{L["Тикер"]}{i}&"|"&{L["Показатели на конец года"]}{i}'
        for h, rh in [("ROE", "ROE = прибыль / средний капитал"), ("P/BV", "P/BV = капитализация / капитал"),
                      ("Денежные средства / активы", "Денежные средства / активы"),
                      ("Прирост денежных средств / активы", "Прирост денежных средств / активы")]:
            src = f"'{RAW_SHEET}'!${R[rh]}$2:${R[rh]}${nraw}"
            keys = f"'{RAW_SHEET}'!${R['Ключ (тикер|год)']}$2:${R['Ключ (тикер|год)']}${nraw}"
            ws[f"{L[h]}{i}"] = f'=IFERROR(INDEX({src},MATCH({key},{keys},0)),"")'
        for val_h, rank_h, cond_h in feats:
            v, c = f"{L[val_h]}{i}", rng(val_h)
            base = f"{yr},${L['Год выкупа (t)']}{i},{inr},\"да\",{c}"
            own = f'(${L["Входит в базу рангов года"]}{i}="да")'
            ws[f"{L[rank_h]}{i}"] = (f'=IF({v}="","",(COUNTIFS({base},"<"&{v})+(COUNTIFS({base},{v})+1+NOT{own})/2)'
                                     f'/(COUNTIFS({base},">-1E+300")+NOT{own}))')
            ws[f"{L[cond_h]}{i}"] = f'=IF({L[rank_h]}{i}="","нет",IF({L[rank_h]}{i}>=2/3,"да","нет"))'
        ws[f"{L['Условие 1: частная компания']}{i}"] = f'=IF({L["Госкомпания"]}{i}="нет","да","нет")'
        c1, c5 = L["Условие 1: частная компания"], L["Условие 5: прирост денег в верхней трети"]
        ws[f"{L['Балл (сколько из 5 условий)']}{i}"] = f'=COUNTIF({c1}{i}:{c5}{i},"да")'
        ws[f"{L['Ключ (тикер|год)']}{i}"] = f'={L["Тикер"]}{i}&"|"&{L["Год выкупа (t)"]}{i}'

    fmt = {"Погашено, доля акций": "0.0%", "Год выкупа (t)": "0", "Показатели на конец года": "0", "Сумма выкупа, млн руб.": "#,##0.0",
           "Выкуп, доля капитализации на конец t−1": "0.00%", "ROE": "0.0%", "P/BV": "0.00",
           "Денежные средства / активы": "0.0%", "Прирост денежных средств / активы": "0.0%",
           **{h: "0.00" for _, h, _ in feats}}
    calc = {h for _, h, _ in feats} | {h for _, _, h in feats} | {"Условие 1: частная компания",
                                                                   "Балл (сколько из 5 условий)", "Ключ (тикер|год)"}
    for j, (h, w) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for row in ws.iter_rows(min_row=2, max_row=n):
        nxt = row[4].value == NEXT
        for j, c in enumerate(row):
            h = COLS[j][0]
            c.font = FONT
            if h in fmt:
                c.number_format = fmt[h]
            if nxt:
                c.fill = NEXT_FILL
            elif h in calc:
                c.fill = CALC_FILL
    style_header(ws, 1, len(COLS), 75)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{n}"
    return n


def company_sheet(wb, d, rows, n):
    ws = wb.create_sheet("По компаниям")
    M = "'Компании по годам'!"
    rng = lambda h: f"{M}${L[h]}$2:${L[h]}${n}"  # noqa: E731
    tk, yr, st, inr, pts, cn = (rng("Тикер"), rng("Год выкупа (t)"), rng("Выкуп в году t (≥ 0,1% капитализации)"),
                                rng("Входит в базу рангов года"), rng("Балл (сколько из 5 условий)"),
                                rng("Погашение акций в году t (файл Мосбиржи)"))
    heads = [("Компания", 26), ("Тикер", 8), ("Отрасль", 17), ("Госкомпания", 11),
             ("Лет в таблице (2017–2025)", 10), ("Лет с известным выкупом", 10), ("Лет с выкупом", 9),
             ("Доля лет с выкупом", 9), ("Годы с выкупом", 34), ("Лет с погашением акций", 10), ("Годы погашения", 22),
             ("Лет с 4–5 условиями (из лет с известным выкупом)", 12), ("Из них с выкупом", 9), ("Средний балл", 9),
             ("Доля лет: ROE в верхней трети", 10), ("Доля лет: P/BV в верхней трети", 10),
             ("Доля лет: деньги / активы в верхней трети", 11), ("Доля лет: прирост денег в верхней трети", 11),
             ("Выкуп в 2025 г.", 12), ("Балл на конец 2025 г. (условия для 2026 г.)", 13)]
    ws.append([h for h, _ in heads])
    known = d[d.bb.notna()]
    per = (d.groupby("company").agg(sector=("sector", "first"), state=("state", "first")).join(
        known.groupby("company").bb.agg(["size", "mean"])).fillna({"size": 0}))
    per = per.assign(name=[NAME_RU.get(c, c) for c in per.index]).sort_values(["mean", "size"], ascending=[False, False])
    join = lambda x: ", ".join(map(str, sorted(x)))  # noqa: E731
    years = known[known.bb == 1].groupby("company").year.apply(join)
    cyears = d[d.cancel == 1].groupby("company").year.apply(join)
    conds = [L["Условие 2: ROE в верхней трети"], L["Условие 3: P/BV в верхней трети"],
             L["Условие 4: деньги / активы в верхней трети"], L["Условие 5: прирост денег в верхней трети"]]
    for i, (co, r) in enumerate(per.iterrows(), start=2):
        b = f"$B{i}"
        last = f'"<={NEXT - 1}"'
        kn = f'{tk},{b},{yr},{last},{inr},"да"'  # годы с известным выкупом
        ws.append([r["name"], co, r.sector, yes_no(r.state),
                   f"=COUNTIFS({tk},{b},{yr},{last})",
                   f"=COUNTIFS({kn})",
                   f'=COUNTIFS({tk},{b},{st},"да")',
                   f'=IF(F{i}=0,"",G{i}/F{i})',
                   years.get(co, ""),
                   f'=COUNTIFS({tk},{b},{cn},"да")',
                   cyears.get(co, ""),
                   f'=COUNTIFS({kn},{pts},">=4")',
                   f'=COUNTIFS({kn},{pts},">=4",{st},"да")',
                   f'=IF(F{i}=0,"",SUMIFS({pts},{kn})/F{i})',
                   *[f'=IF($F{i}=0,"",COUNTIFS({kn},{M}${c}$2:${c}${n},"да")/$F{i})' for c in conds],
                   f'=IFERROR(INDEX({st},MATCH({b}&"|{NEXT - 1}",{rng("Ключ (тикер|год)")},0)),"нет строки")',
                   f'=IFERROR(INDEX({pts},MATCH({b}&"|{NEXT}",{rng("Ключ (тикер|год)")},0)),"нет отчётности за 2025 г.")'])
    m = len(per) + 1
    for j, (h, w) in enumerate(heads, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for row in ws.iter_rows(min_row=2, max_row=m):
        for j, c in enumerate(row):
            c.font = FONT
            c.number_format = {8: "0%", 14: "0.0", 15: "0%", 16: "0%", 17: "0%", 18: "0%"}.get(j + 1, "General")
            if j + 1 not in (1, 2, 3, 4, 9, 11):
                c.fill = CALC_FILL
    style_header(ws, 1, len(heads), 75)
    ws.cell(m + 2, 1, "«Годы с выкупом» и «Годы погашения» — списки, выписанные из dataset.csv при сборке файла (не формулы). "
                      "Остальные серые столбцы считаются формулами по листу «Компании по годам». Средний балл и доли лет "
                      "считаются по годам с известным выкупом.").font = Font(name="Arial", size=9, italic=True)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{m}"


def cancel_sheet(wb, d, n):
    """Все погашения 2017–2025 гг. и условия на конец предыдущего года — ссылками на основной лист."""
    ws = wb.create_sheet("Погашения")
    M = "'Компании по годам'!"
    rng = lambda h: f"{M}${L[h]}$2:${L[h]}${n}"  # noqa: E731
    pull = ["Погашено, доля акций", "Выкуп в году t (≥ 0,1% капитализации)", "Выкуп в прошлом году",
            "ROE", "P/BV", "Денежные средства / активы", "Прирост денежных средств / активы",
            "Условие 1: частная компания", "Условие 2: ROE в верхней трети", "Условие 3: P/BV в верхней трети",
            "Условие 4: деньги / активы в верхней трети", "Условие 5: прирост денег в верхней трети",
            "Балл (сколько из 5 условий)"]
    heads = ["Компания", "Тикер", "Год погашения (t)"] + pull
    ws.append(heads)
    ev = d[d.cancel == 1].sort_values(["company", "year"])
    for i, r in enumerate(ev.itertuples(), start=2):
        row = f"MATCH($B{i}&\"|\"&$C{i},{rng('Ключ (тикер|год)')},0)"
        ws.append([NAME_RU.get(r.company, r.company), r.company, int(r.year)]
                  + [f'=IF(INDEX({rng(h)},{row})="","",INDEX({rng(h)},{row}))' for h in pull])
    m = len(ev) + 1
    fmt = {"Погашено, доля акций": "0.0%", "ROE": "0.0%", "P/BV": "0.00", "Денежные средства / активы": "0.0%",
           "Прирост денежных средств / активы": "0.0%", "Год погашения (t)": "0"}
    for row in ws.iter_rows(min_row=2, max_row=m):
        for j, c in enumerate(row):
            c.font = FONT
            c.number_format = fmt.get(heads[j], "General")
            if j >= 3:
                c.fill = CALC_FILL
    tot = m + 1
    ws.cell(tot, 1, "Итого / среднее").font = BOLD
    for j, h in enumerate(heads, start=1):
        col = get_column_letter(j)
        if h.startswith("Условие") or h in ("Выкуп в году t (≥ 0,1% капитализации)", "Выкуп в прошлом году"):
            c = ws.cell(tot, j, f'=COUNTIF({col}2:{col}{m},"да")/COUNTA({col}2:{col}{m})')
            c.number_format = "0%"
        elif h == "Балл (сколько из 5 условий)":
            c = ws.cell(tot, j, f"=AVERAGE({col}2:{col}{m})")
            c.number_format = "0.0"
        else:
            continue
        c.font, c.fill = BOLD, CALC_FILL
    ws.cell(tot + 1, 1, "Строка «Итого»: для условий и выкупа — доля строк с «да» среди всех погашений "
                        "(пустые значения тоже в знаменателе, поэтому для «Выкупа в прошлом году» это нижняя оценка).").font = Font(name="Arial", size=9, italic=True)
    for j, h in enumerate(heads, start=1):
        ws.column_dimensions[get_column_letter(j)].width = 26 if j == 1 else 11
    style_header(ws, 1, len(heads), 75)
    ws.freeze_panes = "D2"


def check_sheet(wb, n):
    ws = wb.create_sheet("Проверка балла")
    M = "'Компании по годам'!"
    rng = lambda h: f"{M}${L[h]}$2:${L[h]}${n}"  # noqa: E731
    yr, st, inr, pts, cn = (rng("Год выкупа (t)"), rng("Выкуп в году t (≥ 0,1% капитализации)"),
                            rng("Входит в базу рангов года"), rng("Балл (сколько из 5 условий)"),
                            rng("Погашение акций в году t (файл Мосбиржи)"))
    ws.append(["Доля лет с выкупом по числу выполненных условий (годы с известным выкупом, 2017–2025)"])
    ws["A1"].font = Font(name="Arial", size=11, bold=True)
    ws.append([])
    ws.append(["Балл", "Период", "Лет", "Из них с выкупом", "Доля лет с выкупом"])
    style_header(ws, 3, 5, 30)
    r = 4
    scores = [("0", '"0"'), ("1", '"1"'), ("2", '"2"'), ("3", '"3"'), ("4–5", '">=4"')]
    for per, lo, hi in [("2017–2025", 2017, NEXT - 1), ("2017–2021", 2017, 2021), ("2022–2025", 2022, NEXT - 1)]:
        for label, crit in scores:
            base = f'{inr},"да",{yr},">={lo}",{yr},"<={hi}",{pts},{crit}'
            ws.append([label, per, f"=COUNTIFS({base})", f'=COUNTIFS({base},{st},"да")', f'=IF(C{r}=0,"",D{r}/C{r})'])
            r += 1
    for row in ws.iter_rows(min_row=4, max_row=r - 1):
        for j, c in enumerate(row):
            c.font = FONT
            c.number_format = "0%" if j == 4 else "General"
    ws.cell(r + 1, 1, "Сверка с анализом (precursors_results.xlsx, лист с баллом): за 2017–2025 гг. при 0 условий — "
                      "9 лет с выкупом из 90 (10%), при 4–5 условиях — 32 из 68 (47%).").font = Font(name="Arial", size=9, italic=True)
    r += 3
    ws.cell(r, 1, "Доля лет с погашением акций по числу выполненных условий (все годы с торгами, 2017–2025)").font = Font(name="Arial", size=11, bold=True)
    r += 2
    for j, h in enumerate(["Балл", "Период", "Лет", "Из них с погашением", "Доля лет с погашением"], start=1):
        ws.cell(r, j, h)
    style_header(ws, r, 5, 30)
    r0 = r + 1
    for label, crit in scores:
        r += 1
        base = f'{yr},"<={NEXT - 1}",{pts},{crit}'
        for j, v in enumerate([label, "2017–2025", f'=COUNTIFS({base},{cn},"да")+COUNTIFS({base},{cn},"нет")',
                               f'=COUNTIFS({base},{cn},"да")', f'=IF(C{r}=0,"",D{r}/C{r})'], start=1):
            c = ws.cell(r, j, v)
            c.font = FONT
            c.number_format = "0%" if j == 5 else "General"
    ws.cell(r + 2, 1, "Сверка с cancel_precursors.xlsx (лист «Балл и погашение»). Годы без торгов (нет числа акций) не считаются.").font = Font(name="Arial", size=9, italic=True)
    for col, w in zip("ABCDE", [8, 12, 9, 12, 12]):
        ws.column_dimensions[col].width = w


def readme_sheet(wb):
    ws = wb.create_sheet("Как читать")
    lines = [
        ("Сводная таблица по компаниям", True),
        ("Одна строка — компания и год t. Выкуп относится к году t, а показатели (ROE, P/BV, деньги) — к концу года t−1, "
         "то есть к тому, что было известно до выкупа.", False),
        ("Выкуп «да» — компания купила свои акции на сумму не меньше 0,1% капитализации на конец года t−1. "
         "Сумма берётся из отчёта о движении денежных средств (ДДС), а если строки выкупа там нет — из прироста казначейских акций.", False),
        ("«Неизвестно» — выкуп за год определить нельзя; такие годы не входили в анализ выкупов. "
         "Причина указана в отдельном столбце. Балл для них всё равно посчитан (см. про ранги ниже).", False),
        ("Погашение «да» — число акций на конец года t меньше, чем на конец t−1, по файлу Анализ_акций_Мосбиржи.xlsx. "
         "«Нет данных» — акции в этот год не торговались, числа акций в файле нет. Погашение часто идёт из казначейских "
         "акций, купленных в прошлые годы, поэтому выкуп в том же году бывает не всегда. Все погашения собраны на листе «Погашения».", False),
        ("Пять условий балла: 1) частная компания (не госкомпания); 2) ROE в верхней трети компаний того же года; "
         "3) P/BV в верхней трети; 4) денежные средства / активы в верхней трети; 5) прирост денежных средств за год "
         "(в долях активов) в верхней трети. Нет данных по показателю — условие не выполнено.", False),
        ("Ранг — место компании среди компаний того же года с известным выкупом («база рангов»), от 0 до 1 (1 — самое высокое "
         "значение, ничьи делятся пополам). Год, который в базу не входит, ставится на то место, которое занял бы, если его "
         "добавить; на ранги остальных он не влияет. Верхняя треть — ранг не меньше 2/3. Ранги, флаги и балл — формулы Excel.", False),
        ("Жёлтые строки 2026 г. — показатели на конец 2025 г. Выкуп в 2026 г. ещё неизвестен; ранги считаются среди "
         "компаний с отчётностью за 2025 г. Это «прогнозные» строки: какие компании сейчас подходят под условия.", False),
        ("Лист «По компаниям» — итоги за 2017–2025 гг. по каждой компании, отсортированы по доле лет с выкупом; "
         "фильтр по столбцу «Лет с погашением акций» оставляет только гасивших. Лист «Проверка балла» повторяет таблицы "
         "из анализа: 47% лет с выкупом при 4–5 условиях против 10% при нуле, и то же для погашений.", False),
        ("Откуда берутся ROE, P/BV и деньги. Лист «Исходные данные» — все компании и годы: выручка, прибыль, капитал, "
         "активы, деньги, долг, операционный поток, капзатраты, дивиденды, выкуп по ДДС, казначейские акции, число акций и "
         "капитализация. Синие числа введены из источников, серые столбцы — формулы. ROE = чистая прибыль / средний "
         "капитал за год, P/BV = капитализация (все классы акций) / капитал на конец года, деньги / активы и их прирост "
         "за год к активам. Лист «Компании по годам» берёт эти коэффициенты формулой за год t−1.", False),
        ("Откуда исходные числа. Лист «Строки отчётности»: для каждого числа — файл МСФО, лист, номер и название строки, "
         "значение в файле и множитель к млн руб. (тыс. руб. → 0,001; долл. — по курсу ЦБ). Число акций и капитализация — "
         "файл Анализ_акций_Мосбиржи.xlsx (столбцы «Объём год, шт.» и «Кап. год, млрд руб.»).", False),
        ("Оранжевые ячейки — оценки и подстановки: капитализация, пересчитанная по цене Мосбиржи от ближайшего года из файла "
         "(когда в файле её нет), и чистая прибыль, взятая из «итого совокупного дохода», когда строки прибыли нет. "
         "Крайние значения, похожие на ошибки единиц, отброшены: ROE и маржа вне ±500%, P/BV и P/S вне 0–100.", False),
        ("Госкомпании, особые выкупы (по требованию акционеров, стабилизация после IPO, выкуп у нерезидентов) частично "
         "размечены по общим знаниям — это стоит сверить с отчётностью.", False),
        ("Источник: файлы МСФО компаний в папке проекта, цены и капитализация — Московская биржа (ISS). "
         "Собрано скриптом build_summary_table.py из dataset.csv и panel_full.csv.", False),
    ]
    for i, (t, b) in enumerate(lines, start=1):
        c = ws.cell(i, 1, t)
        c.font, c.alignment = (Font(name="Arial", size=12, bold=True) if b else FONT), WRAP
    ws.column_dimensions["A"].width = 120


def main():
    d, rows = load_rows()
    wb = Workbook()
    nraw = raw_sheet(wb)
    n = main_sheet(wb, rows, nraw)
    company_sheet(wb, d, rows, n)
    cancel_sheet(wb, d, n)
    check_sheet(wb, n)
    readme_sheet(wb)
    order = ["Как читать", "Компании по годам", RAW_SHEET, "По компаниям", "Погашения", "Проверка балла", SRC_SHEET]
    wb._sheets = [wb[t] for t in order]
    wb.active = 0
    wb.save(OUT)
    print(OUT, "строк:", len(rows))


if __name__ == "__main__":
    main()
