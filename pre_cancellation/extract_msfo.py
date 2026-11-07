"""Извлечение ключевых строк из файлов МСФО проекта в единый длинный формат.

Запуск: python3 extract_msfo.py  (читает все .xlsx из папки raw_msfo/ в корне репозитория)
Результат: msfo_long.csv — компания, файл, лист, показатель, год, значение (млн руб.),
исходная строка (для ручной проверки сопоставления) и пометка «оценка».
"""
import re
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).absolute().parent.parent / "raw_msfo"  # сюда положить исходные .xlsx МСФО
OUT = Path(__file__).with_name("msfo_long.csv")

# Показатель -> список регулярных выражений по названию строки (по приоритету).
# Сопоставление проверяется вручную по колонке source_row.
ITEMS = {
    "revenue": [r"^(итого |всего )?выручк[аи]( от (реализации|продаж)( продукции| товаров)?)?( всего| итого)?$",
                r"^(итого |всего )?выручка", r"^доходы от реализации", r"^(итого )?выручка, нетто"],
    "net_income": [r"^(чистая )?прибыль\s*/?\s*\(?(убыток)?\)?\s*за (отчетный )?(год|период)",
                   r"^чистая прибыль\s*/?\s*\(?(убыток)?\)?$", r"^чистая прибыль", r"^прибыль за год",
                   r"^(чистый )?убыток\s*/?\s*\(?прибыль\)?\s*за (год|период)", r"^итого прибыль за год",
                   r"^прибыль от продолжающейся деятельности$"],
    "net_income_tci": [r"^прибыль за год и прочий совокупный доход"],
    "net_income_parent": [r"акционерам (материнской|компании|пао|общества|группы)",
                          r"собственникам (материнской|компании|общества)", r"владельцам (материнской|обыкновенных)"],
    "operating_profit": [r"^операционная прибыль", r"^прибыль от операционной деятельности",
                         r"^(прибыль|результат) от (основной|операционной) деятельности"],
    "equity": [r"^итого (собственный )?капитал$", r"^итого собственн(ый|ого) капитал", r"^итого капитал",
               r"^всего (собственный )?капитал", r"^капитал, итого", r"^итого собственные средства"],
    "total_assets": [r"^(итого|всего) актив(ы|ов)$", r"^(итого|всего) актив", r"^активы, итого", r"^баланс$"],
    "total_liabilities": [r"^(итого|всего) обязательств(а)?$", r"^(итого|всего) обязательств"],
    "cash": [r"^денежные средства и (их )?эквиваленты( денежных средств)?$",
             r"^денежные средства и (их )?эквиваленты", r"^денежные средства и краткосрочные депозиты"],
    "ocf": [r"чист\w* (денежн\w* (средства|поток\w*)|поток\w* денежных средств).{0,40}операционн",
            r"^(итого )?денежн\w* (средства|поток).{0,30}операционн", r"операционной деятельности, нетто",
            r"^(итого )?(чистые )?денежные средства от операционной",
            r"^чистая сумма (поступлений )?денежных средств.{0,30}операционн",
            r"^чистое поступление.{0,40}операционн", r"^итого использование денежных средств на операционную"],
    "capex": [r"^(приобретение|покупка|поступление|капитальные вложения в) (объектов )?основн\w* средств",
              r"^капитальные (затраты|вложения)", r"^приобретение основных средств и нематериальных"],
    "dividends_paid": [r"^дивиденды,? выплаченные", r"^выплат\w* дивидендов", r"^выплаченные дивиденды",
                       r"^дивиденды уплаченные", r"^дивиденды выплачен",
                       r"^\(?дивиденды,? (выплаченные|уплаченные)\)?", r"^дивиденды,? уплаченные акционерам",
                       r"^дивиденды и иные аналогичные выплаты"],
    "buyback": [r"(приобретение|выкуп|покупка) (собственных|казначейских|квазиказначейских) акций",
                r"(приобретение|выкуп) собственных долевых"],
    "debt_lt": [r"^долгосрочные (кредиты и займы|заемные средства|кредиты|займы|долговые)",
                r"^кредиты и займы,? долгосрочн"],
    "debt_st": [r"^краткосрочные (кредиты и займы|заемные средства|кредиты|займы|долговые)",
                r"^кредиты и займы,? краткосрочн", r"^текущая часть долгосрочных (кредитов|займов)"],
    "interest_expense": [r"^процентные расходы", r"^финансовые расходы"],
}
# Строки, которые похожи на нужные, но означают другое
EXCLUDE = {
    "equity": r"обязательств",
    "net_income": r"совокупн|приходящ|причитающ|акционерам|неконтрол",
    "ocf": r"\bдо\b|налога на прибыль и процент",
    "dividends_paid": r"неконтрол|меньшинств|полученн|объявлен",
    "cash": r"начало|конец|на 1 января|на 31 декабря",
    "capex": r"отсрочк|продаж|выбыт",
    "revenue": r"себестоим",
}
SHEET_HINT = {  # на каком листе искать (подстроки в названии листа, нижний регистр)
    "balance": ["баланс", "полож", "фп", "bs"],
    "pl": ["офр", "опу", "прибыл", "доход", "совокуп", "pl", "результ"],
    "cf": ["ддс", "движен", "денеж", "cf"],
}
ITEM_SHEET = {"revenue": "pl", "net_income": "pl", "net_income_tci": "pl", "net_income_parent": "pl", "operating_profit": "pl",
              "interest_expense": "pl", "equity": "balance", "total_assets": "balance",
              "total_liabilities": "balance", "cash": "balance", "debt_lt": "balance", "debt_st": "balance",
              "ocf": "cf", "capex": "cf", "dividends_paid": "cf", "buyback": "cf"}


def to_year(v):
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year") and not isinstance(v, (int, float)):
        y = v.year
    else:
        m = re.search(r"(20[0-3]\d|201\d)", str(v))
        if not m:
            return None
        try:
            if re.fullmatch(r"\s*(20\d\d)(\.0+)?\s*(г\.?|год)?\s*", str(v)) is None and not re.search(
                    r"(31\.12\.|декабр|fy|12м|12 мес)", str(v).lower()):
                return None
        except TypeError:
            return None
        y = int(m.group(1))
    return y if 2010 <= y <= 2030 else None


def unit_factor(text):
    t = text.lower()
    usd = "долл" in t or "usd" in t or "$" in t
    # берём единицу, упомянутую первой («млн руб.; 2016–2017 пересчитаны из тыс. руб.» -> млн)
    pos = {f: m.start() for f, pat in [(1000.0, r"млрд|миллиард"), (0.001, r"тыс"), (1.0, r"млн|миллион")]
           if (m := re.search(pat, t))}
    f = min(pos, key=pos.get) if pos else None
    return f, usd


def parse_sheet(df):
    """Находит строку с годами, колонку названий и единицы измерения листа."""
    head = df.iloc[:20]
    best = None
    for i, row in head.iterrows():
        years = {j: to_year(v) for j, v in row.items() if to_year(v)}
        if len(years) >= 3 and (best is None or len(years) > len(best[1])):
            best = (i, years)
    if best is None:
        return None
    hdr, years = best
    # колонка названий — та, где больше всего текстовых ячеек
    textcnt = df.iloc[hdr + 1:].apply(lambda c: c.map(lambda v: isinstance(v, str)).sum())
    label_col = textcnt.idxmax()
    top_text = " ".join(str(v) for v in df.iloc[:hdr + 1].values.ravel() if isinstance(v, str))
    return hdr, years, label_col, top_text


def year_cols_dedupe(years):
    seen, out = set(), {}
    for j, y in years.items():  # первая колонка года (обычно годовые данные слева от кварталов)
        if y not in seen:
            out[j] = y
            seen.add(y)
    return out


def main():
    files = sorted(p for p in ROOT.glob("*.xlsx") if not p.name.startswith("~"))
    rows, log = [], []
    for f in files:
        try:
            book = pd.read_excel(f, sheet_name=None, header=None)
        except Exception as e:
            log.append((f.name, "", f"не читается: {e}"))
            continue
        file_units = None
        for sname, df in book.items():
            parsed = parse_sheet(df)
            if not parsed:
                continue
            hdr, years, label_col, top_text = parsed
            years = year_cols_dedupe(years)
            fac, usd = unit_factor(top_text)
            if fac is None:
                # единицы могут быть указаны в самой строке-заголовке или других листах
                fac, usd = unit_factor(" ".join(str(v) for v in df.iloc[:hdr + 3].values.ravel()))
            if fac is None:
                fac = file_units
            else:
                file_units = fac
            kind = next((k for k, hints in SHEET_HINT.items() if any(h in sname.lower() for h in hints)), None)
            body = df.iloc[hdr + 1:]
            for item, pats in ITEMS.items():
                if kind and ITEM_SHEET[item] != kind:
                    continue
                found = None
                for pat in pats:
                    for i, r in body.iterrows():
                        lab = r[label_col]
                        if not isinstance(lab, str):
                            continue
                        l = re.sub(r"\s+", " ", lab.strip().lower().replace("ё", "е")).lstrip("-–— •")
                        if re.search(pat, l) and not (item in EXCLUDE and re.search(EXCLUDE[item], l)):
                            vals = {y: r[j] for j, y in years.items() if isinstance(r[j], (int, float, np.number)) and pd.notna(r[j])}
                            if vals:
                                found = (i, lab, vals)
                                break
                    if found:
                        break
                if not found:
                    continue
                i, lab, vals = found
                # пометка оценочных значений в строке/листе
                for y, v in vals.items():
                    rows.append({"file": f.name, "sheet": sname, "item": item, "year": y, "hinted": kind is not None,
                                 "value_raw": float(v), "unit_factor": fac, "usd": usd,
                                 "value_mln": float(v) * fac if fac else np.nan,
                                 "source_row": lab.strip()[:90], "row_idx": i})
    long = pd.DataFrame(rows)
    # если показатель найден на нескольких листах — берём первый по порядку листов
    # приоритет у профильных листов (баланс/ОФР/ДДС) над справочными
    long = long.sort_values("hinted", ascending=False, kind="stable").drop_duplicates(["file", "item", "year"], keep="first")
    long.to_csv(OUT, index=False)
    pd.DataFrame(log, columns=["file", "sheet", "msg"]).to_csv(OUT.with_name("extract_log.csv"), index=False)
    print(long.groupby("item").file.nunique().sort_values())
    print("файлов:", long.file.nunique(), "из", len(files))


if __name__ == "__main__":
    main()
