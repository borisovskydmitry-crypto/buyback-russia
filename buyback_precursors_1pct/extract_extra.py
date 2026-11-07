"""Дополнительные строки МСФО для поиска предвестников выкупа.

Запуск: python3 extract_extra.py  ->  msfo_extra.csv
Что извлекается (млн руб., единицы — как в extract_msfo.py, с теми же поправками):
- buyback_cf     — отток на покупку собственных акций в ДДС (сумма всех подходящих строк, только отрицательные);
- treasury_bs    — собственные (казначейские) акции в балансе (обычно отрицательное число в капитале);
- da             — износ и амортизация (из ДДС, иначе из ОФР);
- borrowings     — кредиты, займы и облигации (сумма долгосрочной и краткосрочной части, только в разделе
                   обязательств баланса; аренда, выданные займы и проценты к получению исключены).
Каждая найденная строка сохраняется с исходным названием для ручной проверки (msfo_extra_rows.csv).
"""
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).absolute().parent
sys.path.insert(0, str(HERE.parent / "pre_cancellation"))
from extract_msfo import ROOT, SHEET_HINT, parse_sheet, unit_factor, year_cols_dedupe  # noqa: E402

PATTERNS = {
    "buyback_cf": (r"(приобретение|выкуп|покупка|платежи за( обратный)? выкуп|расходы по выкупу|погашение задолженности за выкуп)"
                   r".{0,25}(собственн|казначейск|квазиказначейск|обыкновенн\w* акци|привилегированн\w* акци|акций$|акций \()"
                   r"|обратн\w* выкуп",
                   r"дочерн|неконтрол|продаж|поступлен|реализац|дол(и|ей) в|доля |миноритар|опцион|ассоциир|"
                   r"инвестиционн|облигац|долгов|возврат|распределен|справочно|дивиденд|затраты, связанные|всего|нетто"),
    "treasury_bs": (r"(собственные|казначейские|выкупленные|квазиказначейские).{0,30}акци|акции,? выкупленные",
                    r"дочерн|ассоциир|привлечен|эмиссион|премия|прибыль на|итого|%|шт\.?|резерв|невыкупленн|"
                    r"справочно|уменьшающие|в т\.ч\."),
    "da": (r"^(износ и )?амортизаци|^амортизация и (износ|обесценение)|^износ основных средств",
           r"дисконт|премии|облигац|финансов|лицензий$|права пользования$|аквизиц|субсиди"),
    "borrowings": (r"кредиты и займы|заемные средства|займы и кредиты|кредиты банков|облигаци|"
                   r"долговые (ценные бумаги|обязательства)|выпущенные долговые|заимствования|^займы( полученные)?$|^кредиты$",
                   r"выданн|предоставлен|к получению|аренд|процент|итого обязательств|дебитор|инвестиц|депозит|"
                   r"нетто|чист\w* долг|^итого$|поступлен|погашени|привлечени|капитал|нераспредел|бессрочн|"
                   r"субординир|структурн|цфа"),
}
WHERE = {"buyback_cf": "cf", "treasury_bs": "balance", "da": "cf", "borrowings": "balance"}
UNIT_OVERRIDE = {"АФК_Система": 1.0, "Селигдар": 0.001}


def drop_subtotal_sum(s):
    """Сумма строк долга за год; строка, равная сумме остальных (подитог), не учитывается."""
    v = s.abs().sort_values(ascending=False).tolist()
    if len(v) >= 3 and abs(v[0] - sum(v[1:])) <= 0.01 * max(v[0], 1):
        v = v[1:]
    return float(sum(v))


def main():
    files = sorted(p for p in ROOT.glob("*.xlsx") if not p.name.startswith("~"))
    found = []
    for f in files:
        if f.name == "Лукойл_отчетность.xlsx":
            continue
        try:
            book = pd.read_excel(f, sheet_name=None, header=None)
        except Exception:  # noqa: BLE001
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
                fac, usd = unit_factor(" ".join(str(v) for v in df.iloc[:hdr + 3].values.ravel()))
            fac = fac if fac is not None else file_units
            file_units = fac
            for k, v in UNIT_OVERRIDE.items():
                if k in f.name:
                    fac = v
            kind = next((k for k, hints in SHEET_HINT.items() if any(h in sname.lower() for h in hints)), None)
            if kind is None:
                continue
            body = df.iloc[hdr + 1:]
            in_liab = False
            for i, r in body.iterrows():
                lab = r[label_col]
                if not isinstance(lab, str):
                    continue
                l = re.sub(r"\s+", " ", lab.strip().lower().replace("ё", "е")).lstrip("-–— •⚠")
                if kind == "balance" and re.search(r"обязательства|^долгосрочные обязательства|^краткосрочные обязательства", l):
                    in_liab = True
                for item, (pat, excl) in PATTERNS.items():
                    if WHERE[item] != kind and not (item == "da" and kind == "pl"):
                        continue
                    if item == "borrowings" and not in_liab:
                        continue
                    if not re.search(pat, l) or re.search(excl, l):
                        continue
                    vals = {y: r[j] for j, y in years.items()
                            if isinstance(r[j], (int, float, np.number)) and pd.notna(r[j])}
                    for y, val in vals.items():
                        found.append({"file": f.name, "sheet": sname, "kind": kind, "item": item, "year": y,
                                      "value_raw": float(val), "unit_factor": fac, "usd": usd,
                                      "source_row": lab.strip()[:100], "row_idx": i})
    rows = pd.DataFrame(found)
    rows.to_csv(HERE / "msfo_extra_rows.csv", index=False)

    out = []
    for (f, item), g in rows.groupby(["file", "item"]):
        if item == "treasury_bs":
            first_row = g.sort_values("row_idx").row_idx.iloc[0]  # первая подходящая строка баланса
            g = g[g.row_idx == first_row]
            agg = g.groupby("year").value_raw.first()
        elif item == "da":
            # ДДС приоритетнее ОФР. Есть общая строка («Амортизация», «Износ и амортизация», «ОС и НМА») —
            # берём её; иначе складываем отдельные строки (ОС + НМА + права пользования).
            sheet = "cf" if (g.kind == "cf").any() else "pl"
            g = g[(g.kind == sheet) & (g.sheet == g[g.kind == sheet].sheet.iloc[0])]
            total = g.source_row.str.lower().str.contains(
                r"^(износ и )?амортизация(\s|$|\()|и нематериальн|ос и нма|ос, |и износ|и истощение|внеоборотных")
            if total.any():
                first_row = g[total].sort_values("row_idx").row_idx.iloc[0]
                agg = g[g.row_idx == first_row].groupby("year").value_raw.first().abs()
            else:
                agg = g.groupby("year").value_raw.apply(lambda s: s.abs().sum())
        elif item == "buyback_cf":
            agg = g[g.value_raw < 0].groupby("year").value_raw.sum()
            agg = agg.reindex(sorted(g.year.unique()), fill_value=0.0)
        else:  # borrowings: строка «совокупный долг», иначе сумма строк долга без подитогов
            sheet = g.sheet.iloc[0]
            g = g[g.sheet == sheet]
            tot = g.source_row.str.lower().str.contains(r"^(совокупный |итого |общий )?долг|^долг:")
            if tot.any():
                first_row = g[tot].sort_values("row_idx").row_idx.iloc[0]
                agg = g[g.row_idx == first_row].groupby("year").value_raw.first().abs()
            else:
                agg = g.groupby("year").value_raw.apply(drop_subtotal_sum)
        fac = g.unit_factor.iloc[0]
        for y, v in agg.items():
            out.append({"file": f, "item": item, "year": y, "value_raw": v, "unit_factor": fac,
                        "usd": bool(g.usd.iloc[0]), "rows": " | ".join(sorted(set(g.source_row)))[:300]})
    out = pd.DataFrame(out)
    out.to_csv(HERE / "msfo_extra.csv", index=False)
    print(out.groupby("item").file.nunique())


if __name__ == "__main__":
    main()
