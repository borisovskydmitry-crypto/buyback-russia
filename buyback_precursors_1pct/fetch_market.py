"""Загрузка дневных свечей (цена закрытия, оборот в рублях) с ISS Мосбиржи для всех тикеров.

Запуск: python3 fetch_market.py [--all]
  без --all — тикеры компаний, по которым есть МСФО (плюс IMOEX);
  с --all   — ещё все тикеры из файла «Анализ акций Мосбиржи».
Кэш: market_cache/<TICKER>.csv (date, close, value). Уже скачанные тикеры пропускаются.
"""
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd
import requests

HERE = Path(__file__).absolute().parent
CACHE = HERE / "market_cache"
ISS = "https://iss.moex.com/iss"
BOARDS = ["TQBR", "TQPI", "TQDE", "TQNE", "EQBR"]
sys.path.insert(0, str(HERE.parent / "pre_cancellation"))
from build_panel import FILE_TICKERS  # noqa: E402


def get(path, **params):
    params.setdefault("iss.meta", "off")
    for attempt in range(7):
        try:
            r = requests.get(f"{ISS}/{path}", params=params, timeout=40)
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError):
            time.sleep(min(2 ** attempt, 30))
    raise RuntimeError(path)


def candles(url, start="2015-01-01"):
    rows, off = [], 0
    while True:
        chunk = get(url, **{"from": start, "interval": 24, "start": off,
                            "candles.columns": "begin,close,value"})["candles"]["data"]
        rows += chunk
        if len(chunk) < 500:
            break
        off += 500
    df = pd.DataFrame(rows, columns=["date", "close", "value"]).dropna(subset=["close"])
    df["date"] = pd.to_datetime(df.date).dt.normalize()
    return df.groupby("date").last()


def fetch(ticker, market="shares"):
    f = CACHE / f"{ticker}.csv"
    if f.exists():
        return ticker, "cached"
    try:
        if market == "index":
            s = candles(f"engines/stock/markets/index/securities/{ticker}/candles.json")
        else:
            s = pd.DataFrame()
            for b in BOARDS:
                if b != BOARDS[0] and not s.empty and s.index.max() > pd.Timestamp.now() - pd.Timedelta(days=30):
                    break
                part = candles(f"engines/stock/markets/shares/boards/{b}/securities/{ticker}/candles.json")
                s = part if s.empty else s.combine_first(part)
        s.to_csv(f)
        return ticker, len(s)
    except Exception as e:  # noqa: BLE001
        return ticker, f"ошибка: {e}"


def main():
    CACHE.mkdir(exist_ok=True)
    tickers = sorted({t for v in FILE_TICKERS.values() for t in v})
    if "--all" in sys.argv:
        df = pd.read_excel(HERE.parent / "pre_cancellation" / "Анализ_акций_Мосбиржи.xlsx", header=1)
        extra = df.iloc[:, 0].astype(str)
        extra = extra[extra.str.fullmatch(r"[A-Z0-9\-]+")]
        tickers += sorted(set(extra) - set(tickers))
    print(fetch("IMOEX", "index"), flush=True)
    print(fetch("MCFTR", "index"), flush=True)
    with ThreadPoolExecutor(4) as ex:
        for res in ex.map(fetch, tickers):
            print(res, flush=True)


if __name__ == "__main__":
    main()
