"""
Stage 2 + strong-fundamentals screener for NSE stocks.

Technical side (weekly bars, Stan Weinstein Stage 1 -> Stage 2):
  1. Weekly close above the 30-week SMA
  2. 30-week SMA rising (above its value 4 weeks ago)
  3. 30-week SMA was flat ~10 weeks ago  -> stock is *entering* Stage 2
  4. Fresh breakout: a weekly close above the prior 26-week high in the last N weeks
  5. Breakout-week volume >= 2x the 10-week average
  6. Relative strength vs Nifty 500 (26-week return beats the index, Mansfield RS > 0)
  7. Liquidity: median daily traded value >= Rs 5 cr
  + Extension check: how far price is above the breakout level (buy the retest if stretched)

Fundamental side (Yahoo Finance data, fetched only for technical survivors):
  market cap, quarterly YoY revenue and profit growth, ROE, debt/equity (non-banks),
  ROA (banks). Pledge data is NOT available from Yahoo: check Screener.in manually,
  or pass a Screener.in export with --screener-file (then its stocks are treated as
  fundamentally pre-approved).

Usage examples:
  python stage2_screener.py                          # Nifty 500 universe
  python stage2_screener.py --symbols MAHABANK PARAGMILK COFORGE
  python stage2_screener.py --tickers-file my_list.txt
  python stage2_screener.py --screener-file screener_export.xlsx
  python stage2_screener.py --loose                  # relax rules (more results)

Output: printed ranking + stage2_results_<date>.xlsx (all scanned stocks, with every check).
This is a screening tool, not investment advice.
"""

from __future__ import annotations

import argparse
import gzip
import io
import json
import math
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date, datetime, timedelta

import pandas as pd
import requests
import yfinance as yf

NIFTY500_URL = "https://archives.nseindia.com/content/indices/ind_nifty500list.csv"
BENCHMARKS = ["^CRSLDX", "^NSEI"]  # Nifty 500, fallback Nifty 50
CRORE = 1e7


@dataclass
class Rules:
    sma_weeks: int = 30
    slope_lookback: int = 4          # SMA now > SMA this many weeks ago
    flat_check_ago: int = 10         # SMA was flat this many weeks ago ...
    flat_window: int = 10            # ... measured over this many weeks
    flat_tol: float = 0.05           # "flat" = moved less than 5%
    base_weeks: int = 26             # breakout above prior 26-week high
    fresh_weeks: int = 6             # breakout must have happened in the last N weeks
    vol_mult: float = 2.0            # breakout-week volume vs 10-week average
    rs_weeks: int = 26
    min_daily_value_cr: float = 5.0
    max_extension: float = 0.20      # > 20% above breakout level = extended
    # fundamentals
    min_mcap_cr: float = 1000
    min_rev_growth: float = 0.15
    min_profit_growth: float = 0.20
    min_roe: float = 0.15
    max_de: float = 0.75
    min_bank_roa: float = 0.01
    max_pledge_pct: float = 5.0        # pledged shares as % of promoter holding
    max_promoter_drop: float = 3.0     # promoter stake fall (percentage points) over ~1 year


def loose(r: Rules) -> Rules:
    r.flat_tol = 0.10
    r.fresh_weeks = 10
    r.vol_mult = 1.5
    r.min_rev_growth = 0.10
    r.min_profit_growth = 0.12
    r.min_roe = 0.12
    r.max_de = 1.0
    return r


# ----------------------------------------------------------------------------- universe
HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(HERE, "lists")


def _cached_list(name: str, fetch) -> list[str]:
    """Fetch a stock list from NSE; keep a copy, and fall back to it when NSE blocks us
    (NSE often refuses requests from cloud servers such as GitHub's)."""
    path = os.path.join(CACHE_DIR, name)
    try:
        syms = fetch()
        if len(syms) > 50:
            os.makedirs(CACHE_DIR, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(syms))
            return syms
    except Exception as e:
        print(f"  NSE list unavailable ({e.__class__.__name__}); using saved copy")
    if os.path.exists(path):
        return load_tickers_file(path)
    sys.exit(f"Could not download {name} from NSE and no saved copy exists in {CACHE_DIR}.")


def load_nifty500() -> list[str]:
    def fetch():
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(NIFTY500_URL, headers=headers, timeout=30)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        return df["Symbol"].astype(str).str.strip().tolist()
    return _cached_list("nifty500.txt", fetch)


NSE_ALL_EQUITY_URL = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"


def load_all_nse(include_be: bool = False) -> list[str]:
    """Every NSE-listed stock in the EQ series (~2,400). BE (trade-to-trade) is optional."""
    def fetch():
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(NSE_ALL_EQUITY_URL, headers=headers, timeout=30)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        df.columns = [c.strip() for c in df.columns]
        series = {"EQ", "BE"} if include_be else {"EQ"}
        df = df[df["SERIES"].astype(str).str.strip().isin(series)]
        return df["SYMBOL"].astype(str).str.strip().tolist()
    return _cached_list("nse_all_be.txt" if include_be else "nse_all.txt", fetch)


def load_tickers_file(path: str) -> list[str]:
    with open(path, encoding="utf-8") as f:
        raw = f.read().replace(",", "\n").split()
    return [s.strip().upper() for s in raw if s.strip()]


def load_screener_file(path: str) -> list[str]:
    """Read a Screener.in export (csv/xlsx). Needs an 'NSE Code' or 'Symbol' column."""
    df = pd.read_excel(path) if path.lower().endswith((".xlsx", ".xls")) else pd.read_csv(path)
    cols = {c.lower().strip(): c for c in df.columns}
    for key in ("nse code", "symbol", "nse symbol", "ticker"):
        if key in cols:
            syms = df[cols[key]].dropna().astype(str).str.strip().str.upper()
            return [s for s in syms if s and s != "NAN"]
    sys.exit(
        "Screener file needs an 'NSE Code' column. In Screener.in, edit your screen's "
        "columns and add 'NSE Code' before exporting."
    )


# ----------------------------------------------------------------------------- data
def download_weekly(symbols: list[str], years: int = 3) -> dict[str, pd.DataFrame]:
    tickers = [f"{s}.NS" for s in symbols]
    out: dict[str, pd.DataFrame] = {}
    chunk = 100
    for i in range(0, len(tickers), chunk):
        part = tickers[i : i + chunk]
        print(f"  downloading {i + 1}-{i + len(part)} of {len(tickers)} ...", flush=True)
        data = yf.download(
            part, period=f"{years}y", interval="1wk", group_by="ticker",
            auto_adjust=True, threads=True, progress=False,
        )
        for t in part:
            try:
                df = data[t] if len(part) > 1 else data
                if isinstance(df.columns, pd.MultiIndex):
                    df = df.droplevel(0, axis=1)
                df = df.dropna(subset=["Close"])
                if len(df) > 0:
                    out[t.removesuffix(".NS")] = df
            except KeyError:
                pass
    return out


def download_benchmark(years: int = 3) -> pd.Series:
    for b in BENCHMARKS:
        df = yf.download(b, period=f"{years}y", interval="1wk", auto_adjust=True, progress=False)
        if isinstance(df.columns, pd.MultiIndex):
            df = df.droplevel(1, axis=1)
        if len(df) > 60:
            print(f"  benchmark: {b}")
            return df["Close"].dropna()
    sys.exit("Could not download a benchmark index from Yahoo Finance.")


# ----------------------------------------------------------------------------- broker data (Upstox / Dhan)
UPSTOX_INSTRUMENTS = "https://assets.upstox.com/market-quote/instruments/exchange/NSE.json.gz"
UPSTOX_CANDLES = "https://api.upstox.com/v3/historical-candle/{key}/days/1/{to}/{frm}"
UPSTOX_NIFTY500 = "NSE_INDEX|Nifty 500"
DHAN_SCRIP_MASTER = "https://images.dhan.co/api-data/api-scrip-master.csv"
DHAN_HISTORICAL = "https://api.dhan.co/v2/charts/historical"


def to_weekly(daily: pd.DataFrame) -> pd.DataFrame:
    """Daily OHLCV -> weekly bars labelled by the week's Friday."""
    if daily.empty:
        return daily
    w = daily.resample("W-FRI").agg(
        {"Open": "first", "High": "max", "Low": "min", "Close": "last", "Volume": "sum"}
    )
    return w.dropna(subset=["Close"])


class RateLimiter:
    """Allow at most `per_sec` calls per second across all threads."""

    def __init__(self, per_sec: float):
        import threading
        self.gap = 1.0 / per_sec
        self.lock = threading.Lock()
        self.next_at = 0.0

    def wait(self) -> None:
        with self.lock:
            now = time.monotonic()
            delay = self.next_at - now
            self.next_at = max(now, self.next_at) + self.gap
        if delay > 0:
            time.sleep(delay)


UPSTOX_LIMIT = RateLimiter(6)   # Upstox allows ~500 requests/minute
DHAN_LIMIT = RateLimiter(4)     # Dhan data APIs are stricter


def _get_with_retry(fn, limiter: RateLimiter, tries: int = 6):
    resp = None
    for attempt in range(tries):
        limiter.wait()
        resp = fn()
        if resp.status_code == 429 or resp.status_code >= 500:
            retry_after = resp.headers.get("Retry-After")
            time.sleep(float(retry_after) if retry_after and retry_after.isdigit() else 5 * (attempt + 1))
            continue
        return resp
    return resp


def _fetch_parallel(symbols, fetch_one, workers: int, label: str) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = {pool.submit(fetch_one, s): s for s in symbols}
        for i, fut in enumerate(as_completed(futs), 1):
            sym = futs[fut]
            try:
                df = fut.result()
                if df is not None and not df.empty:
                    out[sym] = df
            except Exception as e:
                print(f"  {sym}: {e.__class__.__name__}: {e}")
            if i % 100 == 0 or i == len(futs):
                print(f"  {label}: {i}/{len(futs)}", flush=True)
    return out


def upstox_session() -> requests.Session:
    s = requests.Session()
    s.headers["Accept"] = "application/json"
    token = os.environ.get("UPSTOX_ACCESS_TOKEN")  # optional; historical candles work without it
    if token:
        s.headers["Authorization"] = f"Bearer {token}"
    return s


def upstox_daily(session: requests.Session, key: str, years: int) -> pd.DataFrame:
    to = date.today()
    frm = to - timedelta(days=365 * years + 7)
    url = UPSTOX_CANDLES.format(key=requests.utils.quote(key, safe=""), to=to, frm=frm)
    resp = _get_with_retry(lambda: session.get(url, timeout=30), UPSTOX_LIMIT)
    resp.raise_for_status()
    candles = resp.json().get("data", {}).get("candles", [])
    if not candles:
        return pd.DataFrame()
    df = pd.DataFrame(candles, columns=["ts", "Open", "High", "Low", "Close", "Volume", "OI"])
    df.index = pd.to_datetime(df["ts"].str[:10])
    return df.drop(columns=["ts", "OI"]).sort_index()


def download_upstox(symbols: list[str], years: int = 3) -> tuple[dict[str, pd.DataFrame], pd.Series]:
    print("  loading Upstox instrument list ...")
    raw = requests.get(UPSTOX_INSTRUMENTS, timeout=60)
    raw.raise_for_status()
    inst = json.loads(gzip.decompress(raw.content))
    keys = {
        d["trading_symbol"]: d["instrument_key"]
        for d in inst
        if d.get("segment") == "NSE_EQ" and d.get("instrument_type") == "EQ"
    }
    unknown = [s for s in symbols if s not in keys]
    if unknown:
        print(f"  not in Upstox list ({len(unknown)}): {', '.join(unknown[:15])}")
    session = upstox_session()

    def one(sym):
        return to_weekly(upstox_daily(session, keys[sym], years))

    prices = _fetch_parallel([s for s in symbols if s in keys], one, workers=8, label="Upstox")
    bench = to_weekly(upstox_daily(session, UPSTOX_NIFTY500, years))["Close"]
    print("  benchmark: Nifty 500 (Upstox)")
    return prices, bench


def download_dhan(symbols: list[str], years: int = 3) -> tuple[dict[str, pd.DataFrame], pd.Series]:
    client_id = os.environ.get("DHAN_CLIENT_ID")
    token = os.environ.get("DHAN_ACCESS_TOKEN")
    if not client_id or not token:
        sys.exit(
            "Dhan needs two environment variables: DHAN_CLIENT_ID and DHAN_ACCESS_TOKEN.\n"
            "Set them in PowerShell (see README), then run again. Dhan's historical data API\n"
            "also needs the Data API subscription active on your Dhan account."
        )
    print("  loading Dhan scrip master ...")
    sm = pd.read_csv(DHAN_SCRIP_MASTER, low_memory=False)
    sm = sm[(sm.SEM_EXM_EXCH_ID == "NSE") & (sm.SEM_SEGMENT == "E")
            & (sm.SEM_INSTRUMENT_NAME == "EQUITY") & (sm.SEM_SERIES == "EQ")]
    ids = dict(zip(sm.SEM_TRADING_SYMBOL.astype(str), sm.SEM_SMST_SECURITY_ID.astype(str)))
    unknown = [s for s in symbols if s not in ids]
    if unknown:
        print(f"  not in Dhan list ({len(unknown)}): {', '.join(unknown[:15])}")
    session = requests.Session()
    session.headers.update({"access-token": token, "client-id": client_id,
                            "Content-Type": "application/json", "Accept": "application/json"})
    to = date.today()
    frm = to - timedelta(days=365 * years + 7)

    def one(sym):
        body = {"securityId": ids[sym], "exchangeSegment": "NSE_EQ", "instrument": "EQUITY",
                "expiryCode": 0, "oi": False, "fromDate": frm.isoformat(), "toDate": to.isoformat()}
        resp = _get_with_retry(lambda: session.post(DHAN_HISTORICAL, json=body, timeout=30), DHAN_LIMIT)
        if resp.status_code in (401, 403):
            raise RuntimeError(f"Dhan rejected the token (HTTP {resp.status_code}); it may have expired")
        resp.raise_for_status()
        d = resp.json()
        if not d.get("close"):
            return pd.DataFrame()
        df = pd.DataFrame({"Open": d["open"], "High": d["high"], "Low": d["low"],
                           "Close": d["close"], "Volume": d["volume"]},
                          index=pd.to_datetime(d["timestamp"], unit="s", utc=True)
                          .tz_convert("Asia/Kolkata").tz_localize(None).normalize())
        return to_weekly(df.sort_index())

    prices = _fetch_parallel([s for s in symbols if s in ids], one, workers=4, label="Dhan")
    # benchmark from Upstox's public index candles (no login needed)
    bench = to_weekly(upstox_daily(upstox_session(), UPSTOX_NIFTY500, years))["Close"]
    print("  benchmark: Nifty 500 (Upstox public candles)")
    return prices, bench


def drop_partial_week(df: pd.DataFrame, include_current: bool) -> pd.DataFrame:
    """Before Friday's close the latest weekly bar is still forming; drop it unless asked not to."""
    if include_current or df.empty:
        return df
    now = datetime.now()
    today = now.date()
    this_monday = today - timedelta(days=today.weekday())
    week_done = today.weekday() >= 5 or (today.weekday() == 4 and now.hour >= 16)
    if not week_done and df.index[-1].date() >= this_monday:
        return df.iloc[:-1]
    return df


# ----------------------------------------------------------------------------- technical checks
def technical(sym: str, df: pd.DataFrame, bench: pd.Series, r: Rules) -> dict | None:
    need = r.sma_weeks + r.flat_check_ago + r.flat_window + 5
    if len(df) < need:
        return None
    c, h, v = df["Close"], df["High"], df["Volume"]
    sma = c.rolling(r.sma_weeks).mean()
    prior_high = h.shift(1).rolling(r.base_weeks).max()
    avg_vol = v.shift(1).rolling(10).mean()

    last_c = float(c.iloc[-1])
    above_sma = last_c > sma.iloc[-1]
    sma_rising = sma.iloc[-1] > sma.iloc[-1 - r.slope_lookback]
    a, b = sma.iloc[-1 - r.flat_check_ago], sma.iloc[-1 - r.flat_check_ago - r.flat_window]
    was_flat = abs(a / b - 1) < r.flat_tol if b > 0 else False

    # most recent breakout week within the fresh window
    brk_idx = None
    for k in range(1, r.fresh_weeks + 1):
        if c.iloc[-k] > prior_high.iloc[-k]:
            brk_idx = len(c) - k
            break
    fresh_breakout = brk_idx is not None
    brk_level = float(prior_high.iloc[brk_idx]) if fresh_breakout else float(prior_high.iloc[-1])
    vol_ratio = float(v.iloc[brk_idx] / avg_vol.iloc[brk_idx]) if fresh_breakout and avg_vol.iloc[brk_idx] > 0 else float("nan")
    vol_ok = vol_ratio >= r.vol_mult if not math.isnan(vol_ratio) else False
    weeks_since_brk = (len(c) - 1 - brk_idx) if fresh_breakout else None
    extension = last_c / brk_level - 1 if brk_level > 0 else float("nan")
    still_above_brk = last_c >= brk_level * 0.97

    # relative strength vs benchmark
    bench_al = bench.reindex(c.index, method="ffill")
    stock_ret = c.iloc[-1] / c.iloc[-1 - r.rs_weeks] - 1
    bench_ret = bench_al.iloc[-1] / bench_al.iloc[-1 - r.rs_weeks] - 1
    ratio = c / bench_al
    mansfield = float(ratio.iloc[-1] / ratio.rolling(52, min_periods=30).mean().iloc[-1] - 1)
    rs_ok = stock_ret > bench_ret and mansfield > 0

    daily_value_cr = float((c * v).iloc[-10:].median() / 5 / CRORE)
    liquid = daily_value_cr >= r.min_daily_value_cr

    hi52 = float(h.iloc[-52:].max())
    lo52 = float(df["Low"].iloc[-52:].min())

    checks = {
        "above_30w": bool(above_sma),
        "30w_rising": bool(sma_rising),
        "was_stage1_flat": bool(was_flat),
        "fresh_breakout": bool(fresh_breakout and still_above_brk),
        "volume_surge": bool(vol_ok),
        "rel_strength": bool(rs_ok),
        "liquid": bool(liquid),
    }
    return {
        "symbol": sym,
        "close": round(last_c, 2),
        "week": c.index[-1].date().isoformat(),
        **checks,
        "tech_score": sum(checks.values()),
        "breakout_level": round(brk_level, 2),
        "weeks_since_breakout": weeks_since_brk,
        "pct_above_breakout": round(extension * 100, 1),
        "extended": bool(extension > r.max_extension),
        "breakout_vol_x": round(vol_ratio, 2) if not math.isnan(vol_ratio) else None,
        "ret_26w_%": round(stock_ret * 100, 1),
        "nifty500_26w_%": round(bench_ret * 100, 1),
        "mansfield_rs": round(mansfield, 3),
        "pct_from_52w_high": round((last_c / hi52 - 1) * 100, 1),
        "pct_above_52w_low": round((last_c / lo52 - 1) * 100, 1),
        "sma30": round(float(sma.iloc[-1]), 2),
        "daily_value_cr": round(daily_value_cr, 1),
    }


def is_stage2_candidate(row: dict) -> bool:
    core = ("above_30w", "30w_rising", "fresh_breakout", "rel_strength", "liquid")
    return all(row[k] for k in core)


# ----------------------------------------------------------------------------- NSE shareholding / pledge
class NseClient:
    """Reads promoter holding and pledge from NSE's quarterly shareholding-pattern filings.

    Uses the JSON behind nseindia.com's shareholding page plus the XBRL filing it links to.
    These are not an official API: NSE may change or block them, in which case the
    screener carries on and marks pledge as missing.
    """

    HOME = "https://www.nseindia.com/companies-listing/corporate-filings-shareholding-pattern"
    MASTER = "https://www.nseindia.com/api/corporate-share-holdings-master?index=equities&symbol={sym}"
    UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")

    def __init__(self) -> None:
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": self.UA, "Accept": "application/json, text/plain, */*",
                               "Accept-Language": "en-US,en;q=0.9", "Referer": self.HOME})
        self.limit = RateLimiter(2)
        self.ok = self._warm()

    def _warm(self) -> bool:
        try:
            return self.s.get(self.HOME, timeout=20).status_code == 200
        except requests.RequestException:
            return False

    def _get(self, url: str) -> requests.Response:
        for attempt in range(3):
            self.limit.wait()
            resp = self.s.get(url, timeout=30)
            if resp.status_code in (401, 403):   # cookies expired: refresh once and retry
                self._warm()
                continue
            return resp
        return resp

    @staticmethod
    def _xbrl_value(xml: str, tag: str, context: str) -> float | None:
        import re
        m = re.search(r'<in-bse-shp:%s\b[^>]*contextRef="%s"[^>]*>([^<]*)<' % (tag, context), xml)
        try:
            return float(m.group(1)) if m else None
        except ValueError:
            return None

    def shareholding(self, sym: str) -> dict:
        out = {"promoter_%": None, "promoter_chg_1y": None, "pledge_%_of_promoter": None,
               "encumbered_%_of_promoter": None, "shp_quarter": None}
        if not self.ok:
            return out
        try:
            rows = self._get(self.MASTER.format(sym=requests.utils.quote(sym, safe=""))).json()
        except (requests.RequestException, ValueError):
            return out
        by_date: dict[date, dict] = {}
        for x in rows if isinstance(rows, list) else []:
            if not isinstance(x, dict) or x.get("pr_and_prgrp") in (None, "", "-"):
                continue
            try:
                d = datetime.strptime(x["date"], "%d-%b-%Y").date()
                float(x["pr_and_prgrp"])
            except (KeyError, TypeError, ValueError):
                continue
            by_date.setdefault(d, x)  # newest filing for a quarter comes first
        if not by_date:
            return out
        latest_d = max(by_date)
        latest = by_date[latest_d]
        out["shp_quarter"] = latest_d.isoformat()
        out["promoter_%"] = float(latest["pr_and_prgrp"])
        year_ago = [d for d in by_date if 330 <= (latest_d - d).days <= 400]
        if year_ago:
            out["promoter_chg_1y"] = round(out["promoter_%"] - float(by_date[max(year_ago)]["pr_and_prgrp"]), 2)
        if latest.get("xbrl"):
            try:
                self.limit.wait()
                xml = self.s.get(latest["xbrl"], timeout=30).text
                ctx = "ShareholdingOfPromoterAndPromoterGroup_ContextI"
                pl = self._xbrl_value(xml, "EncumberedShareUnderPledgedAsPercentageOfTotalNumberOfShares", ctx)
                en = self._xbrl_value(xml, "EncumberedSharesHeldAsPercentageOfTotalNumberOfShares", ctx)
                import re
                flag = re.search(r'<in-bse-shp:WhetherAnySharesHeldByPromotersAreEncumberedUnderPledged\b[^>]*>\s*(true|false)\s*<',
                                 xml, re.I)
                no_pledge = (flag and flag.group(1).lower() == "false") or out["promoter_%"] == 0
                if pl is None and no_pledge:
                    pl = 0.0
                if en is None and no_pledge and "AreEncumberedOtherThanByWayOfPledgeOrNDU" in xml \
                        and not re.search(r'AreEncumbered(?:UnderNonDisposalUndertaking|OtherThanByWayOfPledgeOrNDU)\b[^>]*>\s*true', xml, re.I):
                    en = 0.0
                out["pledge_%_of_promoter"] = round(pl * 100, 2) if pl is not None else None
                out["encumbered_%_of_promoter"] = round(en * 100, 2) if en is not None else None
            except requests.RequestException:
                pass
        return out


# ----------------------------------------------------------------------------- fundamentals
def fundamentals(sym: str, r: Rules, nse: NseClient | None = None) -> dict:
    out = {"sector": None, "industry": None, "mcap_cr": None, "pe": None,
           "rev_growth_q_%": None, "profit_growth_q_%": None, "roe_%": None,
           "debt_equity": None, "roa_%": None, "insider_%": None,
           "fund_pass": None, "fund_notes": ""}
    shp = nse.shareholding(sym) if nse else {}
    out.update(shp)
    try:
        info = yf.Ticker(f"{sym}.NS").info or {}
    except Exception as e:  # network / rate limit
        out["fund_notes"] = f"fetch failed: {e.__class__.__name__}"
        return out

    def pct(x):
        return round(x * 100, 1) if isinstance(x, (int, float)) else None

    mcap = info.get("marketCap")
    de = info.get("debtToEquity")  # Yahoo reports as a percentage (45 = 0.45)
    out.update({
        "sector": info.get("sector"),
        "industry": info.get("industry"),
        "mcap_cr": round(mcap / CRORE) if mcap else None,
        "pe": round(info["trailingPE"], 1) if isinstance(info.get("trailingPE"), (int, float)) else None,
        "rev_growth_q_%": pct(info.get("revenueGrowth")),
        "profit_growth_q_%": pct(info.get("earningsGrowth") or info.get("earningsQuarterlyGrowth")),
        "roe_%": pct(info.get("returnOnEquity")),
        "debt_equity": round(de / 100, 2) if isinstance(de, (int, float)) else None,
        "roa_%": pct(info.get("returnOnAssets")),
        "insider_%": pct(info.get("heldPercentInsiders")),
    })

    is_bank = "bank" in (out["industry"] or "").lower()
    fails, missing = [], []

    def need(name, val, ok):
        if val is None:
            missing.append(name)
        elif not ok(val):
            fails.append(name)

    need("mcap", out["mcap_cr"], lambda x: x >= r.min_mcap_cr)
    need("rev_growth", out["rev_growth_q_%"], lambda x: x >= r.min_rev_growth * 100)
    need("profit_growth", out["profit_growth_q_%"], lambda x: x >= r.min_profit_growth * 100)
    if is_bank:
        need("roa", out["roa_%"], lambda x: x >= r.min_bank_roa * 100)
    else:
        need("roe", out["roe_%"], lambda x: x >= r.min_roe * 100)
        need("debt_equity", out["debt_equity"], lambda x: x <= r.max_de)
    if nse:
        need("pledge", out.get("pledge_%_of_promoter"), lambda x: x <= r.max_pledge_pct)

    out["fund_pass"] = not fails and len(missing) <= 1
    notes = []
    if is_bank:
        notes.append("bank: ROA used, check GNPA/NNPA")
    if fails:
        notes.append("fails: " + ", ".join(fails))
    if missing:
        notes.append("missing: " + ", ".join(missing))
    if out.get("promoter_%") == 0:
        notes.append("no promoter group (professionally run)")
    chg = out.get("promoter_chg_1y")
    if chg is not None and chg <= -r.max_promoter_drop:
        notes.append(f"promoter stake down {abs(chg)} pts in 1y: check why (OFS/QIP/selling)")
    enc, pl = out.get("encumbered_%_of_promoter"), out.get("pledge_%_of_promoter")
    if enc is not None and pl is not None and enc > pl + 1:
        notes.append(f"other encumbrance (NDU etc.) {enc}%")
    if not nse or out.get("pledge_%_of_promoter") is None:
        notes.append("check pledge on Screener.in")
    out["fund_notes"] = "; ".join(notes)
    return out


# ----------------------------------------------------------------------------- dashboard export
DASH_COLUMNS = [
    "symbol", "close", "week", "verdict", "fund_pass", "extended", "was_stage1_flat", "tech_score",
    "breakout_level", "pct_above_breakout", "weeks_since_breakout", "breakout_vol_x",
    "ret_26w_%", "nifty500_26w_%", "mansfield_rs", "pct_from_52w_high", "daily_value_cr",
    "sector", "industry", "mcap_cr", "pe", "rev_growth_q_%", "profit_growth_q_%", "roe_%",
    "debt_equity", "roa_%", "promoter_%", "promoter_chg_1y", "pledge_%_of_promoter", "fund_notes",
]


def write_dashboard_json(path: str, show: pd.DataFrame, full: pd.DataFrame, r: Rules,
                         args: argparse.Namespace, universe_size: int) -> None:
    cols = [c for c in DASH_COLUMNS if c in show.columns]
    rows = json.loads(show[cols].to_json(orient="records"))  # NaN -> null
    for row in rows:
        row["fund_pass"] = bool(row.get("fund_pass")) if row.get("fund_pass") is not None else None
    with open(__file__, encoding="utf-8") as f:
        code = f.read()
    scope = (" ".join(args.symbols) if args.symbols else
             args.tickers_file or args.screener_file or
             ("All NSE EQ stocks" if args.universe == "all" else "Nifty 500"))
    meta = [
        {"label": "Week ending", "value": str(full["week"].max())},
        {"label": "Run at", "value": datetime.now().strftime("%Y-%m-%d %H:%M")},
        {"label": "Universe", "value": f"{scope} ({universe_size} stocks, {len(full)} with enough history)"},
        {"label": "Price data", "value": f"{args.source} (weekly bars from daily candles)"},
        {"label": "Fundamentals", "value": "Yahoo Finance (growth, ROE, debt); NSE shareholding filings (promoter, pledge)"},
        {"label": "Rules", "value": json.dumps(vars(r))},
        {"label": "stage2_screener.py", "value": code, "type": "code", "lang": "python"},
    ]
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"rows": rows, "meta": meta}, f, ensure_ascii=False)


# ----------------------------------------------------------------------------- main
def main() -> None:
    p = argparse.ArgumentParser(description="NSE Stage 2 + fundamentals screener")
    src = p.add_mutually_exclusive_group()
    src.add_argument("--symbols", nargs="+", help="NSE symbols, e.g. MAHABANK COFORGE")
    src.add_argument("--tickers-file", help="text file of NSE symbols (one per line or comma separated)")
    src.add_argument("--screener-file", help="Screener.in export (csv/xlsx) with an 'NSE Code' column")
    p.add_argument("--universe", choices=["nifty500", "all"], default="nifty500",
                   help="stock list when no symbols/file given: nifty500 (default) or all NSE EQ stocks")
    p.add_argument("--include-be", action="store_true",
                   help="with --universe all, also scan BE (trade-to-trade) series stocks")
    p.add_argument("--min-value-cr", type=float,
                   help="minimum median daily traded value in Rs crore (default 5)")
    p.add_argument("--source", choices=["upstox", "dhan", "yahoo"], default="upstox",
                   help="price data source (default: upstox, no login needed for historical candles)")
    p.add_argument("--loose", action="store_true", help="relax the rules for more results")
    p.add_argument("--include-current-week", action="store_true",
                   help="use the still-forming weekly bar (default: completed weeks only on Mon-Fri)")
    p.add_argument("--all-fundamentals", action="store_true",
                   help="fetch fundamentals for every stock, not just Stage 2 candidates (slow)")
    p.add_argument("--no-nse", action="store_true",
                   help="skip NSE promoter-holding / pledge lookup")
    p.add_argument("--out", help="output .xlsx path")
    p.add_argument("--json-out", help="also write Stage 2 candidates as dashboard JSON {rows, meta}")
    args = p.parse_args()

    r = loose(Rules()) if args.loose else Rules()
    if args.min_value_cr is not None:
        r.min_daily_value_cr = args.min_value_cr
    pre_approved = False
    if args.symbols:
        universe = [s.upper() for s in args.symbols]
    elif args.tickers_file:
        universe = load_tickers_file(args.tickers_file)
    elif args.screener_file:
        universe = load_screener_file(args.screener_file)
        pre_approved = True
    elif args.universe == "all":
        print("Loading full NSE equity list ...")
        universe = load_all_nse(args.include_be)
    else:
        print("Loading Nifty 500 list from NSE ...")
        universe = load_nifty500()
    universe = list(dict.fromkeys(universe))
    print(f"Universe: {len(universe)} stocks")

    print(f"Downloading prices from {args.source} ...")
    if args.source == "upstox":
        prices, bench = download_upstox(universe)
    elif args.source == "dhan":
        prices, bench = download_dhan(universe)
    else:
        bench = download_benchmark()
        prices = download_weekly(universe)
    if args.source != "yahoo":
        gap = [s for s in universe if s not in prices]
        if gap:
            print(f"  filling {len(gap)} missing stocks from Yahoo ...")
            prices.update(download_weekly(gap))
    missing = sorted(set(universe) - set(prices))
    if missing:
        print(f"  no data for {len(missing)}: {', '.join(missing[:15])}{' ...' if len(missing) > 15 else ''}")

    bench = drop_partial_week(bench.to_frame("Close"), args.include_current_week)["Close"]
    rows = []
    for sym, df in prices.items():
        df = drop_partial_week(df, args.include_current_week)
        res = technical(sym, df, bench, r)
        if res:
            res["stage2_candidate"] = is_stage2_candidate(res)
            rows.append(res)
    if not rows:
        sys.exit("No stocks had enough history to scan.")
    tech = pd.DataFrame(rows)
    cands = tech[tech["stage2_candidate"]]
    print(f"Stage 2 candidates (technical): {len(cands)} of {len(tech)}")

    fetch = tech["symbol"] if args.all_fundamentals else cands["symbol"]
    funds = []
    nse = None
    if not args.no_nse and len(fetch):
        nse = NseClient()
        print("NSE shareholding/pledge lookup:", "on" if nse.ok else "unavailable (NSE blocked the request)")
        if not nse.ok:
            nse = None
    if pre_approved:
        print("Screener file given: fundamentals treated as pre-approved.")
    for i, sym in enumerate(fetch, 1):
        print(f"  fundamentals {i}/{len(fetch)}: {sym}", flush=True)
        f = fundamentals(sym, r, nse)
        if pre_approved:
            f["fund_pass"] = True
            f["fund_notes"] = "pre-approved by Screener.in screen; " + f["fund_notes"]
        f["symbol"] = sym
        funds.append(f)
        time.sleep(0.3)
    full = tech.merge(pd.DataFrame(funds), on="symbol", how="left") if funds else tech

    def verdict(row) -> str:
        if not row["stage2_candidate"]:
            return "-"
        if row.get("fund_pass") is True:
            v = "STAGE 2 + STRONG FUNDA"
        elif row.get("fund_pass") is False:
            v = "Stage 2, weak/unknown funda"
        else:
            v = "Stage 2, funda not checked"
        if row["extended"]:
            v += " (extended: wait for retest)"
        elif row["was_stage1_flat"]:
            v += " (just entering)"
        return v

    full["verdict"] = full.apply(verdict, axis=1)
    full["_rank"] = (
        full["stage2_candidate"].astype(int) * 100
        + (full.get("fund_pass") == True).astype(int) * 50  # noqa: E712
        + full["tech_score"] * 5
        - full["extended"].astype(int) * 10
        + full["mansfield_rs"].clip(-1, 1) * 10
    )
    full = full.sort_values("_rank", ascending=False).drop(columns="_rank")

    show = full[full["stage2_candidate"]]
    cols = ["symbol", "close", "verdict", "tech_score", "breakout_level", "pct_above_breakout",
            "breakout_vol_x", "ret_26w_%", "mansfield_rs"]
    for c in ("rev_growth_q_%", "profit_growth_q_%", "roe_%", "debt_equity", "pe",
              "promoter_%", "promoter_chg_1y", "pledge_%_of_promoter"):
        if c in show.columns:
            cols.append(c)
    print()
    if show.empty:
        print("No Stage 2 candidates this week. Try --loose.")
    else:
        with pd.option_context("display.width", 250, "display.max_columns", 30, "display.max_rows", 200):
            print(show[cols].to_string(index=False))

    out = args.out or f"stage2_results_{datetime.now():%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        show.to_excel(xw, sheet_name="Stage2 candidates", index=False)
        full.to_excel(xw, sheet_name="All scanned", index=False)
        pd.DataFrame(
            [{"rule": k, "value": v} for k, v in vars(r).items()]
            + [{"rule": "data_week", "value": full["week"].max()}]
        ).to_excel(xw, sheet_name="Rules", index=False)
    print(f"\nSaved: {out}")
    if args.json_out:
        write_dashboard_json(args.json_out, show, full, r, args, len(universe))
        print(f"Saved: {args.json_out}")
    print("Screening only, not investment advice. Check every chart and the latest results before acting.")


if __name__ == "__main__":
    main()
