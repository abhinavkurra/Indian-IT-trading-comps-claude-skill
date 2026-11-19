"""Phase 5a: market data - price (yfinance), basic shares (AR-anchored, cross-checked), TSM dilution.

Rules
- Price: last non-NaN unadjusted daily Close (never intraday 'currentPrice'); date recorded.
- Basic shares: AR year-end count - treasury + included adjustments from config/share_bridge.yaml
  (perimeter must match the FY financials). yfinance count is a reconciliation check only:
  residual vs AR + all adjustments beyond SHARE_TOL -> BLOCKED.
- Dilution: treasury stock method per plan: n * max(0, 1 - K/P). RSU/PSU at nominal K included.

Usage:  python market_data.py <co_id> [...]
Output: work/<co_id>/market.json ; exit 1 if any company is BLOCKED.
"""
from __future__ import annotations

import datetime as dt
import math
import sys
import warnings

from common import co_work_dir, load_yaml, read_json, write_json

SHARE_TOL = 0.005


def fetch_yf(ticker: str, retries: int = 3) -> dict:
    """Last valid close + yfinance share count. Retries, then falls back NSE (.NS) -> BSE (.BO)."""
    import time
    import yfinance as yf
    warnings.filterwarnings("ignore")
    last_err = None
    for tkr in [ticker] + ([ticker.replace(".NS", ".BO")] if ticker.endswith(".NS") else []):
        for attempt in range(retries):
            try:
                tk = yf.Ticker(tkr)
                h = tk.history(period="1y", auto_adjust=False)
                h = h[h["Close"].notna()]
                if h.empty:
                    raise RuntimeError(f"no valid closes for {tkr}")
                try:
                    shares = tk.fast_info.get("shares")
                except Exception:
                    shares = None
                return {"ticker_used": tkr, "price": float(h["Close"].iloc[-1]),
                        "price_date": h.index[-1].date().isoformat(),
                        "low_52w": float(h["Close"].min()), "high_52w": float(h["Close"].max()),
                        "range_from": h.index[0].date().isoformat(),
                        "shares_yf": int(shares) if shares else None}
            except Exception as e:          # transient Yahoo failures are common
                last_err = e
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"yfinance failed for {ticker}: {last_err}")


def tsm(plans: list[dict], price: float) -> tuple[float, list[dict]]:
    detail, total = [], 0.0
    for p in plans:
        n, k = p["outstanding_year_end"], p["waep_inr"] or 0.0
        add = n * max(0.0, 1 - k / price) if n else 0.0
        total += add
        detail.append({"plan": p["plan"], "instrument": p.get("instrument"), "outstanding": n,
                       "waep_inr": k, "in_the_money": k < price, "dilutive_shares": round(add, 1)})
    return total, detail


def run(co_id: str, peer: dict, bridge: dict) -> dict:
    ex = read_json(co_work_dir(co_id) / "extracted.json")
    ar = ex["shares"]["shares_outstanding_year_end"]
    treasury = (ex["shares"]["treasury_shares"] or {}).get("value", 0) or 0
    if treasury:
        # Some ARs print the share count already net of ESOP-trust shares (e.g. Zensar). Decide from the
        # face share capital: if count x FV already equals it, the count is net -> do not deduct again.
        div = {"million": 1e6, "crore": 1e7, "lakh": 1e5, "thousand": 1e3}[ex["unit_as_printed"]]
        fv, esc = ex["shares"]["face_value_inr"]["value"], ex["bs"]["equity_share_capital"]["value"]
        net_err = abs(ar["value"] * fv / div - esc)
        gross_err = abs((ar["value"] - treasury) * fv / div - esc)
        if net_err <= gross_err:
            treasury = 0
    adj = (bridge or {}).get(co_id, {}).get("adjustments", [])
    basic = ar["value"] - treasury + sum(a["shares"] for a in adj if a["include"])
    explained = ar["value"] - treasury + sum(a["shares"] for a in adj)
    m = fetch_yf(peer["nse"])
    out = {"co_id": co_id, "ticker": m["ticker_used"], "price_inr": m["price"], "price_date": m["price_date"],
           "price_source": "yfinance history() unadjusted Close (last valid)",
           "shares_ar_year_end": ar["value"], "shares_ar_cite": f"printed p.{ar['printed_page']} ({ar['uid']})",
           "treasury_shares": treasury, "share_adjustments": adj, "basic_shares": basic,
           "shares_yf": m["shares_yf"], "fetched_at": dt.datetime.now().isoformat(timespec="seconds")}
    resid = (m["shares_yf"] / explained - 1) if m["shares_yf"] else None
    out["yf_unexplained_residual"] = None if resid is None else round(resid, 5)
    cfg = (bridge or {}).get(co_id, {})
    if cfg.get("basic_shares_source") == "yfinance" and m["shares_yf"]:
        # explicit user override: valuation share count = yfinance; gap vs AR kept visible as a flag
        basic = m["shares_yf"]
        out.update(basic_shares=basic, shares_override="yfinance (user instruction)",
                   shares_override_reason=cfg.get("reason", ""))
    elif resid is None or abs(resid) > SHARE_TOL:
        out.update(status="BLOCKED", reason=(
            f"yfinance {m['shares_yf']:,} vs AR+bridge {explained:,.0f} "
            f"({'n/a' if resid is None else f'{resid:+.2%}'}) exceeds {SHARE_TOL:.1%}: document the "
            f"issuance/buyback in config/share_bridge.yaml (with AR or exchange-filing source)"))
        write_json(co_work_dir(co_id) / "market.json", out)
        return out
    out.update(low_52w=m["low_52w"], high_52w=m["high_52w"], range_52w_from=m["range_from"],
               range_52w_basis="daily unadjusted closes, last 1y (yfinance)")
    subs = []
    for sub in peer.get("listed_subsidiaries") or []:
        if not sub.get("nse") or sub.get("nci_pct") is None:
            continue
        sm = fetch_yf(sub["nse"])
        if not sm["shares_yf"]:
            raise RuntimeError(f"no share count for listed subsidiary {sub['nse']}")
        subs.append({"name": sub["name"], "ticker": sm["ticker_used"], "price_inr": sm["price"],
                     "price_date": sm["price_date"], "shares": sm["shares_yf"], "nci_pct": sub["nci_pct"],
                     "nci_book": sub["nci_book"], "source": sub.get("source", "")})
    out["listed_subsidiaries"] = subs
    dil, detail = tsm(ex["esop_plans"], m["price"])
    out.update(status="OK", tsm_dilutive_shares=round(dil, 1), tsm_detail=detail, diluted_shares=basic + dil)
    write_json(co_work_dir(co_id) / "market.json", out)
    return out


def main(argv):
    peers = {p["co_id"]: p for p in load_yaml("peers.yaml")["peers"]}
    bridge = load_yaml("share_bridge.yaml") or {}
    rc = 0
    for co in argv:
        r = run(co, peers[co], bridge)
        if r["status"] == "OK":
            print(f"[{co}] OK price {r['price_inr']:.2f} ({r['price_date']}) basic {r['basic_shares']:,} "
                  f"(yf residual {r['yf_unexplained_residual']:+.2%}) +TSM {r['tsm_dilutive_shares']:,.0f} -> diluted {r['diluted_shares']:,.0f}"
                  )
        else:
            print(f"[{co}] BLOCKED: {r['reason']}")
            rc = 1
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
