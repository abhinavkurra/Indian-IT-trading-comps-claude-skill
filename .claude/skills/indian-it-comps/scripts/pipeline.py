"""Deterministic pipeline driver for the indian-it-comps skill.

  python pipeline.py prep   [co ...]   index + slice every peer whose PDF is present (cached by mtime),
                                       then print which companies need LLM extraction (TO_EXTRACT line).
  python pipeline.py finish [co ...]   validate -> market data -> build workbook -> verify workbook.
  python pipeline.py status            one line per peer.

Only compact summaries are printed so the orchestrating conversation stays small.
State: work/pipeline_status.json
"""
from __future__ import annotations

import contextlib
import io
import json
import sys

from common import INPUT_DIR, SKILL_DIR, WORK_DIR, load_yaml, read_json, write_json

SCRIPTS = SKILL_DIR / "scripts"


def peers(sel=None):
    ps = load_yaml("peers.yaml")["peers"]
    return [p for p in ps if not sel or p["co_id"] in sel]


def mtime(p):
    return p.stat().st_mtime if p.exists() else 0.0


def quiet(fn, *a, **k):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        res = fn(*a, **k)
    return res, buf.getvalue()


def validation_status(co):
    import validate
    c = validate.validate(co)
    fails = [r for r in c.results if not r["ok"] and r["level"] == "FAIL"]
    hard = [r for r in c.results if r["level"] == "FAIL"]
    write_json(WORK_DIR / co / "validation.json", c.results)
    return (not fails), f"{len(hard) - len(fails)}/{len(hard)}", fails


def prep(sel):
    import index_pdf
    import slice_pages
    index_code = max(mtime(SCRIPTS / f) for f in ("index_pdf.py", "common.py"))
    slice_code = mtime(SCRIPTS / "slice_pages.py")
    status, to_extract = {}, []
    for p in peers(sel):
        co, pdf = p["co_id"], INPUT_DIR / p["pdf"]
        wd, row = WORK_DIR / co, {"name": p["name"]}
        if not pdf.exists():
            row["state"] = f"MISSING_PDF (expected input_pdfs/{p['pdf']})"
            status[co] = row
            continue
        man, pidx = wd / "manifest.json", wd / "page_index.json"
        if mtime(man) > max(mtime(pdf), index_code, slice_code, mtime(pidx)):
            row["slice"] = "cached"
        else:
            if mtime(pidx) > max(mtime(pdf), index_code):      # index still valid: re-slice only
                idx = read_json(pidx)
            else:
                idx, _ = quiet(index_pdf.index_company, co, p["pdf"])
            if idx["problems"]:
                row["state"] = "INDEX_PROBLEMS: " + "; ".join(idx["problems"])
                status[co] = row
                continue
            m, _ = quiet(slice_pages.slice_company, co, p["pdf"])
            row["slice"] = "rebuilt"
            if m["problems"]:
                row["slice_problems"] = m["problems"]
        m = read_json(man)
        row.update(pages=m["n_pages"], approx_tokens=m["approx_tokens"], unit=m["reporting_unit"])
        ex = wd / "extracted.json"
        if ex.exists() and mtime(ex) > mtime(man):
            ok, score, _ = validation_status(co)
            row["extraction"] = f"cached ({'PASS' if ok else 'FAIL'} {score})"
            if not ok:
                to_extract.append(co)
        else:
            row["extraction"] = "needed"
            to_extract.append(co)
        row.setdefault("state", "READY")
        status[co] = row
    write_json(WORK_DIR / "pipeline_status.json", {"prep": status, "to_extract": to_extract})
    for co, r in status.items():
        extra = "" if "pages" not in r else f" | {r['pages']} pages ~{r['approx_tokens'] // 1000}k tok, {r['unit']}, slice {r['slice']}, extraction {r['extraction']}"
        print(f"{co:12s} {r['state']}{extra}" + (f" | slice problems: {r['slice_problems']}" if r.get("slice_problems") else ""))
    print("TO_EXTRACT: " + (" ".join(to_extract) if to_extract else "(none)"))
    return 0


def finish(sel):
    import build_workbook
    import market_data
    import verify_workbook
    bridge = load_yaml("share_bridge.yaml") or {}
    ready, lines, rc = [], [], 0
    for p in peers(sel):
        co = p["co_id"]
        if not (WORK_DIR / co / "extracted.json").exists():
            lines.append(f"{co:12s} SKIPPED (no extraction)")
            continue
        ok, score, fails = validation_status(co)
        if not ok:
            lines.append(f"{co:12s} VALIDATION FAIL {score}: " + "; ".join(f"{f['check']}: {f['detail']}" for f in fails[:4]))
            rc = 1
            continue
        try:
            mk, _ = quiet(market_data.run, co, p, bridge)
        except Exception as e:                       # network / Yahoo failures
            lines.append(f"{co:12s} MARKET DATA ERROR: {e}")
            rc = 1
            continue
        if mk["status"] != "OK":
            lines.append(f"{co:12s} MARKET BLOCKED: {mk['reason']}")
            rc = 1
            continue
        g = SKILL_DIR / "tests" / "golden" / f"{co}.yaml"
        gtxt = ""
        if g.exists():
            import score_golden
            _, out = quiet(score_golden.score, co)
            gtxt = " | " + out.strip().splitlines()[-1].split("] ")[-1]
        lines.append(f"{co:12s} OK checks {score} | price {mk['price_inr']:.2f} ({mk['price_date']}) | "
                     + (f"shares = yfinance OVERRIDE (vs AR {mk['yf_unexplained_residual']:+.2%})" if mk.get("shares_override")
                      else f"shares yf residual {mk['yf_unexplained_residual']:+.2%}") + gtxt
                     + (" | EXCLUDED FROM STATS" if p.get("exclude_from_stats") else ""))
        ready.append(co)
    print("\n".join(lines))
    if not ready:
        print("BUILD: nothing to build")
        return 1
    _, out = quiet(build_workbook.main, ready)
    print("BUILD: " + out.strip().replace("\n", "\n       "))
    vrc, vout = quiet(verify_workbook.main)
    print("VERIFY: " + vout.strip().splitlines()[-1])
    write_json(WORK_DIR / "pipeline_status.json",
               {**(read_json(WORK_DIR / "pipeline_status.json") if (WORK_DIR / "pipeline_status.json").exists() else {}),
                "finish": {"built": ready, "lines": lines}})
    return rc | vrc


def status():
    s = read_json(WORK_DIR / "pipeline_status.json") if (WORK_DIR / "pipeline_status.json").exists() else {}
    print(json.dumps(s, indent=1)[:3000])
    return 0


if __name__ == "__main__":
    cmd, sel = (sys.argv[1] if len(sys.argv) > 1 else "status"), sys.argv[2:]
    sys.exit({"prep": lambda: prep(sel), "finish": lambda: finish(sel), "status": status}[cmd]())
