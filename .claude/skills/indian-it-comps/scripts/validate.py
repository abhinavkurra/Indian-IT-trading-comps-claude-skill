"""Phase 3: deterministic validation of work/<co>/extracted.json.

Checks: (A) JSON schema, (B) every cited value literally appears on its cited page and the
page reference matches the manifest, (C) accounting identities and cross-foots.
Does NOT read golden values (the extractor runs this script; golden scoring is score_golden.py).

Usage:  python validate.py <co_id>
Output: work/<co_id>/validation.json; stdout = FAIL/WARN lines + one summary line. Exit 1 on any FAIL.
"""
from __future__ import annotations

import json
import re
import sys

from common import SKILL_DIR, co_work_dir, load_yaml, read_json, write_json

UNIT_DIVISOR = {"million": 1e6, "crore": 1e7, "lakh": 1e5, "thousand": 1e3}
NUM_TOKEN = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?")


class Checker:
    def __init__(self):
        self.results = []

    def add(self, name, ok, detail="", level="FAIL"):
        self.results.append({"check": name, "ok": bool(ok), "level": level, "detail": detail})

    def eq(self, name, a, b, tol, level="FAIL"):
        if a is None or b is None:
            self.add(name, False, f"missing operand ({a!r} vs {b!r})", level)
            return
        self.add(name, abs(a - b) <= tol, f"{a:,.4f} vs {b:,.4f} (diff {a - b:,.4f}, tol {tol})", level)


def v(x):
    return None if x is None else x["value"]


def page_numbers(text: str) -> list[float]:
    out = []
    for t in NUM_TOKEN.findall(text):
        s = t.strip("()").lstrip("-").replace(",", "")
        try:
            out.append(float(s))
        except ValueError:
            pass
    return out


def walk_cites(obj, path=""):
    if isinstance(obj, dict):
        if "value" in obj and "uid" in obj:
            yield path, obj
        for k, x in obj.items():
            if k != "value":
                yield from walk_cites(x, f"{path}.{k}" if path else k)
    elif isinstance(obj, list):
        for i, x in enumerate(obj):
            yield from walk_cites(x, f"{path}[{i}]")


def validate(co_id: str) -> Checker:
    wd = co_work_dir(co_id)
    c = Checker()
    try:
        ex = read_json(wd / "extracted.json")
    except (FileNotFoundError, json.JSONDecodeError) as e:
        c.add("load", False, str(e))
        return c

    # ---- (A) schema
    import jsonschema
    schema = read_json(SKILL_DIR / "reference" / "extracted.schema.json")
    errs = sorted(jsonschema.Draft202012Validator(schema).iter_errors(ex), key=lambda e: list(e.path))
    for e in errs[:15]:
        c.add("schema", False, f"{'/'.join(map(str, e.path))}: {e.message[:160]}")
    if errs:
        return c
    c.add("schema", True)

    manifest = read_json(wd / "manifest.json")
    pages = {p["uid"]: p for p in manifest["pages"]}
    texts = {}

    # ---- (B) citations
    for path, cite in walk_cites(ex):
        uid = cite["uid"]
        if uid not in pages:
            c.add(f"cite:{path}", False, f"uid {uid} not in manifest pages")
            continue
        if cite.get("printed_page") != pages[uid]["printed_page"]:
            c.add(f"cite:{path}", False, f"printed_page {cite.get('printed_page')} != manifest {pages[uid]['printed_page']}")
            continue
        if uid not in texts:
            texts[uid] = page_numbers((wd / "pages" / f"{uid}.txt").read_text(encoding="utf-8"))
        target = abs(cite["value"])
        found = target == 0 or any(abs(n - target) < 1e-6 for n in texts[uid])
        if not found and path.startswith("shares.") and target >= 1e5:
            # share counts are sometimes printed in million / lakh / crore ("157.75 (No. in Million)")
            found = any(n and abs(n * s - target) <= 0.0005 * target for n in texts[uid] for s in (1e6, 1e5, 1e7))
        c.add(f"cite:{path}", found, "" if found else f"value {cite['value']} not found on {uid}")

    # ---- (C) identities
    pl, bs = ex["pl"], ex["bs"]
    decimals = any(isinstance(x, float) and not float(x).is_integer()
                   for x in [v(pl["revenue_from_operations"]), v(bs["total_assets"]), v(pl["profit_for_year"])])
    def tol(n_terms):           # rounding tolerance for sums of printed (rounded) figures
        return (0.006 if decimals else 0.6) * max(n_terms, 2)

    c.add("unit_matches_manifest", ex["unit_as_printed"] == manifest["reporting_unit"],
          f"{ex['unit_as_printed']} vs {manifest['reporting_unit']}")
    bsd = str(load_yaml("peers.yaml")["balance_sheet_date"])
    c.add("period_end", ex["period_end"] == bsd, f"{ex['period_end']} vs {bsd}")

    c.eq("bs: total assets = total equity & liabilities", v(bs["total_assets"]), v(bs["total_equity_and_liabilities"]), tol(2))
    if bs["equity_owners"] is not None:
        c.eq("bs: total equity = owners + NCI", v(bs["total_equity"]),
             v(bs["equity_owners"]) + (v(bs["nci"]) or 0), tol(2))

    exp = ex["expense_lines"]
    c.eq("pl: sum(expense_lines) = total expenses", sum(x["value"] for x in exp), v(pl["total_expenses"]), tol(len(exp)))
    exp_vals = [x["value"] for x in exp]
    for k in ("finance_costs", "depreciation_amortisation"):
        c.add(f"pl: {k} is one of expense_lines", any(abs(x - v(pl[k])) < 1e-6 for x in exp_vals))
    if pl["fx_line_charge"] is not None:
        c.add("pl: fx_line is one of expense_lines", any(abs(abs(x) - abs(v(pl["fx_line_charge"]))) < 1e-6 for x in exp_vals))

    fxoi = pl.get("fx_in_other_income")
    if fxoi is not None:
        c.add("pl: fx_in_other_income cited from the Other Income note",
              "note:other_income" in pages.get(fxoi["uid"], {}).get("why", []), fxoi["uid"])
        c.add("pl: |fx_in_other_income| is not the whole Other Income line",
              abs(fxoi["value"]) != abs(v(pl["other_income"])) or v(pl["other_income"]) == 0, level="WARN")

    c.eq("pl: revenue + other income - total expenses + share of equity investees = PBEIT",
         v(pl["revenue_from_operations"]) + v(pl["other_income"]) - v(pl["total_expenses"])
         + (v(pl.get("share_of_equity_investees")) or 0),
         v(pl["profit_before_exceptional_and_tax"]), tol(3))
    c.eq("pl: PBEIT - exceptional = PBT", v(pl["profit_before_exceptional_and_tax"]) - (v(pl["exceptional_charge"]) or 0),
         v(pl["profit_before_tax"]), tol(2))
    pat_cont = v(pl["profit_continuing_ops"]) if pl["profit_continuing_ops"] is not None else v(pl["profit_for_year"])
    c.eq("pl: PBT - tax = PAT (continuing)", v(pl["profit_before_tax"]) - v(pl["total_tax_expense"]), pat_cont, tol(2))
    if pl["profit_continuing_ops"] is not None:
        c.eq("pl: continuing + discontinued = profit for year",
             v(pl["profit_continuing_ops"]) + (v(pl["profit_discontinued_ops"]) or 0), v(pl["profit_for_year"]), tol(2))
    elif pl["profit_discontinued_ops"] is not None:
        c.add("pl: discontinued without continuing split", False, "profit_discontinued_ops set but profit_continuing_ops null")
    c.eq("pl: owners + NCI = profit for year", v(pl["pat_owners"]) + (v(pl["pat_nci"]) or 0), v(pl["profit_for_year"]), tol(2))

    cc = ex["cash_components"]
    for bucket, face in (("cash_and_cash_equivalents", bs["cash_and_cash_equivalents"]), ("other_bank_balances", bs["other_bank_balances"])):
        items = [x["value"] for x in cc if x["bucket"] == bucket]
        if face is None and not items:
            continue
        c.eq(f"cash: sum({bucket} components) = face", sum(items), v(face) or 0, tol(len(items)))
    for bucket, face in (("current", bs["current_investments"]), ("non_current", bs["non_current_investments"])):
        items = [x["value"] for x in ex["investment_components"] if x["bucket"] == bucket]
        if face is None and not items:
            continue
        c.eq(f"investments: sum({bucket} components) = face", sum(items), v(face) or 0, tol(len(items)))
    total_cash_lines = sum(x["value"] for x in cc if not x["restricted"])
    carve = sum(x["value"] for x in ex.get("restricted_carveouts", []))
    c.add("cash: carve-outs <= unrestricted cash lines", carve <= total_cash_lines + 1e-6, f"{carve} vs {total_cash_lines}")

    sh = ex["shares"]
    div = UNIT_DIVISOR[ex["unit_as_printed"]]
    implied = v(sh["shares_outstanding_year_end"]) * v(sh["face_value_inr"]) / div
    implied_net = (v(sh["shares_outstanding_year_end"]) - (v(sh["treasury_shares"]) or 0)) * v(sh["face_value_inr"]) / div
    esc = v(bs["equity_share_capital"])
    t_ = max(0.005 * esc, tol(1))
    # face share capital may be gross (issued) or net of ESOP-trust shares - both presentations are valid
    c.add("shares: count x face value = equity share capital (gross or net of treasury)",
          abs(implied - esc) <= t_ or abs(implied_net - esc) <= t_,
          f"gross {implied:,.4f} / net {implied_net:,.4f} vs printed {esc:,.4f}")
    eps_impl = v(pl["pat_owners"]) * div / v(sh["shares_outstanding_year_end"])
    c.add("eps: PAT owners / year-end shares ~ basic EPS (+-5%)", abs(eps_impl / v(pl["eps_basic"]) - 1) <= 0.05,
          f"implied {eps_impl:.2f} vs printed {v(pl['eps_basic']):.2f}", level="WARN")
    c.add("eps: diluted <= basic", v(pl["eps_diluted"]) <= v(pl["eps_basic"]) + 1e-9,
          f"{v(pl['eps_diluted'])} vs {v(pl['eps_basic'])}", level="WARN")
    for i, p in enumerate(ex["esop_plans"]):
        c.add(f"esop[{i}] page in manifest", p["uid"] in pages and pages[p["uid"]]["printed_page"] == p["printed_page"],
              f"{p['uid']} / {p['printed_page']}")
    return c


def main(argv):
    rc = 0
    for co_id in argv:
        c = validate(co_id)
        write_json(co_work_dir(co_id) / "validation.json", c.results)
        fails = [r for r in c.results if not r["ok"] and r["level"] == "FAIL"]
        warns = [r for r in c.results if not r["ok"] and r["level"] == "WARN"]
        for r in fails + warns:
            print(f"{r['level']} {r['check']}: {r['detail']}")
        n_fail_checks = [r for r in c.results if r["level"] == "FAIL"]
        print(f"[{co_id}] {'PASS' if not fails else 'FAIL'} checks={len(n_fail_checks) - len(fails)}/{len(n_fail_checks)} warnings={len(warns)}")
        rc |= bool(fails)
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
