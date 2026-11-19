"""Score work/<co>/extracted.json against tests/golden/<co>.yaml (hand-keyed ground truth).

Run by the orchestrator / developer only - never exposed to the extractor subagent.
Usage:  python score_golden.py <co_id>
"""
from __future__ import annotations

import re
import sys

import yaml

from common import SKILL_DIR, co_work_dir, read_json


def _get(d, path):
    for k in path.split("."):
        d = None if d is None else d.get(k)
    return d


def _sum_where(items, rx, **conds):
    hits = [x for x in items if re.search(rx, x["label"], re.I) and all(x.get(k) == val for k, val in conds.items())]
    if not hits:
        return None, None
    return sum(x["value"] for x in hits), hits[0]["printed_page"]


def extracted_values(ex) -> dict:
    """Map golden keys -> (value, printed_page) from extracted.json."""
    out = {}
    simple = {
        "pl.revenue_from_operations": "pl.revenue_from_operations", "pl.other_income": "pl.other_income",
        "pl.finance_costs": "pl.finance_costs", "pl.depreciation_amortisation": "pl.depreciation_amortisation",
        "pl.exceptional_items": "pl.exceptional_charge", "pl.profit_before_tax": "pl.profit_before_tax",
        "pl.total_tax_expense": "pl.total_tax_expense", "pl.profit_for_year": "pl.profit_for_year",
        "pl.pat_attributable_to_owners": "pl.pat_owners", "pl.diluted_eps_inr": "pl.eps_diluted",
        "bs.total_assets": "bs.total_assets", "bs.borrowings_non_current": "bs.borrowings_non_current",
        "bs.borrowings_current": "bs.borrowings_current",
        "bs.lease_liabilities_non_current": "bs.lease_liabilities_non_current",
        "bs.lease_liabilities_current": "bs.lease_liabilities_current",
        "bs.cash_and_cash_equivalents": "bs.cash_and_cash_equivalents",
        "bs.other_bank_balances": "bs.other_bank_balances", "bs.current_investments": "bs.current_investments",
        "bs.non_current_investments": "bs.non_current_investments", "bs.non_controlling_interest": "bs.nci",
        "bs.equity_share_capital": "bs.equity_share_capital",
        "notes.face_value_per_share_inr": "shares.face_value_inr",
        "notes.shares_outstanding_year_end": "shares.shares_outstanding_year_end",
    }
    for g, e in simple.items():
        c = _get(ex, e)
        out[g] = (0, None) if c is None else (c["value"], c["printed_page"])
    exp = ex["expense_lines"]
    out["pl.purchases_equipment_software"] = _sum_where(exp, r"purchase")
    out["pl.change_in_inventories"] = _sum_where(exp, r"change.*inventor")
    out["pl.employee_benefits_expense"] = _sum_where(exp, r"employee")
    out["pl.other_expenses"] = _sum_where(exp, r"^other\s+expenses")
    inv = ex["investment_components"]
    out["notes.current_investments_mutual_funds"] = _sum_where(inv, r"mutual\s+fund", bucket="current")
    out["notes.current_investments_term_deposits"] = _sum_where(inv, r"term\s+deposit", bucket="current")
    out["notes.restricted_unclaimed_dividend"] = _sum_where(ex["cash_components"], r"dividend", restricted=True)
    rc = ex.get("restricted_carveouts", [])
    out["notes.restricted_deposits_vs_guarantees"] = _sum_where(rc, r"guarantee") if rc else (None, None)
    plans = ex["esop_plans"]
    n = sum(p["outstanding_year_end"] for p in plans)
    waep = sum(p["outstanding_year_end"] * (p["waep_inr"] or 0) for p in plans) / n if n else None
    out["notes.esop_options_outstanding_year_end"] = (n, None)
    out["notes.esop_weighted_avg_exercise_price"] = (waep, None)
    return out


def score(co_id: str) -> int:
    golden = yaml.safe_load(open(SKILL_DIR / "tests" / "golden" / f"{co_id}.yaml", encoding="utf-8"))
    ex = read_json(co_work_dir(co_id) / "extracted.json")
    got = extracted_values(ex)
    n_ok = n = 0
    misses = []
    for section in ("pl", "bs", "notes"):
        for key, g in golden[section].items():
            gk = f"{section}.{key}"
            gv, gp = g["value"], g.get("page")
            if gv is None:
                misses.append(f"  BLANK {gk}: golden value not filled in")
                n += 1
                continue
            if "path" in g:                       # generic: dotted path into extracted.json
                c = _get(ex, g["path"])
                ev, ep = (0, None) if c is None else (c["value"], c["printed_page"])
            elif "sum" in g:                      # generic: sum list items matching label regex + fields
                spec = dict(g["sum"])
                ev, ep = _sum_where(ex.get(spec.pop("list"), []), spec.pop("label", "."), **spec)
            else:
                ev, ep = got.get(gk, (None, None))
            tol = 0.01 if key.endswith("exercise_price") else 1e-6
            v_ok = ev is not None and abs(ev - gv) <= tol
            # golden `page` may be a list when the same figure is printed on several pages
            p_ok = gp is None or ep is None or ep in (gp if isinstance(gp, list) else [gp])
            n += 1
            if v_ok and p_ok:
                n_ok += 1
            else:
                misses.append(f"  MISS {gk}: golden {gv} (p{gp}) vs extracted {ev} (p{ep})")
    print("\n".join(misses))
    print(f"[{co_id}] golden match {n_ok}/{n}")
    return 0 if n_ok == n else 1


if __name__ == "__main__":
    sys.exit(max(score(c) for c in sys.argv[1:]))
