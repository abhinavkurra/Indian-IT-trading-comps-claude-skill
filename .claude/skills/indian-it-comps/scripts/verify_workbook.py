"""Phase 5c: recalculate output/trading_comps.xlsx independently and diff against the Python shadow.

Engine: the `formulas` package (pure-Python Excel calculator) - no Excel/LibreOffice needed.
Also asserts every reconciliation check row is ~0 and no cell evaluates to an Excel error.
Usage:  python verify_workbook.py
"""
from __future__ import annotations

import re
import sys

from common import OUTPUT_DIR, ROOT, WORK_DIR, read_json

TOL = 0.01


def norm(addr: str) -> str:
    sheet, cell = addr.rsplit("!", 1)
    return f"{sheet.strip(chr(39)).upper()}!{cell.replace('$', '').upper()}"


def recalc(path):
    import formulas
    sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    out = {}
    for k, val in sol.items():
        m = re.match(r"'?\[[^\]]+\](.+?)'?!([A-Z]+\d+)$", k)       # single cells only
        if not m:
            continue
        x = val.value[0, 0] if hasattr(val, "value") else val
        out[f"{m.group(1).upper()}!{m.group(2)}"] = x.item() if hasattr(x, "item") else x
    return out


def main():
    shadow, cmap = read_json(WORK_DIR / "comps_shadow.json"), read_json(WORK_DIR / "cellmap.json")
    from pathlib import Path
    path = ROOT / (shadow.get("workbook") or OUTPUT_DIR / "trading_comps.xlsx")   # stored root-relative
    vals = recalc(path)
    fails, n = [], 0

    def cmp(name, addr, expected):
        nonlocal n
        n += 1
        got = vals.get(norm(addr))
        ok = (got == expected) if isinstance(expected, str) else (
            isinstance(got, (int, float)) and abs(got - expected) <= TOL * max(1, abs(expected)) / 100 + TOL)
        if not ok:
            fails.append(f"  DIFF {name} @ {addr}: workbook {got!r} vs shadow {expected!r}")

    keymap = {"ev_rev": "ev_rev", "ev_ebitda": "ev_ebitda", "pe_adj": "pe_adj", "pe_rep": "pe_rep",
              "ev": "ev", "eqv": "equity_value"}
    for co, s in shadow["companies"].items():
        for k, sk in keymap.items():
            cmp(f"{co}.{k}", cmap[f"comps.{k}.{co}"], s[sk])
        for k, sk in (("ebitda", "ebitda"), ("adj_pat", "adj_pat"), ("revenue", "revenue"), ("cash_total", "cash"),
                      ("debt_like", "debt_like"), ("leases", "leases"), ("borrowings", "borrowings")):
            cmp(f"{co}.std.{k}", cmap[f"std.{k}.{co}"], s[sk])
        for chk in ("ebitda_chk", "cash_chk", "inv_chk"):
            cmp(f"{co}.{chk} = 0", cmap[f"std.{chk}.{co}"], 0.0)
    for k, by in shadow["stats"].items():
        for stat, x in by.items():
            cmp(f"stats.{stat}.{k}", cmap[f"stats.{stat}.{k}"], x)

    for k, x in shadow.get("cells", {}).items():      # football field, NCI and other extra cells
        cmp(k, cmap[k], x)

    errors = [f"  ERROR {a}: {x}" for a, x in vals.items() if isinstance(x, str) and x.startswith("#")]
    print("\n".join(fails + errors[:20]))
    status = "PASS" if not fails and not errors else "FAIL"
    print(f"[workbook] {status} {n - len(fails)}/{n} shadow matches, {len(errors)} Excel errors, {len(vals)} cells recalculated")
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
