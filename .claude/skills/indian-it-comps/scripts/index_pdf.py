"""Phase 1: index an Annual Report PDF without any LLM involvement.

Classifies every printed page (spreads are split into halves) as a face statement /
notes page / other, consolidated vs standalone, and selects the best candidate unit for
each consolidated face statement.

Usage:  python index_pdf.py <co_id> [<co_id> ...]      (pdf name taken from peers.yaml)
Output: work/<co_id>/page_index.json   + one-line summary on stdout
"""
from __future__ import annotations

import re
import sys

from common import INPUT_DIR, co_work_dir, detect_unit, fill_folios, iter_units, load_yaml, write_json

TITLES = {
    "bs": re.compile(r"^(consolidated\s+|standalone\s+)?balance\s+sheet\b", re.I),
    "pl": re.compile(r"^(consolidated\s+|standalone\s+)?statement\s+of\s+profit\s+(and|&)\s+loss\b", re.I),
    "cf": re.compile(r"^(consolidated\s+|standalone\s+)?(statement\s+of\s+cash\s+flows?|cash\s+flow\s+statement)\b", re.I),
    "soce": re.compile(r"^(consolidated\s+|standalone\s+)?statement\s+of\s+changes\s+in\s+equity\b", re.I),
}
NOTES_HDR = re.compile(r"^notes\b.*\bfinancial\s+statements\b", re.I)
AMOUNT = re.compile(r"(?<![\w.])\(?\d{1,3}(?:,\d{2,3})+(?:\.\d+)?\)?(?![\w])|(?<![\w.,])\d+\.\d{2}(?![\w])")
HEAD_LINES = 12   # titles must appear in the first N non-empty lines (filters auditor-report prose)
# Section starts that end the consolidated notes region.
END_RX = re.compile(r"^(independent\s+auditor|form\s+(no\.?\s*)?aoc|notice\s+(of\s+the|is\s+hereby)|"
                    r"business\s+responsibility|standalone\s+financial\s+statements$)", re.I)


def head_lines(text: str) -> list[str]:
    return [l.strip() for l in text.splitlines() if l.strip()][:HEAD_LINES]


def classify(text: str) -> dict:
    heads = head_lines(text)
    kind, scope, score = None, None, 0
    # Notes header wins: note pages often contain sub-headings like "Statement of cash flows".
    # Titles may wrap over two lines ("Notes forming part of Consolidated" / "Financial Statements").
    joined = [f"{a} {b}" for a, b in zip(heads[:6], heads[1:7])]
    for line in heads[:6] + joined:
        if NOTES_HDR.match(line):
            kind = "notes"
            scope = "consolidated" if "consolidated" in line.lower() else (
                "standalone" if "standalone" in line.lower() else None)
            break
    if kind is None:
        for line in heads:
            for k, rx in TITLES.items():
                m = rx.match(line)
                if m and len(line) < 110:
                    kind = k
                    scope = "standalone" if not (m.group(1) or "").lower().startswith("cons") else "consolidated"
                    break
            if kind:
                break
    n_amounts = len(AMOUNT.findall(text))
    end_marker = any(END_RX.search(l) for l in heads[:6])
    if kind in TITLES:
        score = 10 + min(n_amounts, 40) // 4 + (3 if detect_unit(text) else 0)
        if n_amounts < 15:          # a title with no table = narrative mention, not the statement
            score -= 12
    return {"kind": kind, "scope": scope, "score": score, "n_amounts": n_amounts, "_end_marker": end_marker}


def index_company(co_id: str, pdf_name: str) -> dict:
    pdf = INPUT_DIR / pdf_name
    raw = list(iter_units(pdf))
    fill_folios(raw)
    # Folios must advance by one per printed page. Take the dominant offset (folio - unit index)
    # and impose it, so stray numbers in footers (years, note refs) cannot corrupt page citations.
    # Use the LOCAL dominant offset (+-20 pages) so reports whose numbering restarts between sections
    # (e.g. integrated report vs financial statements) or that insert unnumbered pages stay correct.
    from collections import Counter
    det = [(i, u.printed_page - i) for i, u in enumerate(raw) if u.printed_page is not None]
    offs = Counter(o for _, o in det)
    folio_offset, folio_votes = offs.most_common(1)[0] if offs else (None, 0)
    WIN = 20
    for i, u in enumerate(raw):
        local = Counter(o for j, o in det if abs(j - i) <= WIN)
        if local:
            o, votes = local.most_common(1)[0]
            if votes >= 3:
                u.printed_page = i + o if i + o > 0 else None
    units = []
    for u in raw:
        c = classify(u.text)
        units.append({**u.meta(), **c, "_unit": detect_unit(u.text), "_chars": len(u.text.strip())})

    # Notes pages without their own header inherit scope from the preceding notes/statement page.
    last_scope = None
    for u in units:
        if u["kind"] in TITLES and u["score"] > 0:
            last_scope = u["scope"]
        elif u["kind"] == "notes":
            if u["scope"] is None:
                u["scope"] = last_scope
            last_scope = u["scope"]

    stmt = lambda u: u["kind"] in TITLES and u["score"] > 0
    best = {}
    for u in units:                                   # standalone: best score is fine (not used downstream)
        if stmt(u) and u["scope"] != "consolidated":
            key = f"sa_{u['kind']}"
            if key not in best or u["score"] > best[key]["score"]:
                best[key] = u

    # Consolidated face statements = the tightest BS / P&L / CF cluster (rejects summary tables
    # elsewhere in the report that carry the same title, e.g. a 'Consolidated Balance Sheet' at p.47).
    pos = {u["uid"]: i for i, u in enumerate(units)}
    cand = {k: [u for u in units if stmt(u) and u["scope"] == "consolidated" and u["kind"] == k] for k in TITLES}
    best_cluster, best_val = None, None
    for pl in cand["pl"]:
        ip = pos[pl["uid"]]
        bs = min(cand["bs"], key=lambda u: abs(pos[u["uid"]] - ip), default=None)
        cf = min(cand["cf"], key=lambda u: abs(pos[u["uid"]] - ip), default=None)
        parts = [x for x in (bs, pl, cf) if x]
        val = sum(x["score"] for x in parts) - 2 * (max(pos[x["uid"]] for x in parts) - min(pos[x["uid"]] for x in parts))
        if best_val is None or val > best_val:
            best_cluster, best_val = {"bs": bs, "pl": pl, "cf": cf}, val
    if best_cluster:
        lo = min(pos[x["uid"]] for x in best_cluster.values() if x)
        hi = max(pos[x["uid"]] for x in best_cluster.values() if x)
        for k, u in best_cluster.items():
            if u:
                best[f"cons_{k}"] = u
        for u in cand["soce"]:
            if lo - 4 <= pos[u["uid"]] <= hi + 6:
                best.setdefault("cons_soce", u)
                hi = max(hi, pos[u["uid"]])
        # Consolidated notes region: after the face block until the next standalone statement,
        # auditor's report, or end-of-section marker. Headerless pages inside it are notes.
        end = len(units)
        for j in range(hi + 1, len(units)):
            u = units[j]
            if (stmt(u) and u["scope"] != "consolidated") or (u["kind"] == "notes" and u["scope"] == "standalone") \
                    or u["_end_marker"]:
                end = j
                break
        hdr_notes = [j for j in range(hi + 1, end) if units[j]["kind"] == "notes" and units[j]["scope"] == "consolidated"]
        if len(hdr_notes) >= 20:       # report labels its notes pages: trust the labelled extent
            end = min(end, hdr_notes[-1] + 1)
        for j in range(hi + 1, end):
            u = units[j]
            if not stmt(u):
                u["kind"], u["scope"] = "notes", "consolidated"
            elif u["scope"] == "consolidated":
                u["kind"] = "notes"          # a statement title inside notes is a sub-heading
        for j in range(lo, hi + 1):
            units[j]["scope"] = "consolidated"
        # anything marked consolidated-notes outside the region is not part of it
        for j, u in enumerate(units):
            if u["kind"] == "notes" and u["scope"] == "consolidated" and not (hi < j < end):
                u["scope"] = "other"

    cons_notes = [u["uid"] for u in units if u["kind"] == "notes" and u["scope"] == "consolidated"]
    units_detected = sorted({u["_unit"] for u in units if u["_unit"] and u["kind"] in TITLES and u["score"] > 0})

    problems = []
    for k in ("cons_bs", "cons_pl", "cons_cf"):
        if k not in best:
            problems.append(f"missing {k}")
    if not cons_notes:
        problems.append("no consolidated notes pages found")
    if len(units_detected) != 1:
        problems.append(f"ambiguous reporting unit: {units_detected}")
    # Units inside the consolidated block with no text layer would need OCR.
    no_text = [u["uid"] for u in units if u["_chars"] < 30 and u["scope"] == "consolidated"]

    result = {
        "co_id": co_id,
        "pdf": pdf_name,
        "n_units": len(units),
        "reporting_unit": units_detected[0] if len(units_detected) == 1 else None,
        "folio_offset": folio_offset, "folio_votes": f"{folio_votes}/{len(raw)}",
        "statements": {k: {"uid": v["uid"], "pdf_page": v["pdf_page"], "side": v["side"],
                           "printed_page": v["printed_page"], "score": v["score"]} for k, v in sorted(best.items())},
        "cons_notes_range": [cons_notes[0], cons_notes[-1]] if cons_notes else None,
        "n_cons_notes_units": len(cons_notes),
        "problems": problems + ([f"no text layer on face statements: {bad}"] if (bad := [x for x in no_text if x in
                    {v["uid"] for k, v in best.items() if k.startswith("cons")}]) else []),
        "warnings": [f"no text layer (OCR needed if a note lives here): {no_text}"] if no_text else [],
        "units": [{k: u[k] for k in ("uid", "pdf_page", "side", "printed_page", "kind", "scope", "score")}
                  for u in units],
    }
    write_json(co_work_dir(co_id) / "page_index.json", result)
    return result


def main(argv: list[str]) -> int:
    peers = {p["co_id"]: p for p in load_yaml("peers.yaml")["peers"]}
    rc = 0
    for co_id in argv:
        r = index_company(co_id, peers[co_id]["pdf"])
        s = r["statements"]
        line = ", ".join(f"{k}={v['uid']}(pp{v['printed_page']})" for k, v in s.items() if k.startswith("cons"))
        status = "OK" if not r["problems"] else "PROBLEMS: " + "; ".join(r["problems"])
        print(f"[{co_id}] unit={r['reporting_unit']} {line} notes={r['cons_notes_range']} "
              f"({r['n_cons_notes_units']} units) -> {status}")
        rc |= bool(r["problems"])
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
