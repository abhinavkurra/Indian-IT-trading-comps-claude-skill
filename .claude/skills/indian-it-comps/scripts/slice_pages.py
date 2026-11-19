"""Phase 1b: slice only the pages the extractor needs, deterministically.

1. Read the consolidated BS / P&L / CF face units found by index_pdf.py.
2. Parse face lines -> note reference for each target line item.
3. Find those note headings inside the consolidated notes (monotone numbering filter).
4. Add keyword-located notes (ESOP, business combinations, NCI, EPS, debt-like items).
5. Write work/<co>/pages/<uid>.txt (one printed page each) + work/<co>/manifest.json.

Usage:  python slice_pages.py <co_id> [<co_id> ...]
"""
from __future__ import annotations

import re
import sys

from common import INPUT_DIR, co_work_dir, fill_folios, iter_units, load_yaml, read_json, write_json

# target -> (statement, face-label regex). First matching face line wins per (statement, regex).
FACE_TARGETS = {
    "bs": {
        "investments": r"\binvestments?\b",
        "cash": r"^(?!.*other\s+than).*\bcash\s+and\s+cash\s+equivalents\b",
        "other_bank_balances": r"\b(other\s+bank\s+balances|bank\s+balances\s+other)",
        "equity_share_capital": r"\bequity\s+share\s+capital\b",
        "nci": r"\bnon[-\s]*controlling\s+interest",
        "borrowings": r"\bborrowings?\b",
        "lease_liabilities": r"\blease\s+liabilit",
        "other_financial_liabilities": r"\bother\s+financial\s+liabilit",
        "finance_lease_receivable": r"\bfinance\s+lease\s+receivable",
        "assets_held_for_sale": r"\bheld\s+for\s+sale\b",
    },
    "pl": {
        # revenue and D&A are read from the face statement; their notes are not sliced.
        "other_income": r"\bother\s+income\b",
        "finance_costs": r"\bfinance\s+costs?\b",
        "exceptional": r"\bexceptional\s+items?\b",
        "fx_line": r"^(?!.*translat).*\b(net\s+)?(loss|gain)?.*exchange\s+(fluctuation|difference|loss|gain)",
        "eps": r"\bearnings\s+per\s+(equity\s+)?share\b",
        "discontinued": r"\bdiscontinued\s+operations?\b",
    },
}

# Notes located by heading keywords (no face reference).
KEYWORD_NOTES = {
    "share_based_payments": (r"(share[-\s]*based|employees?\s+stock|stock\s+option|\besops?\b|\brsus?\b|"
                             r"equity[-\s]settled|incentive\s+plan|stock\s+payments?)"),
    "business_combinations": r"\b(business\s+combinations?|acquisitions?\s+of\s+subsidiar)",
    "nci_material_subs": r"\bnon[-\s]*controlling\s+interests?\b",
    "eps": r"\bearnings\s+per\s+(equity\s+)?share\b",
}
# Debt-like disclosures can be buried in fair-value / financial-instrument notes.
DEBT_LIKE_RX = re.compile(
    r"contingent\s+consideration|earn[-\s]?outs?|deferred\s+consideration|put\s+option|"
    r"redemption\s+liabilit|derecognition\s+of\s+nci|financial\s+liability.{0,40}non[-\s]controlling",
    re.I,
)
RESTRICTED_RX = re.compile(r"margin\s+money|earmarked|unpaid\s+dividend|unclaimed\s+dividend|escrow|lien", re.I)

NOTE_ID = r"(\d{1,2})\.?\s*([A-Z])?\s*(\((?:[ivxIVX]{1,5}|[a-hA-H])\))?"   # 14A, 18 (A), 13(iii), 5(iv), "10. (a)"
FACE_LINE = re.compile(r"^\s*[-–•\s]*(?:\(?[a-z]{1,4}\)\s*)*(?P<label>[A-Za-z][^\d]*?)\s{2,}(?P<note>" + NOTE_ID +
                       r")\s{2,}(?P<rest>.*)$")
HEADING = re.compile(r"^\s*(?:Note\s*(?:No\.?\s*)?)?(?P<id>" + NOTE_ID + r")\s*[\.:\)]?(?:\s+|(?<=\)))"
                     r"(?P<title>[A-Z][A-Za-z0-9 ,&\-–\(\)/':.]{2,95})\s*$")
TABLE_ROW = re.compile(r"\s{3,}|\d{1,3},\d{2,3}")   # column gaps or amounts => a table row, not a heading
UNIT_TAIL = re.compile(r"\s{2,}\([^()]*(million|crore|lakh|mn|inr|rs\.?)[^()]*\)\s*$", re.I)
# A located heading must be about the target; otherwise search the note chain by keyword.
TARGET_KEYWORDS = {
    "investments": r"invest", "cash": r"cash", "other_bank_balances": r"bank", "equity_share_capital": r"share\s*capital",
    "nci": r"non[-\s]*controlling|other\s+equity", "borrowings": r"borrow", "lease_liabilities": r"lease",
    "other_financial_liabilities": r"financial\s+liabilit", "finance_lease_receivable": r"lease",
    "assets_held_for_sale": r"held\s+for\s+sale|discontinued", "other_income": r"other\s+income",
    "finance_costs": r"finance\s+cost", "fx_line": r"exchange|foreign",
}
MAX_NOTE_UNITS = {"default": 3, "share_based_payments": 5, "business_combinations": 4, "eps": 2}
REFER_NOTE = re.compile(r"refer\s+note\s*" + NOTE_ID, re.I)


def norm_id(num: str, letter: str | None, paren: str | None) -> str:
    return f"{int(num)}{(letter or '').upper()}{(paren or '').lower().replace(' ', '')}"   # "18 (A)" -> "18(a)"


def parse_face(text: str, targets: dict) -> dict:
    out = {}
    for line in text.splitlines():
        m = FACE_LINE.match(line)
        r = REFER_NOTE.search(line)          # inline "(Refer note 45)" with no note column
        if m:
            label = re.sub(r"\s+", " ", m.group("label")).strip(" -")
            nid = norm_id(m.group(3), m.group(4), m.group(5))   # groups: label, note, num, letter, paren
        elif r:
            label = re.sub(r"\s+", " ", line[:r.start()]).strip(" -(")
            nid = norm_id(r.group(1), r.group(2), r.group(3))
        else:
            continue
        # drop list markers: "iii. Cash and cash equivalents", "(b) Borrowings", "a) ..."
        label = re.sub(r"^(?:\(?(?:[ivx]{1,5}|[a-h])[\.\)]\s*)+", "", label, flags=re.I).strip()
        for key, rx in targets.items():
            if re.search(rx, label, re.I):
                out.setdefault(key, [])
                if nid not in [x["note"] for x in out[key]]:
                    out[key].append({"note": nid, "face_label": label})
    return out


def find_headings(units: list) -> list[dict]:
    """All candidate note headings in reading order, filtered to the longest monotone chain."""
    cands = []
    for ui, u in enumerate(units):
        lines = u.text.splitlines()
        for li, line in enumerate(lines):
            s = UNIT_TAIL.sub("", line.strip())  # "21. BORROWINGS      (INR million)" -> "21. BORROWINGS"
            m = HEADING.match(s)
            if not m or len(s) > 110:
                continue
            title = m.group("title").strip()
            if len(re.findall(r"[A-Za-z]", title)) < 4 or TABLE_ROW.search(title):
                continue
            cands.append({"n": int(m.group(2)), "id": norm_id(m.group(2), m.group(3), m.group(4)),
                          "title": title, "ui": ui, "li": li, "uid": u.uid,
                          "prefixed": bool(re.match(r"note\b", s, re.I))})
    # Reports that title notes "Note 12 ..." use that form consistently: drop bare-number lines
    # (subsidiary lists, numbered narrative) when the prefixed form dominates.
    if sum(c["prefixed"] for c in cands) >= 10:
        cands = [c for c in cands if c["prefixed"]]
    # Longest non-decreasing chain in n with steps <= 5 (rejects numbered sub-lists inside notes).
    best = [1] * len(cands)
    prev = [-1] * len(cands)
    for i in range(len(cands)):
        for j in range(i):
            if cands[j]["n"] <= cands[i]["n"] <= cands[j]["n"] + 5 and best[j] + 1 > best[i]:
                best[i], prev[i] = best[j] + 1, j
    if not cands:
        return []
    i = max(range(len(cands)), key=lambda k: best[k])
    chain = []
    while i != -1:
        chain.append(cands[i])
        i = prev[i]
    return chain[::-1]


def note_span(chain: list[dict], idx: int, whole_parent: bool, cap: int) -> tuple[list[int], bool]:
    """Unit indices covering the note at chain[idx]. A sub-note ends at the next heading of any id;
    whole_parent (face cited a parent like '14' that has sub-notes) runs to the next top-level number."""
    n = chain[idx]["n"]
    end = idx
    while end + 1 < len(chain) and chain[end + 1]["id"] == chain[idx]["id"]:   # "(continued)" headings
        end += 1
    if whole_parent:
        while end + 1 < len(chain) and chain[end + 1]["n"] == n:
            end += 1
    s_ui = chain[idx]["ui"]
    if end + 1 < len(chain):
        nxt = chain[end + 1]
        e_ui = nxt["ui"] if nxt["li"] > 6 else nxt["ui"] - 1   # include page if note content precedes next heading
    else:
        e_ui = s_ui + 1
    span = list(range(s_ui, max(e_ui, s_ui) + 1))
    return span[:cap], len(span) > cap


def locate(chain, note_id) -> tuple[int | None, bool]:
    """(chain index, whole_parent). Exact sub-note match is tight; parent-number fallback takes the whole note."""
    exact = [i for i, h in enumerate(chain) if h["id"] == note_id]
    if exact:
        h = chain[exact[0]]
        has_subs = sum(1 for x in chain if x["n"] == h["n"]) > 1
        return exact[0], has_subs and h["id"] == str(h["n"])
    top = int(re.match(r"\d+", note_id).group())
    same = [i for i, h in enumerate(chain) if h["n"] == top]
    return (same[0], True) if same else (None, False)


def slice_company(co_id: str, pdf_name: str) -> dict:
    wd = co_work_dir(co_id)
    idx = read_json(wd / "page_index.json")
    st = idx["statements"]
    cons_uids = {u["uid"] for u in idx["units"] if u["scope"] == "consolidated"}
    pdf_pages = sorted({u["pdf_page"] for u in idx["units"] if u["uid"] in cons_uids})
    all_units = list(iter_units(INPUT_DIR / pdf_name, pdf_pages))
    for u in all_units:      # non-breaking / figure / en hyphens and curly quotes break label matching
        u.text = re.sub("[\\u2010-\\u2015\\u2212]", "-", u.text).translate(
            {0x2018: "'", 0x2019: "'", 0x201C: '"', 0x201D: '"'})
    folio = {u["uid"]: u["printed_page"] for u in idx["units"]}   # authoritative folios from the full-document index
    for u in all_units:
        u.printed_page = folio.get(u.uid, u.printed_page)
    by_uid = {u.uid: u for u in all_units}
    notes_units = [u for u in all_units if u.uid in cons_uids and
                   next(x for x in idx["units"] if x["uid"] == u.uid)["kind"] == "notes"]

    selected: dict[str, set] = {}      # uid -> reasons
    def add(uid, reason):
        selected.setdefault(uid, set()).add(reason)

    targets, problems = {}, []
    face_units = {}
    for stmt in ("bs", "pl", "cf"):
        key = f"cons_{stmt}"
        if key not in st:
            problems.append(f"no {key}")
            continue
        uid = st[key]["uid"]
        face_units[stmt] = [uid]
        add(uid, f"face:{stmt}")
        order = [u.uid for u in all_units]
        if stmt == "bs" and not re.search(r"total\s+assets", by_uid[uid].text, re.I):
            prv = order[order.index(uid) - 1] if order.index(uid) > 0 else None
            if prv and re.search(r"total\s+assets", by_uid[prv].text, re.I):
                face_units[stmt].insert(0, prv)
                add(prv, "face:bs:prev")
        # Continuation page if the statement's closing line is not on the first unit.
        closers = {"bs": r"total\s+equity\s+and\s+liabilities", "pl": r"earnings\s+per", "cf": r"cash\s+and\s+cash\s+equivalents\s+at\s+the\s+end"}
        if not re.search(closers[stmt], by_uid[uid].text, re.I):
            order = [u.uid for u in all_units]
            nxt = order[order.index(uid) + 1]
            face_units[stmt].append(nxt)
            add(nxt, f"face:{stmt}:cont")

    chain = find_headings(notes_units)
    nu_index = {u.uid: i for i, u in enumerate(notes_units)}

    for stmt, tmap in FACE_TARGETS.items():
        text = "\n".join(by_uid[u].text for u in face_units.get(stmt, []))
        refs = parse_face(text, tmap)
        for key in tmap:
            for ref in refs.get(key, []):
                ci, whole = locate(chain, ref["note"])
                entry = {"statement": stmt, "note": ref["note"], "face_label": ref["face_label"]}
                kw = TARGET_KEYWORDS.get(key)
                if kw and (ci is None or not re.search(kw, chain[ci]["title"], re.I)):
                    # number match missing or about something else -> keyword search, prefer same top-level n,
                    # then best word overlap with the face label (current vs non-current etc.)
                    top = int(re.match(r"\d+", ref["note"]).group())
                    words = set(re.findall(r"[a-z]{4,}", ref["face_label"].lower()))
                    hits = [i for i, h in enumerate(chain) if re.search(kw, h["title"], re.I) and h["n"] == top]
                    if hits:
                        ci = max(hits, key=lambda i: len(words & set(re.findall(r"[a-z]{4,}", chain[i]["title"].lower()))))
                        whole = False
                        entry["matched_by"] = "keyword"
                    else:
                        ci = None
                        # last resort: a short heading-style line naming the item, e.g. "9. (a)  Equity share
                        # capital" or "(d) Cash and cash equivalents", anywhere in the notes region
                        lw = re.findall(r"[a-z]+", ref["face_label"].lower())[:4]
                        hrx = re.compile(r"^(?:note\s*)?(?:\d{1,2}\s*\.?\s*)?(?:\(?[a-z]{1,4}\)\s*[:.]?\s*)?"
                                         + r"\W*".join(lw) + r"\b", re.I)
                        def heading_like(l):
                            s = UNIT_TAIL.sub("", l.strip())
                            mm = hrx.match(s)
                            return bool(mm) and len(s) < 80 and not TABLE_ROW.search(s[mm.end():])
                        for ui, u in enumerate(notes_units):
                            ln = next((l for l in u.text.splitlines() if heading_like(l)), None)
                            if ln:
                                span = [ui, ui + 1] if ui + 1 < len(notes_units) else [ui]
                                entry.update(heading=ln.strip()[:80], units=[notes_units[i].uid for i in span],
                                             status="ok", matched_by="label_search")
                                for i in span:
                                    add(notes_units[i].uid, f"note:{key}")
                                break
                        if entry.get("matched_by") == "label_search":
                            targets.setdefault(key, []).append(entry)
                            continue
                if ci is None:
                    entry["status"] = "heading_not_found"
                    problems.append(f"{key}: note {ref['note']} heading not found")
                else:
                    span, trunc = note_span(chain, ci, whole, MAX_NOTE_UNITS.get(key, MAX_NOTE_UNITS["default"]))
                    entry.update(heading=chain[ci]["title"], units=[notes_units[i].uid for i in span],
                                 status="truncated" if trunc else "ok")
                    for i in span:
                        add(notes_units[i].uid, f"note:{key}")
                targets.setdefault(key, []).append(entry)

    for key, rx in KEYWORD_NOTES.items():
        hits = [i for i, h in enumerate(chain) if re.search(rx, h["title"], re.I)]
        if not hits:
            if key not in targets:
                targets[key] = [{"status": "not_found"}]
            continue
        # Prefer the last match for NCI (the material-subsidiary disclosure follows the equity note).
        ci = hits[-1] if key == "nci_material_subs" else hits[0]
        span, trunc = note_span(chain, ci, True, MAX_NOTE_UNITS.get(key, MAX_NOTE_UNITS["default"]))
        targets.setdefault(key, []).append({"note": chain[ci]["id"], "heading": chain[ci]["title"],
                                            "units": [notes_units[i].uid for i in span],
                                            "status": "truncated" if trunc else "ok"})
        for i in span:
            add(notes_units[i].uid, f"note:{key}")

    # Skip the accounting-policy section (policy wording is not a balance).
    first_real = next((h["ui"] for h in chain if h["n"] >= 2), 0)
    empty = {u.uid for u in notes_units if len(u.text.strip()) < 30}
    for key, entries in targets.items():
        for e in entries:
            if empty & set(e.get("units") or []):
                problems.append(f"{key}: note pages without text layer {sorted(empty & set(e['units']))} - OCR needed")
    debt_like_units = [u.uid for i, u in enumerate(notes_units)
                       if i >= first_real and DEBT_LIKE_RX.search(u.text)]
    for uid in debt_like_units:
        add(uid, "scan:debt_like")
    targets["debt_like_scan"] = [{"units": debt_like_units, "status": "ok" if debt_like_units else "none_found"}]
    restricted = [uid for uid in selected if RESTRICTED_RX.search(by_uid[uid].text)]
    targets["restricted_cash_mentions"] = [{"units": restricted}]

    pages_dir = wd / "pages"
    pages_dir.mkdir(exist_ok=True)
    for f in pages_dir.glob("*.txt"):
        f.unlink()
    order = [u.uid for u in all_units]
    files, total_chars = [], 0
    for uid in sorted(selected, key=order.index):
        u = by_uid[uid]
        header = (f"### SOURCE pdf={pdf_name} pdf_page={u.pdf_page} half={u.side} "
                  f"printed_page={u.printed_page} uid={uid}\n### WHY {', '.join(sorted(selected[uid]))}\n")
        body = "\n".join(l.rstrip() for l in u.text.splitlines())
        body = re.sub(r"\n{3,}", "\n\n", body)
        (pages_dir / f"{uid}.txt").write_text(header + body, encoding="utf-8")
        total_chars += len(body)
        files.append({"uid": uid, "pdf_page": u.pdf_page, "side": u.side, "printed_page": u.printed_page,
                      "why": sorted(selected[uid])})

    manifest = {
        "co_id": co_id, "pdf": pdf_name, "reporting_unit": idx["reporting_unit"],
        "n_pages": len(files), "approx_tokens": total_chars // 4,
        "pages": files, "targets": targets, "problems": problems,
        "note_chain": [f"{h['id']} {h['title']} @{h['uid']}" for h in chain],
    }
    write_json(wd / "manifest.json", manifest)

    # Keep an existing extraction's page citations in sync with the authoritative folios (metadata only:
    # values and uids are untouched; uids are stable across re-slices).
    exf = wd / "extracted.json"
    if exf.exists():
        import os
        from validate import walk_cites
        ex = read_json(exf)
        changed = 0
        for _, cite in walk_cites(ex):
            if cite["uid"] in folio and cite.get("printed_page") != folio[cite["uid"]]:
                cite["printed_page"], changed = folio[cite["uid"]], changed + 1
        for p in ex.get("esop_plans", []):
            if p.get("uid") in folio and p.get("printed_page") != folio[p["uid"]]:
                p["printed_page"], changed = folio[p["uid"]], changed + 1
        if changed:
            write_json(exf, ex)
            manifest["refolioed_citations"] = changed
            write_json(wd / "manifest.json", manifest)
        # An extraction that still validates against the new slice stays current (no re-extraction needed).
        import validate
        chk = validate.validate(co_id)
        if all(r["ok"] for r in chk.results if r["level"] == "FAIL"):
            os.utime(exf, None)
    return manifest


def main(argv):
    peers = {p["co_id"]: p for p in load_yaml("peers.yaml")["peers"]}
    rc = 0
    for co_id in argv:
        m = slice_company(co_id, peers[co_id]["pdf"])
        status = "OK" if not m["problems"] else "PROBLEMS: " + "; ".join(m["problems"])
        print(f"[{co_id}] {m['n_pages']} pages, ~{m['approx_tokens']} tokens, unit={m['reporting_unit']} -> {status}")
        rc |= bool(m["problems"])
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
