"""Shared helpers: paths, PDF page-unit iteration, number parsing, unit conversion."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]
ROOT = SKILL_DIR.parents[2]            # project root (<root>/.claude/skills/<skill>)
INPUT_DIR = ROOT / "input_pdfs"
WORK_DIR = ROOT / "work"
OUTPUT_DIR = ROOT / "output"
CONFIG_DIR = SKILL_DIR / "config"


def load_yaml(name: str) -> dict:
    import yaml
    with open(CONFIG_DIR / name, encoding="utf-8") as f:
        return yaml.safe_load(f)


def co_work_dir(co_id: str) -> Path:
    d = WORK_DIR / co_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def write_json(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)


def read_json(path: Path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------- page units
# Many Indian ARs are exported as two-page spreads. A "unit" is one printed page:
# (pdf_page, side) where side is "F" (full page), "L" or "R" (spread halves).

@dataclass
class Unit:
    pdf_page: int          # 1-based physical PDF page
    side: str              # F | L | R
    printed_page: int | None
    text: str

    @property
    def uid(self) -> str:
        return f"p{self.pdf_page:03d}{self.side}"

    def meta(self) -> dict:
        d = asdict(self)
        d.pop("text")
        d["uid"] = self.uid
        return d


_PRINTED_RE = re.compile(r"^\s*(\d{1,4})\s*$")
# Folio sharing a footer line with running text, e.g. "300     AI-First Engineering | ..."
_FOLIO_EDGE_RE = re.compile(r"^\s*(\d{1,4})\s{2,}([^\d]+)$|^([^\d]*?(?:annual report[^\d]*\d{4}-\d{2,4})?[^\d]*?)\s{2,}(\d{1,4})\s*$", re.I)


def _printed_page(text: str) -> int | None:
    """Printed folio from the footer (last 2 lines) or header (first 2 lines). None if absent."""
    lines = [l for l in text.splitlines() if l.strip()]
    for l in list(reversed(lines[-2:])) + lines[:2]:
        m = _PRINTED_RE.match(l)
        if m:
            return int(m.group(1))
        m = _FOLIO_EDGE_RE.match(l)
        if m:
            return int(m.group(1) or m.group(4))
    return None


def fill_folios(units: list) -> None:
    """Infer missing printed page numbers from neighbours (folios increase by 1 per unit)."""
    for i, u in enumerate(units):
        if u.printed_page is None:
            if i > 0 and units[i - 1].printed_page is not None:
                u.printed_page = units[i - 1].printed_page + 1
    for i in range(len(units) - 2, -1, -1):
        if units[i].printed_page is None and units[i + 1].printed_page is not None:
            units[i].printed_page = units[i + 1].printed_page - 1


def clean(text: str) -> str:
    # Many AR fonts map the rupee glyph to a backtick; normalise to "INR".
    return text.replace("`", "INR ").replace("\u20b9", "INR ").replace("\u00a0", " ")


def iter_units(pdf_path: Path, pages: list[int] | None = None):
    """Yield Units for the PDF. `pages` (1-based) restricts which PDF pages are read."""
    import pymupdf
    doc = pymupdf.open(pdf_path)
    rng = pages if pages is not None else range(1, doc.page_count + 1)
    for pno in rng:
        page = doc[pno - 1]
        r = page.rect
        if r.width > r.height * 1.2:      # landscape spread -> two printed pages
            w = r.width / 2
            clips = [("L", pymupdf.Rect(0, 0, w, r.height)), ("R", pymupdf.Rect(w, 0, r.width, r.height))]
        else:
            clips = [("F", None)]
        for side, clip in clips:
            t = clean(page.get_text(clip=clip, sort=True) if clip else page.get_text(sort=True))
            yield Unit(pno, side, _printed_page(t), t)


# ---------------------------------------------------------------- numbers
_NUM_RE = re.compile(r"^\(?-?[\d,]+(\.\d+)?\)?$")


def parse_number(s: str) -> float | None:
    """Parse AR-style numbers: '1,48,814', '(1,795.82)', '-' (nil), '—'. Returns None if not numeric."""
    s = s.strip().replace(" ", "")
    if s in {"-", "\u2013", "\u2014", "--", "Nil", "nil"}:
        return 0.0
    if not _NUM_RE.match(s):
        return None
    neg = s.startswith("(") and s.endswith(")") or s.startswith("-")
    v = float(s.strip("()").lstrip("-").replace(",", ""))
    return -v if neg else v


UNIT_TO_CRORE = {"crore": 1.0, "million": 0.1, "lakh": 0.01, "thousand": 0.0001, "rupees": 1e-7}

_UNIT_RE = re.compile(
    r"(?:inr|rs\.?|rupees?)\s*(?:in\s*)?(crores?|cr\b|millions?|mn\b|lakhs?|lacs?|thousands?)"
    r"|(?:amounts?|figures?)\s+(?:are\s+)?in\s+(?:inr\s*|rs\.?\s*)?(crores?|millions?|mn\b|lakhs?)",
    re.I,
)


def detect_unit(text: str) -> str | None:
    m = _UNIT_RE.search(text)
    if not m:
        return None
    u = (m.group(1) or m.group(2)).lower()
    if u.startswith("cr"):
        return "crore"
    if u.startswith("mi") or u == "mn":
        return "million"
    if u.startswith("la"):
        return "lakh"
    if u.startswith("th"):
        return "thousand"
    return None
