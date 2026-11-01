"""
ts_parser.py - Reads a JJM Technical Sanction (TS) cost-estimate abstract (digital or scanned PDF).

Every sanctioned item is verified with Quantity x Rate = Amount, and the item amounts are summed
against the printed Sub Total. Nothing is assumed about a particular scheme: values that the
document's arithmetic cannot confirm are flagged for review instead of being "healed".
"""

import io
import itertools
import logging
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

from abstract_model import Check
from utils import format_inr, money_eq, number_candidates

logger = logging.getLogger(__name__)

PdfSource = Union[str, Path, bytes, io.BytesIO]


@dataclass
class TSItem:
    sl: str
    desc: str
    qty: float
    rate: float
    unit: str
    amount: float
    status: str = "ok"                   # ok | corrected | review
    heading: bool = False
    notes: List[str] = field(default_factory=list)
    uncertain: List[str] = field(default_factory=list)   # words of the text the OCR engines did not settle

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class TSMetaData:
    scheme: str = ""
    estimated_cost_lakhs: float = 0.0
    sub_total: Optional[float] = None
    beneficiary_share: Optional[float] = None
    contingencies: Optional[float] = None
    grand_total: Optional[float] = None
    division: str = ""
    sub_division: str = ""

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class TSParseResult:
    metadata: TSMetaData
    items: List[TSItem]
    checks: List[Check]
    is_scanned: bool
    pages: List[int] = field(default_factory=list)     # 0-based pages the items were read from

    @property
    def priced_items(self) -> List[TSItem]:
        return [it for it in self.items if not it.heading]

    @property
    def total_sanctioned_amount(self) -> float:
        return sum(it.amount for it in self.priced_items)

    @property
    def is_valid(self) -> bool:
        return all(c.passed for c in self.checks if c.level == "critical")


# ---------------------------------------------------------------------------
# Row sources: each yields dicts {role: (text, confidence, alternatives)}
# ---------------------------------------------------------------------------

ROLE_PATTERNS = [
    ("sl", r"\bs[l1i|]\b|^no\.?$|\bno\b|\bs[l1i|r]?\.?\s*no\b|serial"),     # Sl No / SlNo / S1 No / Sr. No
    ("desc", r"name\s*of\s*work|description|particular"),
    ("qty", r"qu?a?n[a-z]*t[iy]|\bqty\b|\bqnty\b"),
    ("rate", r"^rate|\brate\b"),
    ("unit", r"\bu[nm][il1|!]t\b|\bunits?\b"),           # OCR: 'Unlt', 'Un1t'
    ("amount", r"amount|rupees|estimate"),
]


def _assign_roles(headers: List[str]) -> Dict[str, int]:
    roles: Dict[str, int] = {}
    for c, h in enumerate(headers):
        h = h.lower()
        for role, pat in ROLE_PATTERNS:
            if role not in roles and re.search(pat, h):
                roles[role] = c
                break
    return roles


Cell = Tuple[str, float, List[str]]

# "Sub Total" as printed and as OCR mangles it ("su Total", "Sub Tota1", "SubTotal")
SUBTOTAL_RE = re.compile(r"\bs[uv]b?\s*-?\s*t[o0]ta[l1i]", re.I)

# units as printed in TS abstracts
_UNIT_TOKEN = r"(?:R\.?\s?m\.?|Rmt|RM|Job|No\.?s\.?|Nos\.?|Each|Sq\.?\s?m|Cu\.?\s?m|Cum|Qntl|Qtl|Quintal)"
_UNIT_OK = re.compile(_UNIT_TOKEN + r"|M|Kva|KVA|Kg|Set|L\.?S\.?|Ltr|No", re.I)
_UNIT_IN_TEXT = re.compile(r"(?<![^\s)])(" + _UNIT_TOKEN + r")(?!\S)", re.I)


def find_abstract_pages(pdf_source: PdfSource) -> List[int]:
    """
    0-based pages of the "Abstract of Cost Estimate (for Revised Technical Sanction)" table in a TS
    document of any length: the page whose table header reads "... Amount as per TS estimate", plus
    the following pages up to and including the one with the "Sub Total".
    (The underlined title itself OCRs poorly on photos; the column header is reliable.)
    """
    from table_extractor import quick_page_text
    texts = [t.lower() for t in quick_page_text(pdf_source)]

    def header_score(t: str) -> int:
        return sum(bool(re.search(p, t)) for p in (r"qnty|quan|qty", r"\brate", r"\bunit", r"amount"))

    def comparative(t: str) -> bool:
        # recognised by its columns, not its title: some abstracts are headed "Comparative statement"
        # but carry Unit / Rate / Quantity, while the real comparison has (Saving) / Shortfall instead
        return bool(re.search(r"saving|short\s*-?\s*fall", t)) and header_score(t) < 3

    def subtotal_near(i: int) -> bool:
        return any(SUBTOTAL_RE.search(x) for x in texts[i:i + 4])

    # strongest first: the abstract title; then the old header tests (title unreadable on some photos)
    tiers = [
        [i for i, t in enumerate(texts) if re.search(r"abstra", t) and not comparative(t)
         and (re.search(r"ts\s*estim|technical\s*sanction|revised|estimate", t) or header_score(t) >= 2)
         and subtotal_near(i)],
        [i for i, t in enumerate(texts) if re.search(r"ts\s*estim|technical\s*sanction", t)
         and header_score(t) >= 3 and not comparative(t)],
        [i for i, t in enumerate(texts) if header_score(t) == 4 and not comparative(t) and subtotal_near(i)],
    ]
    starts = next((s for s in tiers if s), [])
    if not starts:
        return []
    revised = [i for i in starts if "revised" in texts[i]]
    first = (revised or starts)[0]
    pages = [first]
    for j in range(first, min(len(texts), first + 6)):
        if j != first:
            if comparative(texts[j]):
                break
            pages.append(j)
        if SUBTOTAL_RE.search(texts[j]):
            break
    return pages


def _rows_from_scanned(pdf_source: PdfSource, pages: Optional[List[int]] = None) -> Tuple[List[Dict[str, Cell]], List[str]]:
    from table_extractor import iter_pdf_tables, _join_words
    tables = iter_pdf_tables(pdf_source, pages)       # read page by page: reading stops at the Sub Total
    rows: List[Dict[str, Cell]] = []
    page_texts: List[str] = []
    ref: Optional[Dict[str, int]] = None          # roles from the page that carried the column header
    for t in tables:
        page_texts.append("\n".join(_lines_of(t.page_words)))
        if t.n_cols < 4:
            continue
        hdr, roles = None, None
        for r in range(min(t.n_rows, 8)):
            cand = _assign_roles([t.text(r, c) for c in range(t.n_cols)])
            if {"desc", "qty", "rate"} <= set(cand):
                hdr, roles = r, cand
                # a header split over two rows ("Sl No" printed below the other titles)
                for r2 in (r - 1, r + 1):
                    if 0 <= r2 < t.n_rows:
                        for k, c in _assign_roles([t.text(r2, c) for c in range(t.n_cols)]).items():
                            if k not in roles and c not in roles.values():
                                roles[k] = c
                                hdr = max(hdr, r2)
                break
        if roles is not None:
            if "sl" not in roles and roles["desc"] > 0 and roles["desc"] - 1 not in roles.values():
                roles["sl"] = roles["desc"] - 1          # the unlabelled column left of 'Name of Work'
            ref = roles
        elif ref is None:
            continue
        else:
            # continuation page without a header: its columns may be shifted or split differently
            # (an extra margin column, separate '₹' columns), so they are recognised by their contents
            body = range(t.n_rows)
            options = [_roles_by_content(t, ref, body)]
            if ref.get("_ncols") == t.n_cols:
                options.append(dict(ref))
            options = [o for o in options if o]
            if not options:
                continue
            roles = max(options, key=lambda o: _roles_score(t, o, body))
        roles["_ncols"] = t.n_cols
        # columns without a header of their own inside the description block (e.g. the "110 | mm dia. HDPE"
        # split cells) belong to the description
        num_cols = [roles[k] for k in ("qty", "rate", "unit", "amount") if k in roles]
        desc_extra = [c for c in range(roles["desc"] + 1, min(num_cols) if num_cols else roles["desc"] + 1)
                      if c not in roles.values()]
        page_rows: List[Dict[str, Cell]] = []
        for r in range((hdr + 1) if hdr is not None else 0, t.n_rows):
            row = _read_row(t, r, roles, desc_extra)
            if any(v[0] for k, v in row.items() if not k.startswith("_")):
                page_rows.append(row)
        rows.extend(_merge_split_rows(page_rows))
        if any(SUBTOTAL_RE.search(_txt(r, "desc")) for r in page_rows):
            break                     # the abstract ends here (later pages: comparative statement, estimates)
    return rows, page_texts


_TEXT_ROLES = ("unit", "sl", "desc")          # unit first: a unit read into the description is dropped from it


def _read_row(t, r: int, roles: Dict[str, int], desc_extra: List[int]) -> Dict[str, Cell]:
    """
    One grid row as {role: (text, confidence, alternative readings)}.
    Text cells (Name of Work, Unit, Sl) are the word-by-word consensus of the two OCR engines; number
    cells keep the main engine's reading with the second engine's as one more alternative (the
    Quantity x Rate = Amount check chooses among them).
    Words neither engine settles are listed under "_uncertain" (shown for review).
    """
    from table_extractor import _join_words
    from text_consensus import fold, key, merge
    row: Dict[str, Cell] = {}
    uncertain: List[str] = []
    drop: set = set()
    for role in sorted((k for k in roles if not k.startswith("_")),
                       key=lambda k: _TEXT_ROLES.index(k) if k in _TEXT_ROLES else -1):
        c = roles[role]
        if c >= t.n_cols:
            continue
        ws, alt = list(t.words(r, c)), list(t.alt_words(r, c))
        if role == "desc":
            for c2 in desc_extra:
                ws += t.words(r, c2)
                alt += t.alt_words(r, c2)
        # lines are clustered on the words' own centres (fixed bands split a tilted line in two)
        main, second = _join_words(ws), _join_words(alt)
        score = min((w.score for w in ws), default=1.0)
        if role == "unit":
            # a unit is one of a few words: the reading that is a unit wins; when the main engine found
            # no unit in its cell, the unit was read into the description and is taken out of it there
            r_u, w_u = (_join_wrapped_unit(fold(x).strip("()[] ")) for x in (main, second))
            text = r_u if _UNIT_OK.fullmatch(r_u or "-") else (w_u if _UNIT_OK.fullmatch(w_u or "-") else r_u)
            drop = {key(text)} if text and not _UNIT_OK.fullmatch(r_u or "-") else set()
            row[role] = (text, score, [w_u] if w_u and w_u != text else [])
            continue
        if role in _TEXT_ROLES:
            text, unc = merge(main, second, drop=drop if role == "desc" else set())
            uncertain += unc
            row[role] = (text, score, [fold(second)] if second and fold(second) != text else [])
        else:
            row[role] = (main, score, [a for w in ws for a in w.alts] + ([second] if second and second != main else []))
    if uncertain:
        row["_uncertain"] = (" ".join(uncertain), 1.0, [])
    return row


def _join_wrapped_unit(unit: str) -> str:
    """A unit wrapped inside its narrow cell ('Quin' / 'tal') is one word."""
    parts = unit.split()
    if len(parts) == 2 and parts[0].isalpha() and parts[1].isalpha() and parts[1][0].islower():
        return parts[0] + parts[1]
    return unit


def _txt(row: Dict[str, Cell], role: str) -> str:
    return (row.get(role) or ("", 1.0, []))[0].strip()


def _first_number(text: str) -> Optional[float]:
    c = number_candidates(text, 2) if re.search(r"\d", text or "") else []
    return c[0][0] if c else None


def _roles_score(t, roles: Dict[str, int], body) -> int:
    """Rows whose Quantity x Rate = Amount under this column assignment (+ unit-looking unit cells)."""
    score = 0
    for r in body:
        cells = {k: t.text(r, c) for k, c in roles.items() if not k.startswith("_") and c < t.n_cols}
        q, rt, a = (_first_number(cells.get(k, "")) for k in ("qty", "rate", "amount"))
        if q and rt and a and money_eq(q * rt, a, 0.011):
            score += 3
        if "unit" in cells and _UNIT_OK.fullmatch(cells["unit"].strip() or "-"):
            score += 1
    return score


def _roles_by_content(t, ref: Dict[str, int], body) -> Optional[Dict[str, int]]:
    """
    Column roles of a header-less page, keeping the left-to-right order of the header page's roles:
    the description is the column with the most text; the roles left / right of it are matched to the
    columns left / right of it that carry content (columns holding only a '₹' sign are skipped).
    """
    seq = [k for k, _ in sorted(((k, c) for k, c in ref.items() if not k.startswith("_")), key=lambda kc: kc[1])]
    if "desc" not in seq:
        return None
    n_rows = max(1, len(list(body)))
    used, letters = [], {}
    for c in range(t.n_cols):
        texts = [t.text(r, c).strip() for r in body]
        filled = [x for x in texts if x]
        if len(filled) < max(1, 0.15 * n_rows):
            continue
        if all(len(x) <= 2 and len(re.sub(r"\D", "", x)) <= 1 for x in filled):
            continue                                   # a '₹' column (read as '₹', '2', 'Z', ...)
        used.append(c)
        letters[c] = sum(len(re.findall(r"[A-Za-z]", x)) for x in filled) / len(filled)
    if not used:
        return None
    desc_c = max(used, key=lambda c: letters[c])
    i = seq.index("desc")
    left_roles, right_roles = seq[:i], seq[i + 1:]
    left_cols = [c for c in used if c < desc_c]
    right_cols = [c for c in used if c > desc_c]
    roles = {"desc": desc_c}
    for k, c in zip(reversed(left_roles), reversed(left_cols)):
        roles[k] = c
    # extra columns next to the description (split "110 | mm dia. Pipe" cells) stay with the description
    for k, c in zip(reversed(right_roles), reversed(right_cols)):
        roles[k] = c
    return roles


def _merge_split_rows(rows: List[Dict[str, Cell]]) -> List[Dict[str, Cell]]:
    """
    On photographed pages one item often spans two grid rows: its rate / quantity sit on one row and
    the description and amount on the next (or the amount drops below a wrapped description). Rows
    are joined when they carry complementary values and no Sl number of their own.
    """
    def nums(row):
        return {k for k in ("qty", "rate", "amount") if re.search(r"\d", _txt(row, k))}

    def summary(row):
        return any(re.search(p, _txt(row, "desc"), re.I) for _, p in _SUMMARY)

    def incomplete(row):
        return len(nums(row)) < 3

    def join(a, b):           # a above b
        out = dict(a)
        for k, v in b.items():
            if not v[0].strip():
                continue
            if _txt(out, k) and k in ("desc", "_uncertain"):
                out[k] = (_txt(out, k) + " " + v[0].strip(), min(out[k][1], v[1]), out[k][2] + v[2])
            elif not _txt(out, k):
                out[k] = v
        return out

    # 1. value-only fragments (no Sl, no description) join the neighbouring item that lacks those values
    rows = list(rows)
    i = 0
    while i < len(rows):
        r = rows[i]
        n = nums(r)
        if n and not _txt(r, "sl") and not _txt(r, "desc"):
            cands = []
            for j, below in ((i + 1, True), (i - 1, False)):
                if 0 <= j < len(rows):
                    o = rows[j]
                    if _txt(o, "desc") and not summary(o) and not (n & nums(o)) and incomplete(o):
                        merged = join(r, o) if below else join(o, r)
                        cands.append((len(nums(merged)) == 3, bool(_txt(o, "sl")), below, j, merged))
            if cands:
                _, _, _, j, merged = max(cands)
                rows[j] = merged
                del rows[i]
                continue
        i += 1

    # 2. continuation rows: no Sl, and either values the item above lacks, or a wrapped line of
    #    description (starting in lower case) under an item
    out: List[Dict[str, Cell]] = []
    for r in rows:
        if out and not _txt(r, "sl") and _txt(r, "desc") and not summary(r) and not summary(out[-1]):
            p = out[-1]
            n, pn = nums(r), nums(p)
            values_fill = n and pn and not (n & pn) and incomplete(p)
            wrapped = not n and re.match(r"[a-z(]", _txt(r, "desc")) and pn
            if values_fill or wrapped:
                out[-1] = join(p, r)
                continue
        out.append(r)
    return out


def _rows_from_digital(pdf_source: PdfSource, pages: Optional[List[int]] = None) -> Tuple[List[Dict[str, Cell]], List[str]]:
    import pdfplumber
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    src = io.BytesIO(pdf_source) if isinstance(pdf_source, (bytes, bytearray)) else pdf_source
    rows: List[Dict[str, Cell]] = []
    page_texts: List[str] = []
    roles: Optional[Dict[str, int]] = None
    with pdfplumber.open(src) as pdf:
        for pi, page in enumerate(pdf.pages):
            if pages is not None and pi not in pages:
                continue
            page_texts.append(page.extract_text() or "")
            for tbl in page.extract_tables() or []:
                start = 0
                for i, raw in enumerate(tbl[:6]):
                    cand = _assign_roles([(c or "").replace("\n", " ") for c in raw])
                    if {"desc", "qty", "rate"} <= set(cand):
                        roles, start = cand, i + 1
                        break
                if roles is None:
                    continue
                for raw in tbl[start:]:
                    cells = [(c or "").replace("\n", " ").strip() for c in raw]
                    row = {role: (cells[c] if c < len(cells) else "", 1.0, []) for role, c in roles.items()}
                    if any(v[0] for v in row.values()):
                        rows.append(row)
    return rows, page_texts


def _lines_of(words) -> List[str]:
    ws = sorted(words, key=lambda w: w.cy)
    lines: List[List] = []
    for w in ws:
        if lines and abs(w.cy - sum(v.cy for v in lines[-1]) / len(lines[-1])) < max(w.h, 10) * 0.6:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [" ".join(v.text for v in sorted(ln, key=lambda v: v.x)) for ln in lines]


# ---------------------------------------------------------------------------
# Interpretation
# ---------------------------------------------------------------------------

def _cands(cell: Optional[Cell], decimals: Optional[int], rupee: bool = False) -> List[Tuple[float, float]]:
    if cell is None:
        return []
    text, score, alts = cell
    out: List[Tuple[float, float]] = []
    texts = [(text, 0.0)]
    toks = text.split()
    if len(toks) > 1:                          # leading '₹' misread as '2' / '之' / 'き'
        texts.append((toks[-1], 0.05))
    if rupee:
        # '₹' printed tight against the figure and read as a digit: '₹7,144.00' -> '77,144.00',
        # '-₹41,914.40' -> '-341,914.40' (the arithmetic decides which reading is right)
        m = re.match(r"\s*(-?)\s*\d(?=\d|,\d)", text)
        if m:
            texts.append((m.group(1) + text[m.end():], 0.3))
    texts += [(a, 0.25) for a in alts]
    for t, base in texts:
        for v, c in number_candidates(t, decimals):
            if not any(abs(v - ov) < 1e-9 for ov, _ in out):
                out.append((v, base + c + (0.0 if score >= 0.9 else 0.1)))
    return sorted(out, key=lambda p: p[1])


def _is_blank_amount(text: str) -> bool:
    """'₹ -' style empty amount cells. A lone '₹' is often read as a single '2' / '之' / 'き'."""
    t = (text or "").strip()
    if not re.search(r"\d", t) or re.fullmatch(r"\W*[-—–]\W*", t):
        return True
    return len(re.sub(r"\D", "", t)) == 1 and len(t) <= 2


def _resolve(qc, rc, ac, amount_blank: bool) -> Tuple[float, float, float, str, List[str]]:
    q_opts = qc or [(0.0, 0.5)]
    r_opts = rc or [(0.0, 0.5)]
    a_opts = ([(0.0, 0.0)] if amount_blank else []) + (ac or [])
    if not a_opts:
        a_opts = [(0.0, 1.0)]
    best = None
    for (q, cq), (r, cr), (a, ca) in itertools.product(q_opts[:4], r_opts[:4], a_opts[:5]):
        cost = cq + cr + ca + (0.0 if money_eq(q * r, a, 0.011) else 5.0)
        if best is None or cost < best[0]:
            best = (cost, q, r, a)
    cost, q, r, a = best
    if cost >= 5.0 and r > 0 and a > 0:
        # The printed quantity disagrees with Amount / Rate (e.g. "2 Nos" but amount = 1.4 x rate) or is
        # blank (FHTC "(152 Nos)" only in the description): the sanctioned amount and rate govern,
        # so the quantity is taken as Amount / Rate - exactly as the TPIA abstract does.
        q2 = round(a / r, 3)
        if q2 > 0 and money_eq(q2 * r, a, 0.011):
            note = (f"Quantity taken as Amount / Rate = {q2:g} (printed quantity: {q:g})" if q
                    else f"Quantity blank in TS - taken as Amount / Rate = {q2:g}")
            return q2, r, a, "corrected", [note]
    if cost >= 5.0:
        clean_read = all(c and c[0][1] == 0 for c in (qc, rc, ac))
        why = ("as printed in the TS (the document's own figures disagree)" if clean_read
               else "check the reading")
        return q, r, a, "review", [f"{q:g} x {r:g} = {q * r:,.2f} but amount is {a:,.2f} - {why}"]
    return q, r, a, ("corrected" if cost > 0.3 else "ok"), []


_SUMMARY = [
    ("sub_total", SUBTOTAL_RE.pattern),
    ("grand_total", r"grand\s*total"),
    ("contingencies", r"contingen"),
    ("crore", r"\bin\s*cr\b|\bin\s*crore"),
    ("lakh", r"say\s*rs|in\s*lakh"),
    ("beneficiary_share", r"platform|beneficiary|community\s*contribution"),
    ("provision", r"provision\s+for"),
]


_TITLE_STOP = re.compile(r"\b(?:for|of|amount|estimate|sanction|statement|abstract|revised|technical|cost|"
                         r"jal|jeevan|mission|and|aa|ts)\b", re.I)


def _scheme_from_title(text: str) -> str:
    """
    'ABSTRACT OF REVISED ESTIMATE FOR NALIA PT-III PWSS' -> 'NALIA PT-III PWSS': the words in front of
    'PWSS' on the same line, back to the last title word ('for', 'estimate', 'amount', ...).
    """
    for line in text.splitlines():
        m = re.search(r"(.{0,80}?)\b(P\W{0,2}W\W{0,2}S\W{0,2}S|Piped Water Supply Scheme)\b", line, re.I)
        if not m:
            continue
        head = m.group(1)
        suffix = m.group(2) if len(m.group(2)) > 12 else "PWSS"
        stops = list(_TITLE_STOP.finditer(head))
        if stops:
            head = head[stops[-1].end():]
        head = re.sub(r"^[\s:;,\-–]+|[\s:;,\-–]+$", "", head)
        if re.search(r"[A-Za-z]{3}", head):
            return re.sub(r"\s+", " ", f"{head} {suffix}").strip().upper()
    return ""


def _complete_from_subtotal(items: List[TSItem], sub_total: float) -> None:
    """
    Lines whose reading is incomplete (a factor missing, or an amount unreadable) are completed only
    when the completed amounts then add up to the printed Sub Total exactly - two independent
    figures of the document agreeing, never a guess on its own.
      * one factor missing:  rate = Amount / Qty,  qty = Amount / Rate,  amount = Qty x Rate
      * one line still wrong: its amount = Sub Total - all other lines, accepted only if Qty x Rate agrees
    """
    review = [it for it in items if it.status == "review" and not it.heading]
    if not review:
        return
    proposal = {}
    for it in review:
        q, r, a = it.qty, it.rate, it.amount
        if a and q and not r and money_eq(q * round(a / q, 2), a):
            proposal[id(it)] = (q, round(a / q, 2), a, f"Rate not read - taken as Amount / Qty = {round(a / q, 2):g}")
        elif a and r and not q and money_eq(round(a / r, 3) * r, a):
            proposal[id(it)] = (round(a / r, 3), r, a, f"Quantity not read - taken as Amount / Rate = {round(a / r, 3):g}")
        elif q and r and not a:
            proposal[id(it)] = (q, r, round(q * r, 2), f"Amount not read - taken as Qty x Rate = {q * r:,.2f}")
    def amount(it):
        return proposal[id(it)][2] if id(it) in proposal else it.amount
    total = sum(amount(it) for it in items)
    left = [it for it in review if id(it) not in proposal]
    if len(left) == 1 and left[0].qty and left[0].rate:
        it = left[0]
        rest = total - amount(it)
        if money_eq(it.qty * it.rate, sub_total - rest, 0.011):
            proposal[id(it)] = (it.qty, it.rate, round(sub_total - rest, 2),
                                f"Amount misread - taken as Sub Total - other lines = {sub_total - rest:,.2f} "
                                f"(equals Qty x Rate)")
            total = sum(amount(i) for i in items)
            left = []
    if left or not money_eq(total, sub_total, 0.011):
        return                                  # the document does not confirm the completion: leave for review
    for it in review:
        it.qty, it.rate, it.amount, note = proposal[id(it)]
        it.status = "corrected"
        it.notes = [note + " - confirmed by the printed Sub Total"]


def parse_page_spec(spec: str) -> Optional[List[int]]:
    """'9-10' / '9,10' / '9' (1-based, as printed in a PDF viewer) -> [8, 9]; '' -> None (auto)."""
    spec = (spec or "").strip()
    if not spec:
        return None
    pages: List[int] = []
    for part in re.split(r"[,\s]+", spec):
        m = re.fullmatch(r"(\d+)(?:\s*-\s*(\d+))?", part)
        if not m:
            raise ValueError(f"Page list not understood: '{spec}' (use e.g. 9-10)")
        a = int(m.group(1))
        b = int(m.group(2) or a)
        pages.extend(range(a - 1, b))
    return pages


def parse_ts_pdf(pdf_source: PdfSource, pages: Optional[List[int]] = None) -> TSParseResult:
    """
    pages: 0-based page indexes of the abstract table. None = locate them automatically
    (only those pages are read in detail, so a 100-page TS document takes well under a minute).
    """
    from table_extractor import is_scanned_pdf, open_pdf
    auto = pages is None
    if auto:
        pages = find_abstract_pages(pdf_source) or None
    if pages:
        # decided on the abstract pages themselves: a scanned TS may carry a searchable covering letter
        pdf = open_pdf(pdf_source)
        scanned = any(len(pdf[p].get_textpage().get_text_range().strip()) < 100 for p in pages if p < len(pdf))
    else:
        scanned = is_scanned_pdf(pdf_source)
    rows, page_texts = (_rows_from_scanned if scanned else _rows_from_digital)(pdf_source, pages)
    meta = TSMetaData()
    items: List[TSItem] = []
    summary: Dict[str, float] = {}

    for row in rows:
        desc = (row.get("desc") or ("", 1.0, []))[0].strip()
        sl = re.sub(r"[^\dA-Za-z]", "", (row.get("sl") or ("", 1.0, []))[0])
        kind = next((k for k, pat in _SUMMARY if re.search(pat, desc, re.I)), None)
        amt_cell = row.get("amount")
        if kind:
            c = _cands(amt_cell, 2)
            if c:
                summary[kind] = c[0][0]
            continue
        if not desc or "sub_total" in summary:     # below the Sub Total: platform / contingency / signatures
            continue
        desc = re.sub(r"(?<=[a-z])\d+\.\d{2}\b", "", desc)   # quantity bled in from the next column ("safe6.00")
        qc = _cands(row.get("qty"), None)
        rc = _cands(row.get("rate"), 2)
        ac = _cands(amt_cell, 2, rupee=True)
        blank = _is_blank_amount(amt_cell[0] if amt_cell else "")
        q, r, a, status, notes = _resolve(qc, rc, ac, blank)
        unit = (row.get("unit") or ("", 1.0, []))[0].strip()
        if re.fullmatch(r"[jJlLiI1]ob", unit):    # 'J' read as 'l' / 'I' / '1'
            unit = "Job"
        if (r or a) and not _UNIT_OK.fullmatch(unit or "-"):
            # the unit printed close to the description is read into it ("...Deep Tube well RM using Rotary")
            m = _UNIT_IN_TEXT.search(desc)
            if m:
                unit = m.group(1)
                desc = re.sub(r"\s{2,}", " ", desc[:m.start()] + desc[m.end():]).strip()
        heading = q == 0 and r == 0 and a == 0    # e.g. "Cost of distribution network 0 0 0" (sub-items follow)
        items.append(TSItem(sl=sl, desc=desc, qty=q, rate=r, unit=unit, amount=a,
                            status=status, heading=heading, notes=notes,
                            uncertain=_txt(row, "_uncertain").split()))

    meta.sub_total = summary.get("sub_total")
    if meta.sub_total is not None:
        _complete_from_subtotal(items, meta.sub_total)
    if meta.sub_total is not None and money_eq(sum(it.amount for it in items), meta.sub_total, 0.011):
        # The printed amounts add up to the printed Sub Total, so the amount column is confirmed by the
        # document itself; a line whose printed Qty x Rate disagrees (1030.45 x 342 = 3,52,413.90 printed as
        # 3,52,413.38) keeps its sanctioned amount and the quantity becomes Amount / Rate (rule agreed for the
        # abstract: F = D x E must equal the TS amount).
        for it in items:
            if it.status == "review" and it.rate > 0 and it.amount > 0:
                printed = it.qty
                q = round(it.amount / it.rate, 3)
                it.qty = q if money_eq(q * it.rate, it.amount, 0.001) else it.amount / it.rate
                it.status = "corrected"
                it.notes = [f"TS prints {printed:g} x {it.rate:g} = {printed * it.rate:,.2f} but amount "
                            f"{it.amount:,.2f} (included in the printed Sub Total) - quantity taken as "
                            f"Amount / Rate = {it.qty:.4f}"]
    meta.grand_total = summary.get("grand_total")
    meta.contingencies = summary.get("contingencies")
    meta.beneficiary_share = summary.get("beneficiary_share")
    full = "\n".join(page_texts)
    m = re.search(r"ESTT\.?\s*COST\s*:?\s*([\d.]+)\s*Lakh", full, re.I)
    if m:
        meta.estimated_cost_lakhs = float(m.group(1))
    elif "lakh" in summary:
        meta.estimated_cost_lakhs = summary["lakh"]
    elif "crore" in summary:
        meta.estimated_cost_lakhs = round(summary["crore"] * 100, 2)
    names = [re.sub(r"\s+", " ", m.group(1)).strip(" :-")
             for m in re.finditer(r"of\s*Work\s*[:\-–]*\s*(.+)", full, re.I)]
    names = [n for n in names if re.search(r"pwss|scheme|water|\bpt\b", n, re.I)]
    meta.scheme = names[0] if names else _scheme_from_title(full)

    m = re.search(r"([A-Za-z]+)\s+Sub\s*-?\s*Division", full, re.I)
    if m:
        meta.sub_division = m.group(1).upper()
    # "Dhubri Division" - never the "Sub" of "Dhubri Sub-Division"
    divs = [m.group(1) for m in re.finditer(r"\b([A-Za-z]+)\s*-?\s*Division\b", full, re.I)
            if m.group(1).lower() not in ("sub", "the", "of", "and")]
    if divs:
        meta.division = divs[0].upper()

    result = TSParseResult(metadata=meta, items=items, checks=[], is_scanned=scanned, pages=pages or [])
    priced = result.priced_items
    review = [it for it in items if it.status == "review"]
    unreadable = [it for it in review if "check the reading" in " ".join(it.notes)]
    checks = [
        Check("TS abstract pages", bool(pages) or not auto,
              ("Read pages " + ", ".join(str(p + 1) for p in pages) + (" (found automatically)" if auto else ""))
              if pages else "Could not locate the abstract table automatically - all pages were read; "
                            "enter the page numbers to read only the abstract", "info" if pages else "warning"),
        Check("TS items found", len(priced) > 0, f"{len(priced)} sanctioned items read"),
        Check("Every TS line satisfies Quantity x Rate = Amount", not review,
              "All lines consistent" if not review else
              "; ".join(f"Sl {it.sl or '?'} {it.desc[:30]}: {it.notes[0]}" for it in review[:5]),
              # a discrepancy printed in the TS itself is worth knowing but is not a reading error
              "critical" if unreadable else "warning"),
    ]
    if meta.sub_total is not None:
        ok = money_eq(result.total_sanctioned_amount, meta.sub_total, 0.011)
        checks.append(Check("TS item amounts add up to the printed Sub Total", ok,
                            f"Sum {format_inr(result.total_sanctioned_amount)} vs printed {format_inr(meta.sub_total)}"))
    else:
        checks.append(Check("TS Sub Total found", False, "No 'Sub Total' line was found to verify the item amounts", "warning"))
    result.checks = checks
    return result
