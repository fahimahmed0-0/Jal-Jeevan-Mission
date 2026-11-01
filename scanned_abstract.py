"""
scanned_abstract.py - Reads a scanned TPIA "Abstract of Bill (Payment Recommendation)" into an Abstract.

The scanned table is first turned into a cell grid by table_extractor. This module then
  1. identifies the columns from their header text (with the official column order as fallback),
  2. groups rows into items (merged Sl/Description/Rate cells) and milestone lines,
  3. resolves every number using the document's own arithmetic, so OCR mistakes are either
     corrected *with evidence* or flagged for review - never silently accepted:
        per line : Previous + This Bill = Upto Date,  This Bill = Qty x Rate paid,
                   Previous = Qty x Rate paid,         Rate paid = Rate x milestone %
        per item : Amount = Qty x Rate
        overall  : column sums = the printed "Total Amount" row
"""

import itertools
import logging
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from abstract_model import Abstract, AbstractItem, AbstractMeta, Check, MilestoneRow
from table_extractor import OCRWord, PdfSource, ScannedTable, extract_pdf_tables
from utils import format_inr, money_eq, number_candidates

logger = logging.getLogger(__name__)

STANDARD_PCTS = (0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.6, 0.7, 0.75, 0.8, 0.9, 1.0)

ROLE_ORDER_BEFORE = ["sl", "desc", "unit", "qty", "rate", "amount"]
ROLE_ORDER_AFTER = ["prev_qty", "this_qty", "upto_qty", "prev_amt", "this_amt", "upto_amt", "remarks"]


# ---------------------------------------------------------------------------
# Column roles
# ---------------------------------------------------------------------------

def _header_text(t: ScannedTable, col: int, last_header_row: int) -> str:
    return " ".join(t.text(r, col) for r in range(0, last_header_row + 1)).lower()


def _find_header_end(t: ScannedTable) -> int:
    """Index of the last header row (rows above the first line holding real numbers)."""
    keys = ("description", "remarks", "quantit", "rate", "upto", "previous", "amount", "unit")
    last = -1
    for r in range(min(t.n_rows, 12)):
        txt = " ".join(t.text(r, c) for c in range(t.n_cols)).lower()
        if sum(k in txt for k in keys) >= 2:
            last = r
    return last


def assign_roles(t: ScannedTable, prev_roles: Optional[Dict[str, int]] = None) -> Tuple[Dict[str, int], int]:
    """Maps role -> column index for one page. Returns (roles, last_header_row)."""
    hdr_end = _find_header_end(t)
    heads = [_header_text(t, c, hdr_end) if hdr_end >= 0 else "" for c in range(t.n_cols)]
    paid = next((c for c, h in enumerate(heads) if re.search(r"which|payment|been made", h)), None)
    if paid is None:
        if prev_roles and prev_roles.get("_ncols") == t.n_cols:
            return dict(prev_roles), hdr_end
        raise ValueError(f"Page {t.page_number}: could not find the 'Rate at which payment has been made' column")

    roles: Dict[str, int] = {"rate_paid": paid, "_ncols": t.n_cols}
    before = list(range(0, paid))
    optional = [c for c in before if re.search(r"weit|weight|wise|physical|progress", heads[c])]
    for c in optional:
        roles["weight" if re.search(r"weit|weight|wise", heads[c]) else "phys"] = c
    core = [c for c in before if c not in optional]
    names = list(ROLE_ORDER_BEFORE)
    if len(core) == len(names) - 1 and not re.search(r"\bsl\b|^no$|sl\.", heads[core[0]] if core else ""):
        names = names[1:]                    # Sl. No column not present on this page
    for name, c in zip(names[-len(core):] if len(core) < len(names) else names, core[-len(names):]):
        roles[name] = c
    after = list(range(paid + 1, t.n_cols))
    for name, c in zip(ROLE_ORDER_AFTER, after):
        roles[name] = c
    if len(after) > len(ROLE_ORDER_AFTER):
        roles["pct"] = after[len(ROLE_ORDER_AFTER)]
    return roles, hdr_end


# ---------------------------------------------------------------------------
# Value resolution
# ---------------------------------------------------------------------------

@dataclass
class _Field:
    text: str
    cands: List[Tuple[float, float]]     # (value, cost)
    score: float                          # OCR confidence (min over words)
    primary: Optional[float] = None       # the plain reading of the printed text (None if unreadable)

    @property
    def confident(self) -> bool:
        return self.primary is not None and self.score >= 0.9

    @property
    def literal(self) -> Optional[float]:
        return self.primary if self.confident else None


def _field(words: List[OCRWord], decimals: Optional[int] = 2, empty_cost: float = 0.8) -> _Field:
    text = " ".join(w.text for w in sorted(words, key=lambda w: (round(w.cy / 20), w.x))).strip()
    score = min((w.score for w in words), default=1.0)
    cands: List[Tuple[float, float]] = []
    seen = set()

    def add(v: float, c: float):
        k = round(v, 6)
        for i, (ov, oc) in enumerate(cands):
            if round(ov, 6) == k:
                if c < oc:
                    cands[i] = (v, c)
                return
        cands.append((v, c))

    primary = None
    for v, c in number_candidates(text, decimals):
        if primary is None and c <= 0.1:          # plain reading (a space standing for '.' counts as plain)
            primary = v
        add(v, c + (0.0 if score >= 0.9 else 0.15))
    for w in words:
        for alt in w.alts:
            if alt in seen:
                continue
            seen.add(alt)
            for v, c in number_candidates(alt, decimals):
                add(v, 0.25 + c)
    if not text:
        add(0.0, empty_cost)
    cands.sort(key=lambda p: p[1])
    return _Field(text, cands[:5], score, primary)


def _std_pct(rate_paid: float, rate: float) -> Optional[float]:
    """The standard milestone % for which rate x % equals rate_paid to the paisa."""
    if rate <= 0:
        return None
    for p in STANDARD_PCTS:
        if money_eq(rate * p, rate_paid):
            return p
    return None


def _resolve_line(G: _Field, H: _Field, I: _Field, K: _Field, L: _Field, M: _Field,
                  rate: float, printed_pct: Optional[float],
                  alternatives: Optional[List[Dict[str, float]]] = None) -> Tuple[Dict[str, float], str, List[str], List[str]]:
    """
    Picks the most plausible reading of one milestone line. Returns (values, status, notes, unconfirmed).
    Other self-consistent readings are appended to `alternatives` (best first).
    status 'ok'        - literal readings satisfy Prev + This = Upto
           'corrected' - a repaired/derived value was needed and is backed by another identity
           'review'    - no reading is backed by the document's arithmetic
    """
    def opts(f: _Field, derived: Sequence[Tuple[float, float]] = ()):
        o = list(f.cands)
        for v, c in derived:
            if not any(abs(v - ov) < 1e-6 for ov, _ in o):
                o.append((v, c))
        return o or [(0.0, 1.0)]

    g_opts = opts(G)
    if rate > 0:
        pcts = [printed_pct] if printed_pct else []
        pcts += [p for p in STANDARD_PCTS if p not in pcts]
        for p in pcts:
            exact = round(rate * p, 2)
            # printed rate-paid is often shown rounded (e.g. 27818.7 for 27818.67): accept the exact value
            if any(abs(exact - v) <= 0.051 for v, _ in g_opts):
                g_opts.append((exact, 0.05))
                break
    best = None
    consistent: List[Tuple] = []
    if alternatives is None:
        alternatives = []
    for (g, cg), (h, ch), (i, ci), (k, ck) in itertools.product(g_opts, opts(H, ()), opts(I, ()), opts(K)):
        l_opts = opts(L, [(round(i * g, 2), 1.4)])
        for l, cl in l_opts:
            m_opts = opts(M, [(round(k + l, 2), 1.4)])
            for m, cm in m_opts:
                cost = cg + ch + ci + ck + cl + cm
                if not money_eq(k + l, m, 0.006):
                    cost += 6.0
                if not money_eq(i * g, l, 0.006):
                    cost += 0.9
                if not money_eq(h * g, k, 0.006):
                    cost += 0.6
                if rate > 0 and g > 0:
                    p = _std_pct(g, rate)
                    if p is None:
                        cost += 0.5
                    elif printed_pct and abs(p - printed_pct) > 1e-6:
                        cost += 0.3
                if best is None or cost < best[0]:
                    best = (cost, g, h, i, k, l, m)
                if money_eq(k + l, m, 0.006):
                    consistent.append((cost, g, h, i, k, l, m))
    cost, g, h, i, k, l, m = best
    vals = dict(rate_paid=g, prev_qty=h, this_qty=i, prev_amt=k, this_amt=l, upto=m)
    # other readings that also satisfy Prev + This = Upto (used if the column totals disagree)
    seen = {(round(k, 2), round(l, 2))}
    for c in sorted(consistent):
        key = (round(c[4], 2), round(c[5], 2))
        if key not in seen and len(alternatives) < 6:
            seen.add(key)
            alternatives.append(dict(rate_paid=c[1], prev_qty=c[2], this_qty=c[3], prev_amt=c[4], this_amt=c[5], upto=c[6]))
    if not money_eq(k + l, m, 0.006):
        return vals, "review", [f"Prev {format_inr(k)} + This {format_inr(l)} does not equal Upto {format_inr(m)}"], \
            ["prev amount", "this amount", "upto amount"]

    fields = {"rate paid": (G, g), "prev qty": (H, h), "this qty": (I, i),
              "prev amount": (K, k), "this amount": (L, l), "upto amount": (M, m)}
    # Evidence for each value other than the reading itself
    support = {
        "rate paid": _std_pct(g, rate) is not None,
        "prev qty": money_eq(h * g, k, 0.006),
        "this qty": money_eq(i * g, l, 0.006),
        "prev amount": money_eq(h * g, k, 0.006) or (L.confident and M.confident and L.primary == l and M.primary == m),
        "this amount": money_eq(i * g, l, 0.006) or (K.confident and M.confident and K.primary == k and M.primary == m),
        "upto amount": K.confident and L.confident and K.primary == k and L.primary == l,
    }
    notes: List[str] = []
    repaired, unconfirmed = [], []
    for name, (f, val) in fields.items():
        as_read = f.primary is not None and abs(f.primary - val) < 1e-9
        if not as_read:
            repaired.append(name)
            notes.append(f"{name}: read '{f.text or '(blank)'}' -> {val:,.2f}")
        if (not as_read or not f.confident) and not support[name]:
            unconfirmed.append(name)
    if unconfirmed:
        return vals, "review", notes + [f"unconfirmed: {', '.join(unconfirmed)}"], unconfirmed
    return vals, ("corrected" if repaired else "ok"), notes, []


def _resolve_item(D: _Field, E: _Field, F: _Field, row_rates: List[float]) -> Tuple[float, float, Optional[float], str, List[str]]:
    """Picks qty/rate/amount for an item. Rate is cross-checked against the rows' rate-paid values."""
    best = None
    for (d, cd), (e, ce), (f, cf) in itertools.product(D.cands or [(0.0, 1.0)], E.cands or [(0.0, 1.0)],
                                                       F.cands or [(0.0, 1.0)]):
        cost = cd + ce + cf
        if not money_eq(d * e, f, 0.051):     # printed amounts are sometimes shown to 1 decimal
            cost += 2.0
        if row_rates and e > 0 and not all(_std_pct(g, e) is not None or abs(g - e) < 0.051 for g in row_rates if g):
            cost += 1.0
        if best is None or cost < best[0]:
            best = (cost, d, e, f)
    cost, d, e, f = best
    status, notes = "ok", []
    if cost >= 2.0:
        status = "review"
        notes.append(f"Qty {d:g} x Rate {e:,.2f} = {d * e:,.2f} but amount reads {f:,.2f}")
    elif cost > 0:
        status = "corrected"
    return d, e, (f if F.cands else None), status, notes


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

_UNIT_FIX = {"JO8": "JOB", "J0B": "JOB", "JQB": "JOB", "J08": "JOB", "R.M": "RM", "NO.S": "Nos", "NOS": "Nos"}


def _clean_unit(u: str) -> str:
    u = u.strip().strip(".")
    return _UNIT_FIX.get(u.upper(), u)


def _clean_remarks(text: str) -> Tuple[str, Optional[float]]:
    """Removes a milestone % that bled in from the narrow % column; returns (remarks, pct)."""
    pct = None
    for m in re.finditer(r"(?<![\w.])(\d{1,3})\s*%", text):
        pct = int(m.group(1)) / 100.0
    text = re.sub(r"(?<![\w.])\d{1,3}\s*%", " ", text)
    text = re.sub(r"\s+\d{1,3}\s*$", "", text)          # truncated '100%' read as '1'
    return re.sub(r"\s+", " ", text).strip(), pct


def _is_sub_item(desc: str) -> bool:
    return bool(re.match(r"^\s*[A-Za-z]\s*\)", desc))


def _words_text(ws: List[OCRWord]) -> str:
    return " ".join(w.text for w in sorted(ws, key=lambda w: (round(w.cy / 25), w.x))).strip()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _parse_meta(first: ScannedTable) -> AbstractMeta:
    meta = AbstractMeta()
    words = [w for w in first.page_words if w.cy < first.top + 5]
    labels = [
        ("scheme", r"scheme"), ("contractor", r"contrac"), ("work_order_no", r"work\s*order"),
        ("contract_value_text", r"contract\s*value"), ("division", r"^name\s*of\s*division|^division"),
        ("sub_division", r"sub-?\s*div"), ("status_text", r"completion|status"),
    ]
    for key, pat in labels:
        lab = next((w for w in words if re.search(pat, w.text, re.I)), None)
        if not lab:
            continue
        right = [w for w in words if w is not lab and w.x > lab.x + lab.w - 5 and abs(w.cy - lab.cy) < max(lab.h, w.h) * 0.7
                 and not re.search(r"name of|contrac|order no|value:|division|completion", w.text, re.I)]
        if not right:
            continue
        val = min(right, key=lambda w: w.x)
        text = val.text
        if val.score < 0.85 and val.alts:
            text = max(val.alts + [text], key=lambda s: sum(ch.isalpha() for ch in s))
        text = re.sub(r"^MIS\b", "M/S", text.strip(" :-"))
        setattr(meta, key, text)
    status = next((w.text for w in words if re.search(r"work\s+in\s+progress|completed", w.text, re.I)), None)
    if status:
        meta.status_text = status.upper()
    return meta


def parse_scanned_abstract(pdf_source: PdfSource, tables: Optional[List[ScannedTable]] = None) -> Abstract:
    tables = tables if tables is not None else extract_pdf_tables(pdf_source)
    if not tables:
        raise ValueError("The PDF has no pages.")
    meta = _parse_meta(tables[0])
    items: List[AbstractItem] = []
    stated: Dict[str, float] = {}
    work_order_amount: Optional[float] = None
    notes: List[str] = []
    roles: Optional[Dict[str, int]] = None
    row_alternatives: Dict[int, List[Dict[str, float]]] = {}

    for t in tables:
        roles, hdr_end = assign_roles(t, roles)
        R = roles
        r = hdr_end + 1
        while r < t.n_rows:
            row_all = " ".join(t.text(r, c) for c in range(t.n_cols)).strip()
            if not row_all:
                r += 1
                continue
            low = row_all.lower()
            # summary rows
            if re.search(r"total\s+amount", low):
                for key, role in (("contract", "amount"), ("prev", "prev_amt"), ("this", "this_amt"), ("upto", "upto_amt")):
                    if role in R:
                        f = _field(t.words(r, R[role]))
                        if f.cands:
                            stated[key] = f.cands[0][0]
                r += 1
                continue
            if re.search(r"work\s*order\s*amount", low):
                f = _field(t.words(r, R.get("amount", 0)))
                if f.cands:
                    work_order_amount = f.cands[0][0]
                r += 1
                continue

            # the item group = rows covered by the merged Rate cell (fallback: Description cell)
            key_col = R.get("rate", R.get("desc"))
            a, b = t.span(r, key_col)
            a, b = max(a, r), max(b, r)
            group = list(range(a, b + 1))
            g_rows = [x for x in group if _field(t.words(x, R["rate_paid"])).cands and t.words(x, R["rate_paid"])]

            def span_words(role):
                if role not in R:
                    return []
                ws: List[OCRWord] = []
                for x in group:
                    ws.extend(t.words(x, R[role]))
                return ws

            sl_txt = _words_text(span_words("sl"))
            desc = _words_text(span_words("desc"))
            if not g_rows:
                # heading line ("10 Supplying, fitting ...", "OTHER WORKS:-") - no payment values
                if desc or sl_txt:
                    items.append(AbstractItem(sl=re.sub(r"\D", "", sl_txt), desc=desc, heading=True))
                r = b + 1
                continue

            item = AbstractItem(sl=re.sub(r"[^\dA-Za-z]", "", sl_txt), desc=desc,
                                unit=_clean_unit(_words_text(span_words("unit"))))
            row_rates = []
            raw_lines = []
            for x in g_rows:
                G = _field(t.words(x, R["rate_paid"]))
                row_rates.append(G.cands[0][0] if G.cands else 0.0)
                raw_lines.append(x)
            D = _field(span_words("qty"), decimals=None)
            E = _field(span_words("rate"))
            F = _field(span_words("amount"))
            d, e, f, istatus, inotes = _resolve_item(D, E, F, row_rates)
            item.qty, item.rate, item.stated_amount = d, e, f
            item.notes += inotes

            # A value printed slightly off its line can land in a neighbouring (value-less) grid row
            # of the same item. Hand it to the nearest payment line whose cell is empty.
            value_roles = [k for k in ("prev_qty", "this_qty", "upto_qty", "prev_amt", "this_amt", "upto_amt", "pct")
                           if k in R]
            cellw: Dict[Tuple[int, str], List[OCRWord]] = {(x, k): list(t.words(x, R[k])) for x in group for k in value_roles}
            for x in group:
                if x in raw_lines:
                    continue
                for k in value_roles:
                    if not cellw[(x, k)]:
                        continue
                    for target in sorted(raw_lines, key=lambda g: (abs(g - x), g > x)):
                        if not cellw[(target, k)]:
                            cellw[(target, k)] = cellw[(x, k)]
                            cellw[(x, k)] = []
                            break

            # Remarks wrap over several grid rows: text above the first payment line belongs to it,
            # text below a payment line belongs to that line.
            rem: Dict[int, List[OCRWord]] = {x: [] for x in raw_lines}
            if "remarks" in R:
                for x in group:
                    owner = max((g for g in raw_lines if g <= x), default=raw_lines[0])
                    rem[owner] = rem[owner] + list(t.words(x, R["remarks"]))

            for x in raw_lines:
                remarks, pct_in_rem = _clean_remarks(_words_text(rem[x]))
                pct_txt = _words_text(cellw[(x, "pct")]) if "pct" in R else ""
                printed_pct = None
                m = re.search(r"(\d{1,3})\s*%", pct_txt)
                if m:
                    printed_pct = int(m.group(1)) / 100.0
                printed_pct = printed_pct or pct_in_rem
                fields = {"rate_paid": _field(t.words(x, R["rate_paid"]))}
                for k in ("prev_qty", "this_qty", "prev_amt", "this_amt", "upto_amt"):
                    fields[k] = _field(cellw.get((x, k), []), decimals=(None if k.endswith("qty") else 2))
                alts: List[Dict[str, float]] = []
                vals, status, lnotes, unconfirmed = _resolve_line(
                    fields["rate_paid"], fields["prev_qty"], fields["this_qty"],
                    fields["prev_amt"], fields["this_amt"], fields["upto_amt"], e, printed_pct, alts)
                if istatus == "review":
                    status = "review"
                    lnotes = lnotes + item.notes
                    unconfirmed = unconfirmed + ["contract qty/rate"]
                pct = printed_pct
                if e > 0 and vals["rate_paid"] > 0:
                    p = _std_pct(vals["rate_paid"], e)
                    pct = p if p is not None else (printed_pct or round(vals["rate_paid"] / e, 6))
                row = MilestoneRow(
                    remarks=remarks, pct=pct, rate_paid=vals["rate_paid"], prev_qty=vals["prev_qty"],
                    this_qty=vals["this_qty"], prev_amt=vals["prev_amt"], this_amt=vals["this_amt"],
                    stated_upto=vals["upto"], status=status, notes=lnotes, unconfirmed=unconfirmed,
                    source=f"page {t.page_number}, table row {x + 1}")
                item.rows.append(row)
                row_alternatives[id(row)] = alts
            items.append(item)
            r = b + 1

    _infer_missing_sl(items)
    _reconcile_with_totals(items, stated, notes, row_alternatives)

    # footer: payment recommended / words (last page, below the table)
    last = tables[-1]
    below = sorted([w for w in last.page_words if w.cy > last.bottom - 5], key=lambda w: (round(w.cy / 30), w.x))
    for i, w in enumerate(below):
        if re.search(r"recommend", w.text, re.I):
            for v in below[i:i + 4]:
                c = number_candidates(re.sub(r"[^\d.,]", "", v.text), 2)
                if c and c[0][0] > 0:
                    stated["recommended"] = c[0][0]
                    break
    abstract = Abstract(meta=meta, items=items, source_kind="scanned_abstract",
                        work_order_amount=work_order_amount, stated=stated, notes=notes)
    abstract.source_checks = [Check("Document type", True, "Scanned TPIA Abstract of Bill - read with grid OCR", "info")]
    return abstract


def _infer_missing_sl(items: List[AbstractItem]) -> None:
    """Fills a missing Sl. No. when it is the only number between two numbered neighbours."""
    numbered = [(i, int(it.sl)) for i, it in enumerate(items) if it.sl.isdigit()]
    for (i1, s1), (i2, s2) in zip(numbered, numbered[1:]):
        gap = [k for k in range(i1 + 1, i2)
               if not items[k].sl and not _is_sub_item(items[k].desc)
               and not re.search(r":-?\s*$", items[k].desc)]          # section titles like "OTHER WORKS:-"
        if s2 - s1 - 1 == len(gap):
            for n, k in enumerate(gap, start=1):
                items[k].sl = str(s1 + n)
                # the number may have been read into the description cell
                items[k].desc = re.sub(rf"\s\b{s1 + n}\b\s", " ", items[k].desc).strip()


def _reconcile_with_totals(items: List[AbstractItem], stated: Dict[str, float], notes: List[str],
                           alternatives: Optional[Dict[int, List[Dict[str, float]]]] = None) -> None:
    """
    1. If the column sums miss the printed totals row, look for the single line whose alternative
       (also self-consistent) reading closes all three gaps exactly; adopt it only if it is unique.
    2. If every column total then matches, amounts that a line could not confirm on its own are
       confirmed by the totals (quantities and rates are not - totals say nothing about them).
    """
    rows = [r for it in items for r in it.rows]
    if not rows or not all(k in stated for k in ("prev", "this", "upto")):
        return

    def gaps():
        return (stated["prev"] - sum(r.prev_amt for r in rows), stated["this"] - sum(r.this_amt for r in rows),
                stated["upto"] - sum(r.upto_amt for r in rows))

    dk, dl, dm = gaps()
    if alternatives and not (money_eq(dk, 0, 0.011) and money_eq(dl, 0, 0.011) and money_eq(dm, 0, 0.011)):
        fixes = []
        for r in rows:
            for alt in alternatives.get(id(r), []):
                if (money_eq(alt["prev_amt"] - r.prev_amt, dk, 0.011) and money_eq(alt["this_amt"] - r.this_amt, dl, 0.011)
                        and money_eq(alt["upto"] - r.upto_amt, dm, 0.011)):
                    fixes.append((r, alt))
                    break
        if len(fixes) == 1:
            r, alt = fixes[0]
            before = (r.prev_amt, r.this_amt)
            r.rate_paid, r.prev_qty, r.this_qty = alt["rate_paid"], alt["prev_qty"], alt["this_qty"]
            r.prev_amt, r.this_amt, r.stated_upto = alt["prev_amt"], alt["this_amt"], alt["upto"]
            r.status, r.unconfirmed = "corrected", []
            r.notes.append(f"Re-read to match the printed column totals: previous/this bill "
                           f"{before[0]:,.2f}/{before[1]:,.2f} -> {r.prev_amt:,.2f}/{r.this_amt:,.2f}")
            notes.append("One line was re-read so that the column sums equal the printed totals row.")

    dk, dl, dm = gaps()
    if money_eq(dk, 0, 0.011) and money_eq(dl, 0, 0.011) and money_eq(dm, 0, 0.011):
        amount_fields = {"prev amount", "this amount", "upto amount"}
        for r in rows:
            if (r.status == "review" and r.unconfirmed and set(r.unconfirmed) <= amount_fields
                    and money_eq(r.upto_amt, r.stated_upto if r.stated_upto is not None else r.upto_amt, 0.006)):
                r.status = "corrected"
                r.unconfirmed = []
                r.notes.append("confirmed by the printed column totals")
        notes.append("Column totals of Previous / This Bill / Upto Date match the printed totals row exactly.")
