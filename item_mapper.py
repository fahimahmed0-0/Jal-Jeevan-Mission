"""
item_mapper.py - Builds the TPIA Abstract from a digital contractor e-bill.

Two routes:

1. With the previous TPIA abstract (Excel template) of the same scheme  ->  abstract_from_template()
   Each e-bill line is placed on the template row(s) for which the bill's own figures hold exactly:
       rate              = template contract rate (E)
       amount this bill  = sum of the rows' L after the new quantities are entered
       amount up to date = sum of the rows' K + L   (K = what the template says was paid before)
       qty up to date    = H + I                    (single-milestone lines)
   No keyword guessing: a line either satisfies all identities on a row set, or it is reported as
   unmatched. The model is then recomputed by evaluating the template's own formulas, so what the
   app shows is exactly what Excel will calculate.

2. Without a template  ->  abstract_from_ebill()
   Items are built from the e-bill lines (milestone % = up-to-date amount / (qty x rate)), optionally
   completed with the sanctioned items and quantities of a Technical Sanction.

ts_benchmark() compares any abstract with a Technical Sanction (rates / quantities / missing items).
"""

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import openpyxl

from abstract_model import Abstract, AbstractItem, AbstractMeta, Check, MilestoneRow
from ebill_parser import EBillItem, EBillParseResult
from formula_eval import FormulaError, WorksheetEvaluator, formula_refs
from ts_parser import TSItem, TSParseResult
from utils import format_inr, money_eq

STANDARD_PCTS = (0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.6, 0.7, 0.75, 0.8, 0.9, 1.0)


# ---------------------------------------------------------------------------
# Template reading
# ---------------------------------------------------------------------------

@dataclass
class TemplateLayout:
    sheet: str
    start_row: int
    end_row: int                 # last item row
    total_row: int               # "Total Amount ..." row
    words_cell: Optional[Tuple[int, int]] = None
    recommend_cell: Optional[Tuple[int, int]] = None
    work_order_cell: Optional[Tuple[int, int]] = None


def _num(v) -> float:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def find_layout(ws) -> TemplateLayout:
    start = None
    for r in range(1, 40):
        a = str(ws.cell(r, 1).value or "").strip()
        if a in ("1", "1.0") or (isinstance(ws.cell(r, 1).value, (int, float)) and ws.cell(r, 1).value == 1):
            start = r
            break
    if start is None:
        raise ValueError("Template: could not find the first item row (Sl. No. 1 in column A).")
    total = None
    for r in range(start, min(ws.max_row, 400) + 1):
        if re.search(r"total", str(ws.cell(r, 2).value or ""), re.I):
            total = r
            break
    if total is None:
        raise ValueError("Template: could not find the 'Total Amount' row (column B).")
    lay = TemplateLayout(sheet=ws.title, start_row=start, end_row=total - 1, total_row=total)
    for r in range(total + 1, min(ws.max_row, total + 15) + 1):
        for c in range(1, 20):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("=") and "ROUNDUP" in v.upper():
                lay.recommend_cell = (r, c)
            elif isinstance(v, str) and re.search(r"in\s*words", v, re.I):
                lay.words_cell = (r, c)
            elif isinstance(v, str) and lay.work_order_cell is None and re.match(r"\s*work\s*order\s*amount", v, re.I):
                for cc in range(c + 1, 20):
                    if ws.cell(r, cc).value not in (None, ""):
                        lay.work_order_cell = (r, cc)
                        break
    return lay


def read_template(template_bytes: bytes, i_values: Optional[Dict[int, float]] = None,
                  sheet_name: Optional[str] = None) -> Tuple[List[AbstractItem], TemplateLayout, AbstractMeta, Optional[float]]:
    """
    Reads the template into items/rows, with column I optionally replaced by i_values
    {excel_row: qty}. All amounts come from evaluating the template's own formulas.
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=False)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets[0]
    lay = find_layout(ws)
    for r, q in (i_values or {}).items():
        ws.cell(r, 9).value = q
    ev = WorksheetEvaluator(ws)

    items: List[AbstractItem] = []
    current: Optional[AbstractItem] = None
    for r in range(lay.start_row, lay.end_row + 1):
        a = ws.cell(r, 1).value
        b = str(ws.cell(r, 2).value or "").strip()
        d_raw, e_raw = ws.cell(r, 4).value, ws.cell(r, 5).value
        g_raw = ws.cell(r, 7).value
        has_price = d_raw not in (None, "") or e_raw not in (None, "")
        sl = "" if a in (None, "") else (str(int(a)) if isinstance(a, float) and a.is_integer() else str(a)).strip()
        if sl or (b and has_price):
            if not has_price and g_raw in (None, ""):
                items.append(AbstractItem(sl=sl, desc=b, heading=True, excel_row=r))
                current = None
                continue
            current = AbstractItem(sl=sl, desc=b, unit=str(ws.cell(r, 3).value or "").strip(),
                                   qty=ev.number(r, 4), rate=ev.number(r, 5), excel_row=r)
            try:
                current.stated_amount = ev.number(r, 6)
            except FormulaError:
                current.stated_amount = None
            items.append(current)
        elif b and not has_price and g_raw in (None, ""):
            items.append(AbstractItem(sl="", desc=b, heading=True, excel_row=r))
            current = None
            continue
        if current is None or g_raw in (None, ""):
            continue
        l_formula = ws.cell(r, 12).value
        l_fixed = not (isinstance(l_formula, str) and f"I{r}" in formula_refs(l_formula))
        pct = ws.cell(r, 15).value
        row = MilestoneRow(
            remarks=str(ws.cell(r, 14).value or "").strip(),
            pct=float(pct) if isinstance(pct, (int, float)) else None,
            rate_paid=ev.number(r, 7), prev_qty=ev.number(r, 8), this_qty=ev.number(r, 9),
            prev_amt=ev.number(r, 11), this_amt=ev.number(r, 12),
            excel_row=r, l_fixed=l_fixed, source=f"template row {r}")
        current.rows.append(row)

    meta = AbstractMeta(
        scheme=str(ws.cell(4, 3).value or "").strip(),
        contractor=str(ws.cell(5, 3).value or "").strip(),
        work_order_no=str(ws.cell(6, 3).value or "").strip(),
        contract_value_text=str(ws.cell(7, 3).value or "").strip(),
        division=str(ws.cell(8, 3).value or "").strip(),
        sub_division=str(ws.cell(8, 13).value or "").strip(),
        status_text=str(ws.cell(9, 3).value or "WORK IN PROGRESS").strip(),
    )
    wo = None
    if lay.work_order_cell:
        try:
            wo = ev.number(*lay.work_order_cell) or None
        except FormulaError:
            wo = None
    return items, lay, meta, wo


# ---------------------------------------------------------------------------
# Route 1: e-bill + template
# ---------------------------------------------------------------------------

def _match_line(eb: EBillItem, items: List[AbstractItem], used: Set[int]) -> Tuple[Optional[List[MilestoneRow]], Dict[int, float], str]:
    """Finds the template rows for one e-bill line. Returns (rows, new I values, explanation)."""
    candidates = [it for it in items if not it.heading and it.rows and money_eq(it.rate, eb.rate)]
    candidates.sort(key=lambda it: (not money_eq(it.qty, eb.qty, 1e-6), it.excel_row or 0))
    found = []
    for it in candidates:
        free = [r for r in it.rows if r.excel_row not in used]
        options: List[List[MilestoneRow]] = [[r] for r in free]
        if len(free) > 1:
            options.append(free)                     # the line claims the whole item (all milestones)
        for opt in options:
            new_i: Dict[int, float] = {}
            l_sum = m_sum = 0.0
            valid = True
            for r in opt:
                if r.l_fixed:
                    l = r.this_amt                       # reconciliation formula, independent of I
                else:
                    if len(opt) == 1:
                        if r.rate_paid == 0:
                            valid = valid and money_eq(eb.since_amt, 0)
                            q = 0.0
                        else:
                            q = eb.since_amt / r.rate_paid
                    else:
                        q = it.qty - r.prev_qty          # each milestone completed
                    if q < -1e-9:
                        valid = False
                    new_i[r.excel_row] = q
                    l = q * r.rate_paid
                l_sum += l
                m_sum += r.prev_amt + l
            if not valid or not money_eq(l_sum, eb.since_amt, 0.011) or not money_eq(m_sum, eb.upto_amt, 0.011):
                continue
            qty_ok = True
            if len(opt) == 1 and not opt[0].l_fixed and eb.qty > 0:
                qty_ok = abs(opt[0].prev_qty + new_i[opt[0].excel_row] - eb.qty) < 1e-6
            found.append((not qty_ok, len(opt) > 1, opt[0].excel_row, opt, new_i, it))
    if not found:
        return None, {}, "no template row satisfies rate, amount-this-bill and amount-up-to-date"
    found.sort(key=lambda f: f[:3])
    _, _, _, opt, new_i, it = found[0]
    rows_txt = ", ".join(str(r.excel_row) for r in opt)
    why = f"rate {eb.rate:,.2f} = item Sl {it.sl or '-'}; rows {rows_txt}: this bill and up-to-date amounts match exactly"
    if found[0][0]:
        why += " (quantity differs from e-bill)"
    return opt, new_i, why


def abstract_from_template(ebill: EBillParseResult, template_bytes: bytes,
                           sheet_name: Optional[str] = None) -> Abstract:
    items, lay, meta, wo = read_template(template_bytes, sheet_name=sheet_name)
    used: Set[int] = set()
    i_values: Dict[int, float] = {}
    line_rows: Dict[int, List[int]] = {}
    unmatched: List[EBillItem] = []
    explanations: Dict[int, str] = {}
    for eb in ebill.items:
        rows, new_i, why = _match_line(eb, items, used)
        if rows is None:
            unmatched.append(eb)
            explanations[eb.sl] = why
            continue
        for r in rows:
            used.add(r.excel_row)
            r.ebill_ref = eb.sl
        i_values.update(new_i)
        line_rows[eb.sl] = [r.excel_row for r in rows]
        explanations[eb.sl] = why

    abstract = build_template_abstract(template_bytes, i_values, ebill, sheet_name)
    for it, r in abstract.all_rows():
        for sl, rows in line_rows.items():
            if r.excel_row in rows:
                r.ebill_ref = sl
                r.notes.append(f"E-bill line {sl}: {explanations[sl]}")
    abstract.source_checks.append(Check(
        "Every e-bill line placed on the template", not unmatched,
        f"{len(ebill.items) - len(unmatched)}/{len(ebill.items)} lines matched exactly"
        + ("" if not unmatched else " - unmatched: " + "; ".join(
            f"line {e.sl} ({e.desc[:40]}, rate {e.rate:,.2f}, this bill {format_inr(e.since_amt)})" for e in unmatched))))
    # rows that carry a this-bill amount in the template but are not on the e-bill
    stray = [r for _, r in abstract.all_rows() if r.ebill_ref is None and abs(r.this_amt) > 0.005]
    if stray:
        abstract.source_checks.append(Check(
            "No template amounts outside the e-bill", False,
            "Template rows with a this-bill amount not claimed on the e-bill: "
            + ", ".join(f"row {r.excel_row} ({format_inr(r.this_amt)})" for r in stray), "warning"))
    abstract.notes.append(f"Template rows {lay.start_row}-{lay.end_row}; {len(used)} rows billed from the e-bill.")
    return abstract


def build_template_abstract(template_bytes: bytes, i_values: Dict[int, float],
                            ebill: Optional[EBillParseResult] = None, sheet_name: Optional[str] = None) -> Abstract:
    """(Re)evaluates the template with the given column-I quantities into an Abstract."""
    items, lay, meta, wo = read_template(template_bytes, i_values, sheet_name)
    if ebill is not None:
        m = ebill.metadata
        meta.contractor = meta.contractor or m.contractor
        meta.work_order_no = meta.work_order_no or m.work_order_no
        meta.scheme = meta.scheme or m.scheme
        meta.bill_ref = m.bill_ref
        meta.bill_date = m.bill_date
    stated: Dict[str, float] = {}
    if ebill is not None and ebill.metadata.net_this_bill is not None:
        stated["claim_net"] = ebill.metadata.net_this_bill
    if ebill is not None and ebill.metadata.total_upto_date is not None:
        stated["claim_upto"] = ebill.metadata.total_upto_date
    abstract = Abstract(meta=meta, items=items, source_kind="ebill_template", work_order_amount=wo,
                        stated=stated, template_bytes=template_bytes, template_sheet=lay.sheet)
    if ebill is not None:
        abstract.source_checks = list(ebill.checks)
    return abstract


def rebuild_template_abstract(prev: Abstract, ebill: Optional[EBillParseResult],
                              edits: Dict[int, float]) -> Abstract:
    """
    Re-evaluates the template after the engineer changed this-bill quantities ({excel_row: qty}).
    Every amount is recomputed by the template's own formulas, so the screen equals the Excel file.
    """
    i_values = {r.excel_row: r.this_qty for _, r in prev.all_rows() if r.excel_row and not r.l_fixed}
    i_values.update({row: q for row, q in edits.items()})
    new = build_template_abstract(prev.template_bytes, i_values, ebill, prev.template_sheet)
    old = {r.excel_row: r for _, r in prev.all_rows()}
    for _, r in new.all_rows():
        o = old.get(r.excel_row)
        if o is not None:
            r.ebill_ref, r.notes, r.status = o.ebill_ref, list(o.notes), o.status
        if r.excel_row in edits and (o is None or abs(o.this_qty - edits[r.excel_row]) > 1e-12):
            r.status = "edited"
            r.notes.append(f"Edited by engineer: this-bill quantity = {edits[r.excel_row]:g}")
    new.meta, new.work_order_amount = prev.meta, prev.work_order_amount
    new.source_checks, new.notes, new.stated = prev.source_checks, prev.notes, prev.stated
    return new


# ---------------------------------------------------------------------------
# Route 2: e-bill without template (optionally completed from a TS)
# ---------------------------------------------------------------------------

def _snap_pct(x: float) -> Optional[float]:
    for p in STANDARD_PCTS:
        if abs(x - p) < 1e-6:
            return p
    return None


def _short_desc(desc: str, limit: int = 160) -> str:
    desc = re.sub(r"\s+", " ", desc).strip()
    if len(desc) <= limit:
        return desc
    cut = desc[:limit]
    return (cut[: cut.rfind(" ")] if " " in cut else cut) + " ..."


def _clean_qty(amount: float, rate_paid: float) -> Optional[float]:
    """amount / rate_paid if that quantity reproduces the amount to the paisa with <= 3 decimals."""
    if rate_paid <= 0:
        return 0.0 if money_eq(amount, 0) else None
    q = round(amount / rate_paid, 3)
    return q if money_eq(q * rate_paid, amount) else None


def _ts_match(rate: float, unit: str, ts_items: List[TSItem], taken: Set[int]) -> Optional[int]:
    best = None
    for i, t in enumerate(ts_items):
        if i in taken or t.heading or t.rate <= 0:
            continue
        diff = abs(t.rate - rate)
        if diff <= max(0.51, 0.001 * rate):
            score = (diff, unit.lower()[:2] != t.unit.lower()[:2])
            if best is None or score < best[0]:
                best = (score, i)
    return best[1] if best else None


def abstract_from_ebill(ebill: EBillParseResult, ts: Optional[TSParseResult] = None) -> Abstract:
    m = ebill.metadata
    meta = AbstractMeta(scheme=m.scheme, contractor=m.contractor, work_order_no=m.work_order_no,
                        bill_ref=m.bill_ref, bill_date=m.bill_date)
    groups: Dict[Tuple, List[EBillItem]] = {}
    order: List[Tuple] = []
    for eb in ebill.items:
        key = (round(eb.rate, 2), eb.unit.strip().lower(), re.sub(r"\s+", " ", eb.desc)[:80].lower())
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(eb)

    ts_items = ts.items if ts else []
    taken: Set[int] = set()
    built: List[Tuple[Optional[int], AbstractItem]] = []
    for key in order:
        lines = groups[key]
        first = lines[0]
        it = AbstractItem(sl="", desc=_short_desc(first.desc), unit=first.unit, rate=first.rate,
                          qty=max(e.qty for e in lines))
        ts_idx = _ts_match(first.rate, first.unit, ts_items, taken) if ts_items else None
        if ts_idx is not None:
            taken.add(ts_idx)
            t = ts_items[ts_idx]
            it.qty = t.qty
            it.desc = t.desc
            if t.unit and t.unit not in ("0",):
                it.unit = t.unit
            it.notes.append(f"Contract quantity from TS: {t.qty:g} {t.unit}")
            if not money_eq(t.rate, first.rate):
                it.notes.append(f"TS rate {t.rate:,.2f} vs billed rate {first.rate:,.2f}")
        for eb in lines:
            pct = None
            if eb.qty > 0 and eb.rate > 0:
                pct = _snap_pct(eb.upto_amt / (eb.qty * eb.rate))
            g = round(eb.rate * pct, 6) if pct else eb.rate
            k, l = eb.prev_amt, eb.since_amt
            h = _clean_qty(k, g)
            i = _clean_qty(l, g)          # may be negative: a measured reduction (e.g. -1.35 Qntl)
            row = MilestoneRow(
                remarks=("On Completion of the work in all respect." if pct in (None, 1.0)
                         else f"Milestone {pct * 100:g}% (as per e-bill line {eb.sl})"),
                pct=pct, rate_paid=g, prev_amt=k, this_amt=l, ebill_ref=eb.sl,
                source=f"e-bill line {eb.sl}")
            row.prev_qty = h if h is not None else (eb.qty if k > 0 else 0.0)
            row.this_qty = i if i is not None else 0.0
            if i is None:
                row.notes.append(f"This-bill amount {format_inr(l)} is an adjustment (not qty x rate) - carried as an amount")
            if h is None and k:
                row.notes.append(f"Previous amount {format_inr(k)} is not qty x rate - carried as an amount")
            it.rows.append(row)
        built.append((ts_idx, it))

    items: List[AbstractItem]
    if ts_items:
        # follow the TS order; unbilled sanctioned items appear with zero amounts
        by_ts = {idx: it for idx, it in built if idx is not None}
        items = []
        for i, t in enumerate(ts_items):
            if i in by_ts:
                by_ts[i].sl = t.sl
                items.append(by_ts[i])
            elif t.heading:
                items.append(AbstractItem(sl=t.sl, desc=t.desc, heading=True))
            else:
                items.append(AbstractItem(sl=t.sl, desc=t.desc, unit=t.unit, qty=t.qty, rate=t.rate,
                                          rows=[MilestoneRow(remarks="On Completion of the work in all respect.",
                                                             pct=1.0, rate_paid=t.rate, source="TS (not billed)")]))
        extra = [it for idx, it in built if idx is None]
        if extra:
            items.append(AbstractItem(desc="ITEMS BILLED BUT NOT FOUND IN THE TS:-", heading=True))
            items.extend(extra)
    else:
        items = [it for _, it in built]
        for n, it in enumerate(items, start=1):
            it.sl = str(n)

    stated: Dict[str, float] = {}
    if m.net_this_bill is not None:
        stated["claim_net"] = m.net_this_bill
    abstract = Abstract(meta=meta, items=items, source_kind="digital_ebill", stated=stated,
                        source_checks=list(ebill.checks))
    if m.deduct_previous is not None:
        ok = money_eq(abstract.total_prev, m.deduct_previous, 0.011)
        abstract.source_checks.append(Check("Previous-bill total matches the e-bill", ok,
                                            f"{format_inr(abstract.total_prev)} vs e-bill {format_inr(m.deduct_previous)}"))
    if ts_items:
        missing = [it for idx, it in built if idx is None]
        abstract.source_checks.append(Check("Every billed item found in the TS", not missing,
                                            "All billed items are sanctioned items" if not missing else
                                            "; ".join(f"{it.desc[:40]} (rate {it.rate:,.2f})" for it in missing[:5]),
                                            "warning"))
    return abstract


# ---------------------------------------------------------------------------
# TS benchmark
# ---------------------------------------------------------------------------

_SYNONYMS = {
    "rwpm": "raw water pumping main", "cwpm": "clear water pumping main", "dtw": "deep tube well",
    "fhtc": "functional household tap connection", "ugr": "under ground service reservoir",
    "esr": "elevated service reservoir", "tp": "treatment plant", "bfm": "bulk flow meter",
    "dg": "dg set", "hdpe": "hdpe pipe", "gi": "gi pipe", "piller": "pillar", "dozing": "dosing",
    "underground": "under ground", "toilet": "toilet unit", "submersible": "raw water pump",
    "centrifugal": "clear water pump",
}
_STOP = {"cost", "construction", "constructing", "supply", "supplying", "fitting", "fixing", "providing", "testing",
         "commissioning", "commisioniong", "commisioning", "installation", "including", "within", "with", "and",
         "the", "for", "from", "work", "job", "different", "size", "type", "nos", "each", "mm", "dia", "set", "of"}


def _tokens(text: str) -> Set[str]:
    words = re.findall(r"[a-z]+|\d+", re.sub(r"['’`]", "", text.lower()))     # OCR: "DT'W" -> "dtw"
    out: Set[str] = set()
    for w in words:
        for x in _SYNONYMS.get(w, w).split():
            if x not in _STOP and (len(x) >= 2 and not x.isdigit() or (x.isdigit() and x not in ("0", "1"))):
                out.add(x)
    return out


def _pair_score(it: AbstractItem, t: TSItem) -> float:
    score = 0.0
    if t.rate > 0 and abs(t.rate - it.rate) <= max(0.51, 0.001 * it.rate):
        score += 10.0
    if t.amount > 0 and money_eq(t.amount, it.amount, 1.0):
        score += 8.0
    a, b = _tokens(it.desc), _tokens(t.desc)
    if a and b:
        inter = len(a & b)
        score += 10.0 * inter / min(len(a), len(b)) + 5.0 * inter / len(a | b)
    return score


def ts_benchmark(abstract: Abstract, ts: TSParseResult) -> Tuple[List[Dict], List[Check]]:
    """
    Compares the abstract's items with the sanctioned items. Items are paired globally by the
    strongest evidence first (same rate, same amount, same wording), so one weak description match
    cannot steal the TS line that belongs to another item.
    """
    ts_items = [t for t in ts.items if not t.heading]
    ab_items = [it for it in abstract.items if not it.heading]
    pairs = sorted(((_pair_score(it, t), ai, ti) for ai, it in enumerate(ab_items) for ti, t in enumerate(ts_items)),
                   reverse=True)
    match: Dict[int, int] = {}
    used_ts: Set[int] = set()
    for score, ai, ti in pairs:
        if score < 6.0:
            break
        if ai in match or ti in used_ts:
            continue
        match[ai] = ti
        used_ts.add(ti)

    rows: List[Dict] = []
    for ai, it in enumerate(ab_items):
        if ai not in match:
            rows.append({"Sl": it.sl, "Abstract item": it.desc[:70], "Abstract qty": it.qty, "Abstract rate": it.rate,
                         "TS item": "", "TS qty": None, "TS rate": None, "TS amount": None, "Status": "⚠️ Not found in TS"})
            continue
        t = ts_items[match[ai]]
        notes = []
        if money_eq(t.rate, it.rate) and abs(t.qty - it.qty) < 1e-6:
            status = "✅ Matches TS"
        else:
            if abs(t.rate - it.rate) <= max(0.51, 0.001 * it.rate) and not money_eq(t.rate, it.rate):
                notes.append(f"rate rounded ({it.rate:,.2f} vs TS {t.rate:,.2f})")
            elif not money_eq(t.rate, it.rate):
                notes.append(f"rate {it.rate:,.2f} vs TS {t.rate:,.2f}")
            if abs(t.qty - it.qty) > 1e-6:
                notes.append(f"qty {it.qty:g} vs TS {t.qty:g}")
            amt_diff = it.amount - t.amount
            notes.append("same amount" if money_eq(amt_diff, 0, 1.0) else f"amount {amt_diff:+,.2f} vs TS")
            status = "ℹ️ " + "; ".join(notes)
        rows.append({"Sl": it.sl, "Abstract item": it.desc[:70], "Abstract qty": it.qty, "Abstract rate": it.rate,
                     "TS item": t.desc[:70], "TS qty": t.qty, "TS rate": t.rate, "TS amount": t.amount,
                     "Status": status})
    for ti, t in enumerate(ts_items):
        if ti not in used_ts:
            rows.append({"Sl": "", "Abstract item": "", "Abstract qty": None, "Abstract rate": None,
                         "TS item": t.desc[:70], "TS qty": t.qty, "TS rate": t.rate, "TS amount": t.amount,
                         "Status": "ℹ️ Sanctioned item not in abstract"})
    not_in_ts = [r for r in rows if r["Status"].startswith("⚠️")]
    differ = [r for r in rows if r["Abstract item"] and r["TS item"] and not r["Status"].startswith("✅")]
    checks = [
        Check("Every abstract item exists in the TS", not not_in_ts,
              "All items found" if not not_in_ts else f"{len(not_in_ts)} item(s) not found in TS", "warning"),
        Check("Abstract rates and quantities equal the TS", not differ,
              "All equal" if not differ else f"{len(differ)} item(s) differ - see TS benchmark tab", "info"),
    ]
    return rows, checks
