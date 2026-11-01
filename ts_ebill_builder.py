"""
ts_ebill_builder.py - TPIA Abstract of Bill from the Technical Sanction + the contractor's e-bill.

  TS  "Abstract of Cost Estimate for Revised Technical Sanction"
        -> columns A-F (items, quantity, rate, amount) in the standard TPIA order, split into the
           JJM payment milestones (column G = rate x milestone %, remarks N, % O)
  E-bill (Running Account Bill-C)
        -> one line per milestone already claimed: rate = full TS rate,
           milestone % = Up-to-date / (Qty x Rate), previous = Up-to-date - Since previous bill
        -> columns H/K (previous bill) and I/L (this bill)

Recovery lines (negative "since previous bill", e.g. a rate reduced in the revised TS) are written the
way the TPIA abstract does: previous amount as paid, this bill = F - K (whole item) or D x G - K
(one milestone), so that the up-to-date amount equals the e-bill.
"""

import re
from dataclasses import dataclass
from itertools import combinations
from typing import Dict, List, Optional, Sequence, Tuple

from abstract_model import Abstract, AbstractItem, AbstractMeta, Check, MilestoneRow
from ebill_parser import EBillItem, EBillParseResult
from ts_parser import TSItem, TSParseResult
from utils import format_inr, money_eq

DONE = "On Completion of the work in all respect."

# Milestone schedules (remarks, fraction of the rate)
SCHEDULES: Dict[str, List[Tuple[str, float]]] = {
    "pumping_main": [("1. On Supply & Delivery of Pipe =", 0.30), ("2. On Completion of Laying =", 0.60),
                     ("3. On completion of Hydraulic testing & Commissioning =", 0.10)],
    "pump": [("1. On Supply & Delivery of Pumps =", 0.40), ("2. On Installation & Commissioning =", 0.60)],
    "ugr": [("1. Upto Completion of Sump Wall =", 0.40),
            ("2. After Completion of sump roof and commissioning thereof =", 0.60)],
    "esr": [("1. On completion of Foundation works =", 0.15), ("2. On Completion of the staging =", 0.25),
            ("3. On completion of the installation of GRP Tank =", 0.40), ("4. On commissioning of the ESR =", 0.20)],
    "hdpe": [("On Supply & Delivery of Pipe =", 0.20), ("On Completion of laying =", 0.50),
             ("On completion of FHTC =", 0.20), ("On completion of Hydraulic testing & Commissioning =", 0.10)],
    "pump_house": [("1. On completion upto Lintel level", 0.40), ("2. On completion of Remaining Works", 0.60)],
    "power": [("On Supply & Delivery of Transformer & Installation", 0.80), ("On Commissioning", 0.20)],
    "single": [(DONE, 1.0)],
}


@dataclass
class Component:
    key: str
    pattern: str
    schedule: str = "single"
    group: Optional[str] = None          # "hdpe" / "valve": sub-items under one numbered heading


# Order = the standard TPIA abstract layout. Rules are tried top to bottom (first match wins), so the
# more specific ones come first (e.g. "DTW protector" before "DTW", "electrification of pump house"
# before "pump house", "drain around boundary wall" before "boundary wall").
MAIN = ["dtw", "rwpm", "cwpm", "sub_pump", "cent_pump", "tp", "ugr", "fhtc", "esr", "hdpe", "valve"]
OTHER = ["interconnection", "bfm", "pump_house", "sign_board", "electrification", "boundary_wall", "chemical",
         "iron_gate", "site_dev", "power", "dtw_protector", "soil", "toilet", "plinth", "trial_run", "gi_pipe",
         "campus_light", "drains", "pathway", "site_camp", "soak_pit"]
RULES: List[Component] = [
    Component("dtw_protector", r"protector"),
    Component("toilet", r"toilet"),                      # "toilet unit at TP site" is not the TP
    Component("interconnection", r"connection\s+between|inter\s*-?\s*connection"),
    Component("electrification", r"electrif"),
    Component("drains", r"\bdrains?\b"),
    Component("soak_pit", r"soak\s*-?\s*pit"),
    Component("campus_light", r"campus"),
    Component("site_camp", r"\bcamp\b|site\s+clearance"),
    Component("rwpm", r"raw\s*water\s*pumping\s*main|\brwpm\b", "pumping_main"),
    Component("cwpm", r"clear\s*water\s*pumping\s*main|\bcwpm\b", "pumping_main"),
    Component("sub_pump", r"submers|raw\s*water\s*pump", "pump"),
    Component("cent_pump", r"centrifu|clear\s*water\s*pump", "pump"),
    Component("chemical", r"chemical|dosing|dozing"),
    # "Power connection at Treatment Plant" is the power item, not the TP
    Component("power", r"apdcl|power\s*line|power\s*connection|sub\s*-?\s*station|transformer", "power"),
    Component("tp", r"treatment\s*plant|sand\s*filter|\btp\b"),
    Component("ugr", r"under\s*ground|\bugr\b|\bsump\b", "ugr"),
    Component("fhtc", r"fhtc|household\s*tap|tap\s*connection"),
    Component("esr", r"elevated|\besr\b|staging", "esr"),
    Component("hdpe", r"hdpe|diameter\s*-|distribution\s*net\w*\s.*\bdia\b", "hdpe", group="hdpe"),
    Component("valve", r"valve\s*chamber", group="valve"),
    Component("bfm", r"bulk|flow\s*meter|\bbfm\b"),
    Component("pump_house", r"pump\s*house", "pump_house"),
    Component("boundary_wall", r"boundary"),
    Component("iron_gate", r"\bgate\b|heavy\s*duty\s*iron"),     # text is sometimes cut before "gate"
    Component("site_dev", r"site\s*develop"),
    Component("soil", r"soil"),
    Component("plinth", r"plinth"),
    Component("trial_run", r"trial"),
    Component("gi_pipe", r"\bg\.?\s*i\.?\s*pipe"),
    Component("pathway", r"pathway"),
    Component("sign_board", r"sign\s*bo"),
    Component("dtw", r"tube\s*-?\s*well|\bd\.?t\.?\s*'?w\b"),
]
GROUP_TITLES = {
    "hdpe": "Supplying, fitting, fixing, testing and commissioning of distribution network (HDPE pipes)",
    "valve": "Construction, testing and commissioning of valve chambers for sluice valves of different sizes",
}


def classify(desc: str) -> Optional[Component]:
    d = re.sub(r"['’`]", "", desc.lower())
    for comp in RULES:
        if re.search(comp.pattern, d):
            return comp
    return None


def _clean_desc(desc: str) -> str:
    d = re.sub(r"^\s*\(?[a-z]\)\s*", "", desc.strip(), flags=re.I)      # leading "(a)" / "(b)" markers
    return re.sub(r"\s+", " ", d).strip()


def _norm_unit(u: str) -> str:
    u = re.sub(r"[^a-z]", "", (u or "").lower())
    return {"no": "nos", "nos": "nos", "each": "nos", "rm": "rm", "rmt": "rm", "m": "rm", "sqm": "sqm",
            "cum": "cum", "job": "job", "qntl": "qtl", "qtl": "qtl", "quintal": "qtl"}.get(u, u)


def _clean_qty(amount: float, g: float) -> Optional[float]:
    """amount / g when that quantity (3 decimals) reproduces the amount to the paisa."""
    if g <= 0:
        return 0.0 if money_eq(amount, 0) else None
    q = round(amount / g, 3)
    return q if money_eq(q * g, amount) else None


# ---------------------------------------------------------------------------
# Layout from the TS
# ---------------------------------------------------------------------------

def _schedule_rows(it: AbstractItem, t: TSItem, comp: Optional[Component]) -> None:
    for remarks, pct in SCHEDULES[comp.schedule if comp else "single"]:
        it.rows.append(MilestoneRow(remarks=remarks, pct=pct, rate_paid=t.rate * pct,
                                    source="TS" + (f" Sl {t.sl}" if t.sl else "")))


def _layout_copy_ts(ts: TSParseResult) -> Tuple[List[AbstractItem], List[str]]:
    """
    The TS abstract copied as printed: same rows in the same order, the TS's own Sl numbers, headings,
    unnumbered sub-items and zero-amount lines, the text exactly as read. Each priced line gets its
    payment milestone rows; sub-items take their milestones from their heading ("Laying of
    distribution network" -> HDPE pipe milestones).
    """
    notes: List[str] = []
    items: List[AbstractItem] = []
    parent: Optional[str] = None
    unknown: List[str] = []
    for n, t in enumerate(ts.items):
        desc = re.sub(r"\s+", " ", t.desc).strip()
        nxt = ts.items[n + 1] if n + 1 < len(ts.items) else None
        # a title line: no unit ("Deep Tube Well & Treatment Plant"), or a numbered line whose unnumbered
        # sub-items follow ("18 Valve chamber including cost of sluice valve." -> "80 mm dia 6 Nos")
        title = t.heading and (not t.unit or (t.sl and nxt is not None and not nxt.sl))
        if title:
            items.append(AbstractItem(sl=t.sl, desc=desc, unit=t.unit, heading=True))
            parent = desc
            continue
        if t.sl:
            parent = None
        comp = classify(desc)
        if comp is None and parent:
            comp = classify(parent + " " + desc)
        if comp is None and t.amount:
            unknown.append(desc[:40])
        it = AbstractItem(sl=t.sl, desc=desc, unit=t.unit, qty=t.qty, rate=t.rate, stated_amount=t.amount)
        it.notes += t.notes
        if t.uncertain:
            it.notes.append("Check the wording (the two OCR readings differ): " + ", ".join(t.uncertain))
        _schedule_rows(it, t, comp)
        items.append(it)
    if unknown:
        notes.append("Items paid in one milestone (no standard payment schedule known for them): "
                     + "; ".join(unknown))
    return items, notes


def _layout_from_ts(ts: TSParseResult) -> Tuple[List[AbstractItem], List[str]]:
    """Abstract items (headings included) in the standard order. Returns (items, notes)."""
    notes: List[str] = []
    buckets: Dict[str, List[Tuple[TSItem, Component]]] = {}
    unknown: List[TSItem] = []
    parent: Optional[str] = None          # numbered heading such as "15 Valve chamber including cost of sluice"
    for t in ts.items:
        if t.heading:
            parent = t.desc if t.sl else None
            continue
        if t.sl:
            parent = None
        if not t.amount:
            notes.append(f"TS item with zero amount left out: {t.desc[:60]}")
            continue
        comp = classify(t.desc)
        if comp is None and parent and not t.sl:
            comp = classify(parent + " " + t.desc)   # unnumbered sub-item ("80 mm dia 6 Nos") of that heading
        if comp is None:
            unknown.append(t)
        else:
            buckets.setdefault(comp.key, []).append((t, comp))

    def make(t: TSItem, comp: Optional[Component]) -> AbstractItem:
        it = AbstractItem(sl="", desc=_clean_desc(t.desc), unit=t.unit, qty=t.qty, rate=t.rate,
                          stated_amount=t.amount)
        it.notes += t.notes
        sched = SCHEDULES[comp.schedule if comp else "single"]
        for remarks, pct in sched:
            it.rows.append(MilestoneRow(remarks=remarks, pct=pct, rate_paid=t.rate * pct,
                                        source="TS" + (f" Sl {t.sl}" if t.sl else "")))
        return it

    items: List[AbstractItem] = []
    sl = 0

    def add_component(key: str):
        nonlocal sl
        entries = buckets.pop(key, [])
        if not entries:
            return
        comp = entries[0][1]
        if comp.group:
            sl += 1
            items.append(AbstractItem(sl=str(sl), desc=GROUP_TITLES[comp.group], heading=True))
            for n, (t, c) in enumerate(entries):
                it = make(t, c)
                it.desc = f"{chr(ord('a') + n)}) {it.desc}"
                items.append(it)
        else:
            for t, c in entries:
                sl += 1
                it = make(t, c)
                it.sl = str(sl)
                items.append(it)

    for key in MAIN:
        add_component(key)
    items.append(AbstractItem(desc="OTHER WORKS:-", heading=True))
    for key in OTHER:
        add_component(key)
    # anything recognised but not in the standard lists, then unrecognised items, in TS order
    for key in list(buckets):
        add_component(key)
    for t in unknown:
        sl += 1
        it = make(t, None)
        it.sl = str(sl)
        items.append(it)
    if unknown:
        notes.append("Items not in the standard list (added under OTHER WORKS in TS order): "
                     + "; ".join(t.desc[:40] for t in unknown))
    return items, notes


# ---------------------------------------------------------------------------
# E-bill placement
# ---------------------------------------------------------------------------

def _find_item(eb: EBillItem, items: List[AbstractItem]) -> Tuple[Optional[AbstractItem], str]:
    priced = [it for it in items if not it.heading and it.rows]
    exact = [it for it in priced if money_eq(it.rate, eb.rate)]
    how = "same rate"
    if not exact:
        exact = [it for it in priced if abs(it.rate - eb.rate) <= max(0.51, 0.0005 * eb.rate)]
        how = "rate within rounding"
    if not exact:
        return None, "no TS item has this rate"
    if len(exact) == 1:
        return exact[0], how
    # several TS items share the rate: unit, quantity, then wording decide
    from item_mapper import _tokens
    ebt = _tokens(eb.desc)

    def score(it: AbstractItem):
        return (_norm_unit(it.unit) == _norm_unit(eb.unit), money_eq(it.qty, eb.qty, 1e-6),
                len(ebt & _tokens(it.desc)))
    ranked = sorted(exact, key=score, reverse=True)
    if score(ranked[0]) == score(ranked[1]):
        return None, f"{len(exact)} TS items have rate {eb.rate:,.2f} and the e-bill line does not tell them apart"
    return ranked[0], how + ", " + "unit/quantity/wording"


def _milestones_for(eb: EBillItem, it: AbstractItem, taken: set) -> Optional[List[MilestoneRow]]:
    free = [r for r in it.rows if id(r) not in taken]
    if len(it.rows) == 1:
        return free or None
    if eb.qty <= 0 or it.rate <= 0:
        return None
    pct = eb.upto_amt / (eb.qty * it.rate)
    for size in range(1, len(free) + 1):
        for combo in combinations(free, size):
            if abs(sum(r.pct for r in combo) - pct) < 1e-6:
                return list(combo)
    return None


def _fill(it: AbstractItem, r: MilestoneRow, qty: float, prev: float, since: float, upto: float, eb: EBillItem):
    g = r.rate_paid
    r.ebill_ref = eb.sl
    r.stated_upto = upto
    r.prev_amt, r.this_amt = prev, since
    if since < -0.005:
        # recovery of an earlier over-payment: previous as paid, this bill closes the line at the e-bill's upto
        h = _clean_qty(prev, g)
        r.prev_qty = h if h is not None else qty
        r.this_qty = 0.0
        r.l_fixed = True
        if len(it.rows) == 1 and money_eq(upto, it.qty * it.rate):
            r.l_formula = "f_minus_k"
        elif money_eq(upto, it.qty * g):
            r.l_formula = "dg_minus_k"
        r.notes.append(f"E-bill line {eb.sl}: recovery {format_inr(since)} (previously paid {format_inr(prev)}, "
                       f"now {format_inr(upto)} up to date)")
        return
    h, i = _clean_qty(prev, g), _clean_qty(since, g)
    r.prev_qty = h if h is not None else (qty if prev > 0 else 0.0)
    r.this_qty = i if i is not None else 0.0
    if h is None or i is None:
        r.notes.append(f"E-bill line {eb.sl}: amount is not quantity x rate - carried as an amount")
    r.notes.append(f"E-bill line {eb.sl}: previous {format_inr(prev)}, this bill {format_inr(since)}")


LAYOUTS = {
    "ts": "As in the TS (exact copy, TS order and Sl numbers)",
    "standard": "Standard TPIA order (items regrouped, serial numbers)",
}


def abstract_from_ts_and_ebill(ts: TSParseResult, ebill: Optional[EBillParseResult],
                               layout: str = "ts") -> Abstract:
    """
    ebill=None: the abstract is laid out from the TS alone, with every quantity / amount at zero.
    layout: "ts" copies the TS abstract row for row; "standard" regroups it in the standard TPIA order.
    """
    has_ebill = ebill is not None
    if ebill is None:
        from ebill_parser import EBillMetadata
        ebill = EBillParseResult(metadata=EBillMetadata(), items=[])
    items, notes = (_layout_copy_ts if layout == "ts" else _layout_from_ts)(ts)
    taken: set = set()
    unmatched: List[Tuple[EBillItem, str]] = []
    for eb in ebill.items:
        if money_eq(eb.upto_amt, 0) and money_eq(eb.since_amt, 0):
            continue
        it, how = _find_item(eb, items)
        if it is None:
            unmatched.append((eb, how))
            continue
        rows = _milestones_for(eb, it, taken)
        if not rows:
            pct = eb.upto_amt / (eb.qty * it.rate) if eb.qty and it.rate else 0
            unmatched.append((eb, f"matched Sl {it.sl or it.desc[:20]} by {how}, but its up-to-date amount is "
                                  f"{pct:.2%} of qty x rate - no free milestone (or combination) has that share"))
            continue
        for r in rows:
            taken.add(id(r))
        if len(rows) == 1:
            _fill(it, rows[0], eb.qty, eb.prev_amt, eb.since_amt, eb.upto_amt, eb)
        else:
            # one e-bill line covering several milestones (e.g. APDCL claimed 100% = 80% + 20%)
            total_pct = sum(r.pct for r in rows)
            for r in rows:
                share = r.pct / total_pct
                upto = eb.qty * r.rate_paid
                if money_eq(eb.prev_amt, 0):
                    prev, since = 0.0, upto
                elif money_eq(eb.since_amt, 0):
                    prev, since = upto, 0.0
                else:
                    prev, since = eb.prev_amt * share, eb.since_amt * share
                    r.notes.append("Previous / this-bill split between milestones in proportion to their %")
                _fill(it, r, eb.qty, prev, since, upto, eb)

    if unmatched:
        items.append(AbstractItem(desc="E-BILL LINES NOT FOUND IN THE TS:-", heading=True))
        for eb, why in unmatched:
            pct = 1.0
            it = AbstractItem(sl="", desc=_clean_desc(eb.desc)[:160], unit=eb.unit, qty=eb.qty, rate=eb.rate)
            r = MilestoneRow(remarks=f"E-bill line {eb.sl}", pct=pct, rate_paid=eb.rate, status="review",
                             unconfirmed=["item"], notes=[why], source=f"e-bill line {eb.sl}")
            _fill(it, r, eb.qty, eb.prev_amt, eb.since_amt, eb.upto_amt, eb)
            r.status = "review"
            it.rows.append(r)
            items.append(it)

    m = ebill.metadata
    meta = AbstractMeta(scheme=m.scheme or ts.metadata.scheme, contractor=m.contractor,
                        work_order_no=m.work_order_no, division=ts.metadata.division,
                        sub_division=ts.metadata.sub_division or ts.metadata.division,
                        bill_ref=m.bill_ref, bill_date=m.bill_date)
    sub_total = ts.metadata.sub_total if ts.metadata.sub_total else ts.total_sanctioned_amount
    if sub_total:
        meta.contract_value_text = f"Rs. {format_inr(sub_total, include_symbol=False)} (TS Sub Total)"
    stated: Dict[str, float] = {}
    if ts.metadata.sub_total:
        stated["contract"] = ts.metadata.sub_total
    if m.deduct_previous is not None:
        stated["prev"] = m.deduct_previous
    if m.net_this_bill is not None:
        stated["this"] = m.net_this_bill
    if m.total_upto_date is not None:
        stated["upto"] = m.total_upto_date

    abstract = Abstract(meta=meta, items=items, source_kind="ts_ebill" if has_ebill else "ts_only",
                        work_order_amount=sub_total or None, stated=stated, notes=notes)
    abstract.source_checks = list(ts.checks) + list(ebill.checks)
    if not has_ebill:
        abstract.source_checks.append(Check(
            "Contractor e-bill", False,
            "No e-bill uploaded: the abstract is laid out from the TS with all quantities at zero. Enter "
            "'Prev qty' / 'This qty' in the table, or upload the e-bill to fill them automatically.", "info"))
        return abstract
    abstract.source_checks.append(Check(
        "Every e-bill line placed on a TS item / milestone", not unmatched,
        f"{len(ebill.items) - len(unmatched)}/{len(ebill.items)} lines placed"
        + ("" if not unmatched else " - not placed: " + "; ".join(f"line {e.sl} ({why})" for e, why in unmatched[:4]))))
    return abstract
