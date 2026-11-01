"""
abstract_model.py - The single data model behind the TPIA Abstract of Bill.

Every input (scanned TPIA abstract, digital contractor e-bill, previous-abstract Excel template)
is converted into an `Abstract`. The UI edits it, the checks are computed from it, and the Excel
writer builds the workbook from it - so the numbers on screen, the checks and the Excel file can
never drift apart.

Layout follows the PHED Assam TPIA format:
    A Sl | B Description | C Unit | D Qty | E Rate | F Amount (=D*E)
    G Rate paid (=E*milestone%) | H Qty upto previous | I Qty this bill | J Qty upto date
    K Amount previous | L Amount this bill | M Amount upto date | N Remarks | O Milestone factor
"""

import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from utils import format_inr, money_eq

TOL = 0.005          # one-half paisa


@dataclass
class MilestoneRow:
    """One payment line (Excel row) of an item: a milestone stage, or the whole item if single-stage."""
    remarks: str = ""
    pct: Optional[float] = None          # milestone factor (col O); None = unknown
    rate_paid: float = 0.0               # G
    prev_qty: float = 0.0                # H
    this_qty: float = 0.0                # I
    prev_amt: float = 0.0                # K
    this_amt: float = 0.0                # L
    stated_upto: Optional[float] = None  # M as printed in the source document (for verification)
    status: str = "ok"                   # ok | corrected | review | edited
    notes: List[str] = field(default_factory=list)
    unconfirmed: List[str] = field(default_factory=list)   # fields the document's arithmetic cannot confirm
    source: str = ""                     # where the values came from, e.g. "page 2, row 7"
    excel_row: Optional[int] = None      # template mode: worksheet row this line lives in
    l_fixed: bool = False                # L does not use I (reconciliation / recovery row)
    l_formula: str = ""                  # generated layout: "" (I x G or value) | "f_minus_k" | "dg_minus_k"
    ebill_ref: Optional[int] = None      # e-bill line number this row was billed against

    @property
    def upto_qty(self) -> float:
        return self.prev_qty + self.this_qty

    @property
    def upto_amt(self) -> float:
        return self.prev_amt + self.this_amt

    def this_from_qty(self) -> bool:
        """True when L equals I x G to the paisa (so Excel can carry it as a live formula)."""
        return money_eq(self.this_qty * self.rate_paid, self.this_amt)

    def prev_from_qty(self) -> bool:
        return money_eq(self.prev_qty * self.rate_paid, self.prev_amt)


@dataclass
class AbstractItem:
    sl: str = ""
    desc: str = ""
    unit: str = ""
    qty: float = 0.0                     # D (contract / sanctioned quantity)
    rate: float = 0.0                    # E
    stated_amount: Optional[float] = None    # F as printed in the source (verification)
    rows: List[MilestoneRow] = field(default_factory=list)
    heading: bool = False                # text-only line such as "OTHER WORKS:-"
    notes: List[str] = field(default_factory=list)
    excel_row: Optional[int] = None      # template mode: row holding A..F

    @property
    def amount(self) -> float:
        return self.qty * self.rate


@dataclass
class AbstractMeta:
    scheme: str = ""
    contractor: str = ""
    work_order_no: str = ""
    contract_value_text: str = ""
    division: str = ""
    sub_division: str = ""
    status_text: str = "WORK IN PROGRESS"
    bill_ref: str = ""
    bill_date: str = ""


@dataclass
class Check:
    name: str
    passed: bool
    detail: str
    level: str = "critical"              # critical | warning | info

    @property
    def icon(self) -> str:
        if self.passed:
            return "✅"
        return {"critical": "❌", "warning": "⚠️"}.get(self.level, "ℹ️")


@dataclass
class Abstract:
    meta: AbstractMeta
    items: List[AbstractItem]
    source_kind: str                     # scanned_abstract | digital_ebill | ebill_template
    work_order_amount: Optional[float] = None
    # Figures printed in the source documents, used only for verification:
    #   contract / prev / this / upto  -> totals row of a scanned abstract
    #   claim_net / claim_upto / claim_prev -> contractor's e-bill summary
    stated: Dict[str, float] = field(default_factory=dict)
    source_checks: List[Check] = field(default_factory=list)   # produced by the parser
    notes: List[str] = field(default_factory=list)
    template_bytes: Optional[bytes] = None
    template_sheet: Optional[str] = None

    # ------------------------------------------------------------------ totals
    def all_rows(self):
        for it in self.items:
            for r in it.rows:
                yield it, r

    @property
    def total_contract(self) -> float:
        return sum(it.amount for it in self.items if not it.heading)

    @property
    def total_prev(self) -> float:
        return sum(r.prev_amt for _, r in self.all_rows())

    @property
    def total_this(self) -> float:
        return sum(r.this_amt for _, r in self.all_rows())

    @property
    def total_upto(self) -> float:
        return self.total_prev + self.total_this

    @property
    def recommended(self) -> int:
        """Payment recommended = ROUNDUP(this-bill total, 0), exactly as the official format."""
        v = round(self.total_this, 2)
        return int(math.ceil(v - 1e-9)) if v >= 0 else -int(math.ceil(-v - 1e-9))

    @property
    def progress(self) -> Optional[float]:
        if self.work_order_amount:
            return self.total_upto / self.work_order_amount
        return None

    def review_rows(self) -> List[MilestoneRow]:
        return [r for _, r in self.all_rows() if r.status == "review"]

    # ------------------------------------------------------------------ checks
    def run_checks(self) -> List[Check]:
        """Data checks. Recomputed after every edit, so they always describe the current numbers."""
        checks: List[Check] = list(self.source_checks)
        items = [it for it in self.items if not it.heading]
        checks.append(Check("Line items found", len(items) > 0, f"{len(items)} items / {sum(len(i.rows) for i in items)} payment lines"))

        # Row arithmetic against the printed Upto-Date amount (scanned abstracts)
        stated_rows = [(it, r) for it, r in self.all_rows() if r.stated_upto is not None and r.status != "edited"]
        if stated_rows:
            bad = [(it, r) for it, r in stated_rows if not money_eq(r.upto_amt, r.stated_upto, 0.011)]
            detail = (f"{len(stated_rows) - len(bad)}/{len(stated_rows)} lines satisfy Previous + This Bill = Upto Date"
                      + ("" if not bad else " - check: " + ", ".join(f"Sl {it.sl or '?'} ({r.remarks[:25]})" for it, r in bad[:6])))
            checks.append(Check("Line arithmetic (Prev + This = Upto)", not bad, detail))

        unresolved = self.review_rows()
        checks.append(Check(
            "All values confirmed",
            not unresolved,
            "Every value is confirmed by the document's own arithmetic" if not unresolved else
            f"{len(unresolved)} line(s) need review - values highlighted in the table could not be confirmed",
        ))

        edited = [r for _, r in self.all_rows() if r.status == "edited"]
        source_level = "warning" if edited else "critical"   # deliberate engineer changes are not reading errors

        def total_check(key: str, label: str, value: float, level: str = source_level):
            if key in self.stated:
                ok = money_eq(value, self.stated[key], 0.011)
                diff = value - self.stated[key]
                checks.append(Check(label, ok,
                                    f"Computed {format_inr(value)} vs printed {format_inr(self.stated[key])}"
                                    + ("" if ok else f" (difference {format_inr(diff)})")
                                    + ("" if ok or not edited else f" - {len(edited)} line(s) edited by the engineer"),
                                    level))

        total_check("contract", "Contract amount total matches document", self.total_contract)
        total_check("prev", "Previous-bill total matches document", self.total_prev)
        total_check("this", "This-bill total matches document", self.total_this)
        total_check("upto", "Upto-date total matches document", self.total_upto)
        if "recommended" in self.stated:
            ok = self.recommended == int(round(self.stated["recommended"]))
            checks.append(Check("Payment recommended matches document", ok,
                                f"ROUNDUP(this bill) = {format_inr(self.recommended)} vs printed {format_inr(self.stated['recommended'])}",
                                source_level))

        # Contractor's claim (digital e-bill) - the TPIA may legitimately recommend a different amount
        if "claim_net" in self.stated:
            ok = money_eq(self.total_this, self.stated["claim_net"], 0.011)
            checks.append(Check(
                "This-bill amount equals contractor's claim", ok,
                f"Recommended {format_inr(self.total_this)} vs claimed {format_inr(self.stated['claim_net'])}"
                + ("" if ok else f" (difference {format_inr(self.total_this - self.stated['claim_net'])})"),
                "warning"))
        if "claim_upto" in self.stated and self.source_kind != "digital_ebill":
            ok = money_eq(self.total_upto, self.stated["claim_upto"], 0.011)
            checks.append(Check("Upto-date amount equals contractor's claim", ok,
                                f"{format_inr(self.total_upto)} vs claimed {format_inr(self.stated['claim_upto'])}", "warning"))

        # Business sanity checks
        over = [(it, r) for it, r in self.all_rows()
                if it.qty > 0 and r.upto_qty > it.qty + 1e-6 and not r.l_fixed]
        checks.append(Check("Executed quantity within contract quantity", not over,
                            "All lines within contract quantity" if not over else
                            "; ".join(f"Sl {it.sl or '?'}: {r.upto_qty:g} > {it.qty:g} {it.unit}" for it, r in over[:5]),
                            "warning"))
        # printed amounts are sometimes displayed to one decimal (27818.67 shown as 27818.7)
        f_bad = [it for it in items if it.stated_amount is not None and not money_eq(it.amount, it.stated_amount, 0.051)]
        if any(it.stated_amount is not None for it in items):
            checks.append(Check("Contract amount = Qty x Rate", not f_bad,
                                "All items consistent" if not f_bad else
                                "; ".join(f"Sl {it.sl or '?'}: {it.qty:g} x {it.rate:g} = {format_inr(it.amount)} vs printed {format_inr(it.stated_amount)}"
                                          for it in f_bad[:5]), "warning"))
        if self.work_order_amount is None:
            checks.append(Check("Work order amount available", False,
                                "Enter the work order amount to compute physical / financial progress", "warning"))
        return checks

    def blocking_failures(self) -> List[Check]:
        return [c for c in self.run_checks() if not c.passed and c.level == "critical"]

    # ------------------------------------------------------------------ editing
    def row_by_id(self, row_id: str) -> Tuple[AbstractItem, MilestoneRow]:
        ii, ri = (int(x) for x in row_id.split("."))
        it = self.items[ii]
        return it, it.rows[ri]

    def row_ids(self):
        for ii, it in enumerate(self.items):
            for ri, r in enumerate(it.rows):
                yield f"{ii}.{ri}", it, r

    def apply_text_edit(self, row_id: str, field_name: str, text: str) -> bool:
        """The engineer corrects the wording of an item (Description / Unit) to match the document."""
        it, _ = self.row_by_id(row_id)
        attr = {"desc": "desc", "unit": "unit"}[field_name]
        text = " ".join(str(text).split())
        if getattr(it, attr) == text:
            return False
        it.notes.append(f"{'Description' if attr == 'desc' else 'Unit'} corrected by engineer "
                        f"(was: {getattr(it, attr)!r})")
        setattr(it, attr, text)
        return True

    def apply_edit(self, row_id: str, field_name: str, value: float) -> bool:
        """
        Applies one engineer edit. Quantities keep their amount in step when the amount was
        quantity x rate before the edit; amounts edited directly become explicit values.
        Returns True if something changed.
        """
        it, r = self.row_by_id(row_id)
        value = float(value)
        before = (it.qty, it.rate, r.rate_paid, r.prev_qty, r.this_qty, r.prev_amt, r.this_amt)
        prev_linked, this_linked = r.prev_from_qty(), r.this_from_qty()
        if field_name == "qty":
            it.qty = value
        elif field_name == "rate":
            it.rate = value
        elif field_name == "rate_paid":
            r.rate_paid = value
            if prev_linked:
                r.prev_amt = r.prev_qty * value
            if this_linked:
                r.this_amt = r.this_qty * value
        elif field_name == "prev_qty":
            r.prev_qty = value
            if prev_linked:
                r.prev_amt = value * r.rate_paid
        elif field_name == "this_qty":
            r.this_qty = value
            if this_linked:
                r.this_amt = value * r.rate_paid
        elif field_name == "prev_amt":
            r.prev_amt = value
        elif field_name == "this_amt":
            r.this_amt = value
        else:
            raise ValueError(field_name)
        changed = before != (it.qty, it.rate, r.rate_paid, r.prev_qty, r.this_qty, r.prev_amt, r.this_amt)
        if changed and r.l_formula and field_name in ("this_qty", "this_amt", "prev_amt", "qty", "rate", "rate_paid"):
            r.l_formula = ""        # the recovery formula no longer describes the edited line
            r.l_fixed = field_name != "this_qty"
        if changed:
            r.status = "edited"
            r.unconfirmed = []
            r.notes.append(f"Edited by engineer: {field_name} = {value:g}")
        return changed
