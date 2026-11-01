"""
excel_writer.py - Writes the TPIA "Abstract of Bill (Payment Recommendation)" workbook.

Two strategies, both driven only by the reviewed Abstract model:
  * build_workbook()  - generates the official PHED Assam layout from scratch with live formulas
                        (F=D*E, G=E*%, J=H+I, K=G*H, L=I*G, M=K+L, SUM totals, ROUNDUP, progress).
                        Amounts that the source document carries as adjustments (not qty x rate)
                        are written as values and highlighted, so the totals stay exact.
  * fill_template()   - writes only column I (this-bill quantities), the words line and missing
                        header fields into the engineer's previous abstract; every template formula,
                        style and merged cell is left untouched.

verify_workbook() recalculates the produced file (formula_eval) and compares every line and total
with the model. The app only offers a download when this verification passes.
"""

import io
import re
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from abstract_model import Abstract, Check
from formula_eval import FormulaError, WorksheetEvaluator
from utils import amount_to_words, format_inr, money_eq

INR_FMT = r'[>=10000000]##\,##\,##\,##0.00;[>=100000]##\,##\,##0.00;##,##0.00'
QTY_FMT = "General"
PCT_FMT = "0%"

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_HEAD = PatternFill("solid", start_color="EAEAEA", end_color="EAEAEA")
FILL_TOTAL = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
FILL_ADJ = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")      # value carried, not qty x rate
FILL_REVIEW = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")   # needs engineer review


def _pct_text(p: float) -> str:
    return f"{p * 100:g}%"


def _words_line(abstract: Abstract) -> str:
    return f"In Words: {amount_to_words(abstract.recommended)}."


def _contract_value_text(abstract: Abstract) -> str:
    if abstract.meta.contract_value_text:
        return abstract.meta.contract_value_text
    if abstract.work_order_amount:
        return f"{abstract.work_order_amount / 100000:.4f} LAKH"
    return ""


# ---------------------------------------------------------------------------
# Strategy 1: official layout from scratch
# ---------------------------------------------------------------------------

def build_workbook(abstract: Abstract) -> Tuple[bytes, Dict[int, int]]:
    """Returns (xlsx bytes, {id(MilestoneRow): excel row})."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Payment Recommendation"
    bold = Font(name="Calibri", size=10, bold=True)
    regular = Font(name="Calibri", size=10)
    title = Font(name="Calibri", size=12, bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    widths = {"A": 7, "B": 42, "C": 8, "D": 10, "E": 13, "F": 15, "G": 14, "H": 11, "I": 11, "J": 11,
              "K": 15, "L": 15, "M": 15, "N": 34, "O": 7}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    for r, text in ((1, "(JAL JEEVAN MISSION)"), (2, "PUBLIC HEALTH DEPARTMENT, GOVT. OF ASSAM"),
                    (3, "ABSTRACT OF BILL(Payment Recommendation)")):
        ws.cell(r, 1, text).font = title
        ws.cell(r, 1).alignment = center
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)

    m = abstract.meta
    header = [(4, "Name of Scheme", m.scheme), (5, "Name of Contractor:-", m.contractor),
              (6, "Work Order No:-", m.work_order_no), (7, "Contract Value:-", _contract_value_text(abstract)),
              (8, "Name of Division:", m.division), (9, "Completion Date", m.status_text or "WORK IN PROGRESS")]
    for r, label, value in header:
        ws.cell(r, 1, label).font = bold
        ws.cell(r, 3, value).font = bold if r in (4, 9) else regular
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
    ws.cell(8, 10, "Name of Sub-Division:-").font = bold
    ws.cell(8, 13, m.sub_division or m.division)
    if m.bill_ref or m.bill_date:
        ws.cell(4, 10, "Bill:").font = bold
        ws.cell(4, 11, " ".join(x for x in (m.bill_ref, f"dated {m.bill_date}" if m.bill_date else "") if x))

    heads10 = {1: "SL. No", 2: "Description of item", 3: "Unit", 4: "As per Contract",
               7: "Rate at which payment has been made (Rs)", 8: "Quantity executed", 11: "Amount (Rs.)",
               14: "Remarks", 15: "%"}
    heads11 = {4: "Quantity", 5: "Rate (Rs.)", 6: "Amount (Rs.)", 8: "Upto Previous Bill", 9: "This Bill",
               10: "Upto Date", 11: "Previous Bill", 12: "This Bill", 13: "Upto Date"}
    for c in range(1, 16):
        for r in (10, 11):
            cell = ws.cell(r, c)
            cell.font, cell.alignment, cell.fill, cell.border = bold, center, FILL_HEAD, BORDER
    for c, t in heads10.items():
        ws.cell(10, c, t)
    for c, t in heads11.items():
        ws.cell(11, c, t)
    for rng in ("A10:A11", "B10:B11", "C10:C11", "D10:F10", "G10:G11", "H10:J10", "K10:M10", "N10:N11", "O10:O11"):
        ws.merge_cells(rng)
    ws.row_dimensions[10].height = 30
    ws.row_dimensions[11].height = 30

    row = 12
    rowmap: Dict[int, int] = {}
    for it in abstract.items:
        if it.heading:
            ws.cell(row, 1, it.sl or None)
            ws.cell(row, 2, it.desc).font = bold
            ws.cell(row, 3, it.unit or None)
            for c in range(1, 16):
                ws.cell(row, c).border = BORDER
                ws.cell(row, c).alignment = left if c == 2 else center
            row += 1
            continue
        first = row
        ws.cell(row, 1, it.sl or None)
        ws.cell(row, 2, it.desc)
        ws.cell(row, 3, it.unit)
        if (it.stated_amount and it.rate and abs(it.qty - round(it.qty, 3)) > 1e-9
                and money_eq(it.qty * it.rate, it.stated_amount)):
            # quantity taken as TS Amount / Rate (the TS's own Qty x Rate disagrees with its amount)
            c = ws.cell(row, 4, f"={it.stated_amount:.2f}/{it.rate:g}")
            c.number_format = "0.000"
            c.comment = Comment("Quantity = TS amount / rate. " + "; ".join(it.notes)[:800], "TPIA tool")
        else:
            ws.cell(row, 4, it.qty)
        ws.cell(row, 5, it.rate)
        ws.cell(row, 6, f"=D{row}*E{row}")
        for mr in it.rows:
            r = row
            rowmap[id(mr)] = r
            if mr.pct is not None and money_eq(it.rate * mr.pct, mr.rate_paid):
                ws.cell(r, 7, f"=E{first}*{_pct_text(mr.pct)}" if mr.pct != 1 else f"=E{first}")
            else:
                ws.cell(r, 7, mr.rate_paid)
            ws.cell(r, 8, mr.prev_qty)
            ws.cell(r, 9, mr.this_qty)
            ws.cell(r, 10, f"=H{r}+I{r}")
            recovery = mr.l_formula in ("f_minus_k", "dg_minus_k")
            if mr.prev_from_qty() and not recovery:
                ws.cell(r, 11, f"=G{r}*H{r}")
            else:
                c = ws.cell(r, 11, round(mr.prev_amt, 2))
                if not recovery:
                    c.fill = FILL_ADJ
                c.comment = Comment("Previous payment as per the e-bill (Up to date - Since previous bill).", "TPIA tool")
            if mr.l_formula == "f_minus_k":         # recovery: whole item now paid at its revised amount
                ws.cell(r, 12, f"=F{first}-K{r}")
            elif mr.l_formula == "dg_minus_k":      # recovery: this milestone now paid at qty x milestone rate
                ws.cell(r, 12, f"=D{first}*G{r}-K{r}")
            elif mr.this_from_qty():
                ws.cell(r, 12, f"=I{r}*G{r}")
            else:
                c = ws.cell(r, 12, round(mr.this_amt, 2))
                c.fill = FILL_ADJ
                c.comment = Comment("This-bill amount as per the source document (adjustment, not quantity x rate).",
                                    "TPIA tool")
            ws.cell(r, 13, f"=K{r}+L{r}")
            ws.cell(r, 14, mr.remarks)
            if mr.pct is not None:
                ws.cell(r, 15, mr.pct).number_format = PCT_FMT
            if mr.status == "review":
                for c in (7, 8, 9, 11, 12, 13):
                    ws.cell(r, c).fill = FILL_REVIEW
                ws.cell(r, 14).comment = Comment("NEEDS REVIEW: " + "; ".join(mr.notes)[:900], "TPIA tool")
            row += 1
        if not it.rows:
            row += 1
        last = row - 1
        for r in range(first, last + 1):
            for c in range(1, 16):
                cell = ws.cell(r, c)
                cell.border = BORDER
                if cell.font is None or not cell.font.b:
                    cell.font = regular
                cell.alignment = left if c in (2, 14) else (center if c in (1, 3, 15) else right)
            for c in (5, 6, 7, 11, 12, 13):
                ws.cell(r, c).number_format = INR_FMT
            for c in (4, 8, 9, 10):
                ws.cell(r, c).number_format = QTY_FMT
        if last > first:
            for c in range(1, 7):
                ws.merge_cells(start_row=first, start_column=c, end_row=last, end_column=c)
        # give a long (merged) description enough height: ~48 characters per line at width 42
        lines_needed = max(1, -(-len(it.desc) // 48))
        per_row = max(15.0, 13.0 * lines_needed / (last - first + 1))
        for r in range(first, last + 1):
            remark_lines = max(1, -(-len(str(ws.cell(r, 14).value or "")) // 38))
            ws.row_dimensions[r].height = max(per_row, 13.0 * remark_lines)

    last_item_row = row - 1
    tot = row
    ws.cell(tot, 2, "Total Amount(TS Amount without Benificary)")
    for c, col in ((6, "F"), (11, "K"), (12, "L"), (13, "M")):
        ws.cell(tot, c, f"=SUM({col}12:{col}{last_item_row})").number_format = INR_FMT
    for c in range(1, 16):
        ws.cell(tot, c).font, ws.cell(tot, c).fill, ws.cell(tot, c).border = bold, FILL_TOTAL, BORDER
    wo = tot + 1
    ws.cell(wo, 2, "Work order Amount").font = bold
    if abstract.work_order_amount:
        ws.cell(wo, 6, abstract.work_order_amount).number_format = INR_FMT
    ws.cell(wo, 6).font = bold
    rec = tot + 2
    ws.cell(rec, 5, "Payment recommendaded in this Bill = ").font = bold
    c = ws.cell(rec, 9, f"=ROUNDUP(ROUND(L{tot},2),0)")
    c.font, c.number_format = Font(name="Calibri", size=12, bold=True), INR_FMT
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=rec, start_column=9, end_row=rec, end_column=11)   # room for ₹ 1,00,00,000.00
    ws.cell(tot + 3, 6, _words_line(abstract)).font = bold
    for off, label in ((5, "Physical Progress with respect to Work order Amount"),
                       (6, "Financial Progress with respect to Work order Amount")):
        ws.cell(tot + off, 7, label)
        p = ws.cell(tot + off, 14, f'=IF(F{wo}>0,M{tot}/F{wo},"")')
        p.number_format = "0.00%"
    ws.cell(tot + 9, 9, "Seal and Signature of FE/PC/CME of TPIA").font = bold
    note_r = tot + 11
    ws.cell(note_r, 2, "Shaded cells: orange = amount carried from the source document (not quantity x rate); "
                       "yellow = value flagged for engineer review.").font = Font(name="Calibri", size=8, italic=True)

    ws.freeze_panes = "D12"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "10:11"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), rowmap


# ---------------------------------------------------------------------------
# Strategy 2: fill the engineer's template
# ---------------------------------------------------------------------------

def fill_template(abstract: Abstract) -> Tuple[bytes, Dict[int, int]]:
    from item_mapper import find_layout
    if not abstract.template_bytes:
        raise ValueError("No template attached to this abstract.")
    wb = openpyxl.load_workbook(io.BytesIO(abstract.template_bytes), data_only=False)
    ws = wb[abstract.template_sheet] if abstract.template_sheet in wb.sheetnames else wb.worksheets[0]
    lay = find_layout(ws)
    rowmap: Dict[int, int] = {}
    for _, mr in abstract.all_rows():
        if mr.excel_row is None:
            continue
        rowmap[id(mr)] = mr.excel_row
        if not mr.l_fixed:
            ws.cell(mr.excel_row, 9).value = mr.this_qty
    if lay.recommend_cell:
        r, c = lay.recommend_cell
        f = str(ws.cell(r, c).value)
        m = re.fullmatch(r"=ROUNDUP\((L\d+),\s*0\)", f.replace(" ", ""), re.I)
        if m:   # identical result, immune to binary floating-point noise in the SUM
            ws.cell(r, c).value = f"=ROUNDUP(ROUND({m.group(1)},2),0)"
    if lay.words_cell:
        ws.cell(*lay.words_cell).value = _words_line(abstract)
    if lay.work_order_cell and abstract.work_order_amount:
        cur = ws.cell(*lay.work_order_cell).value
        if not (isinstance(cur, (int, float)) and money_eq(cur, abstract.work_order_amount)):
            ws.cell(*lay.work_order_cell).value = abstract.work_order_amount
    m = abstract.meta
    for r, c, v in ((4, 3, m.scheme), (5, 3, m.contractor), (6, 3, m.work_order_no), (7, 3, m.contract_value_text),
                    (8, 3, m.division), (8, 13, m.sub_division), (9, 3, m.status_text)):
        cell = ws.cell(r, c)
        if v and type(cell).__name__ != "MergedCell" and str(cell.value or "").strip() != v.strip():
            cell.value = v
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), rowmap


def generate_workbook(abstract: Abstract) -> Tuple[bytes, Dict[int, int]]:
    return fill_template(abstract) if abstract.template_bytes else build_workbook(abstract)


# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------

def verify_workbook(xlsx: bytes, abstract: Abstract, rowmap: Dict[int, int]) -> List[Check]:
    """Recalculates the workbook and compares it with the model, line by line and in total."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx), data_only=False)
    ws = wb[abstract.template_sheet] if abstract.template_sheet in wb.sheetnames else wb.worksheets[0]
    ev = WorksheetEvaluator(ws)
    checks: List[Check] = []
    try:
        tot = next(r for r in range(12, ws.max_row + 1) if re.search(r"total\s*amount", str(ws.cell(r, 2).value or ""), re.I))
    except StopIteration:
        return [Check("Excel totals row present", False, "No 'Total Amount' row found in the generated workbook")]
    try:
        bad_lines = []
        for _, mr in abstract.all_rows():
            r = rowmap.get(id(mr))
            if r is None:
                continue
            k, l, mm = ev.number(r, 11), ev.number(r, 12), ev.number(r, 13)
            if not (money_eq(k, mr.prev_amt) and money_eq(l, mr.this_amt) and money_eq(mm, mr.upto_amt)):
                bad_lines.append(f"row {r}: K {k:,.2f}/{mr.prev_amt:,.2f}, L {l:,.2f}/{mr.this_amt:,.2f}")
        checks.append(Check("Excel lines recalculate to the reviewed values", not bad_lines,
                            f"{len(rowmap)} lines recalculated" if not bad_lines else "; ".join(bad_lines[:5])))
        for col, label, val in ((6, "Contract amount", abstract.total_contract), (11, "Previous bill", abstract.total_prev),
                                (12, "This bill", abstract.total_this), (13, "Upto date", abstract.total_upto)):
            got = ev.number(tot, col)
            checks.append(Check(f"Excel total - {label}", money_eq(got, val),
                                f"Excel {format_inr(got)} vs reviewed {format_inr(val)}"))
        rec_cell = next(((r, c) for r in range(tot + 1, tot + 6) for c in range(1, 16)
                         if "ROUNDUP" in str(ws.cell(r, c).value or "").upper()), None)
        if rec_cell:
            got = ev.number(*rec_cell)
            checks.append(Check("Excel payment recommended", money_eq(got, abstract.recommended),
                                f"Excel {format_inr(got)} vs reviewed {format_inr(abstract.recommended)}"))
        else:
            checks.append(Check("Excel payment recommended", False, "ROUNDUP cell not found"))
        words = next((str(ws.cell(r, c).value) for r in range(tot + 1, tot + 6) for c in range(1, 16)
                      if re.search(r"in\s*words", str(ws.cell(r, c).value or ""), re.I)), "")
        expected = amount_to_words(abstract.recommended)
        checks.append(Check("Excel amount in words", expected in words, words or "no 'In Words' line found"))
        if abstract.work_order_amount:
            prog = next(((r, 14) for r in range(tot + 1, tot + 8)
                         if "/" in str(ws.cell(r, 14).value or "") and str(ws.cell(r, 14).value).startswith("=")), None)
            if prog:
                got = ev.value(*prog)
                ok = isinstance(got, float) and abs(got - abstract.progress) < 1e-9
                checks.append(Check("Excel progress %", ok, f"Excel {got if not isinstance(got, float) else f'{got:.2%}'} "
                                                           f"vs reviewed {abstract.progress:.2%}"))
    except FormulaError as ex:
        checks.append(Check("Excel formulas evaluate", False, f"Formula error: {ex}"))
    return checks
