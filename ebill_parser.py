"""
ebill_parser.py - Parser for digital PHED Assam contractor e-bills (Running Account Bill-C, jjmassam.in).

The portal PDF has a real text layer, so values are read exactly (no OCR). The bill is then checked
against its own summary:
    sum of "Since previous bill" amounts  = "Net value of work supplies since previous bill"
    sum of "Up to date" amounts           = "Total"
    Total - "Deduct value ... previous bill" = Net value

Scanned bills / abstracts are handled by scanned_abstract.py (see pipeline.py for routing).
"""

import io
import logging
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Union

logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

import pdfplumber

from abstract_model import Check
from utils import clean_number, format_inr, money_eq

logger = logging.getLogger(__name__)

PdfSource = Union[str, Path, io.BytesIO, bytes]


@dataclass
class EBillItem:
    """One line of the contractor's e-bill."""
    sl: int
    unit: str
    qty: float            # quantity executed / supplied up to date (as per measurement book)
    desc: str
    rate: float
    upto_amt: float       # amount up to date
    since_amt: float      # amount since previous bill (this bill); negative = recovery

    @property
    def prev_amt(self) -> float:
        return self.upto_amt - self.since_amt

    @property
    def is_negative(self) -> bool:
        return self.since_amt < 0

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class EBillMetadata:
    contractor: str = ""
    scheme: str = ""
    work_order_no: str = ""
    bill_serial: str = ""
    bill_ref: str = ""             # e.g. "2nd RA Bill"
    bill_date: str = ""
    agreement_no: str = ""
    total_upto_date: Optional[float] = None      # "Total ₹ x/-" (sum of up-to-date amounts)
    total_rounded: Optional[float] = None        # "'Total value of work' done ... (A) ₹ x/-"
    deduct_previous: Optional[float] = None
    net_this_bill: Optional[float] = None
    figure_in_words: str = ""

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class EBillParseResult:
    metadata: EBillMetadata
    items: List[EBillItem]
    checks: List[Check] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def sum_since_amt(self) -> float:
        return sum(it.since_amt for it in self.items)

    @property
    def sum_upto_amt(self) -> float:
        return sum(it.upto_amt for it in self.items)

    @property
    def is_valid(self) -> bool:
        return all(c.passed for c in self.checks if c.level == "critical")


def _first(pattern: str, text: str, flags: int = re.I) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def parse_ebill_pdf(pdf_source: PdfSource) -> EBillParseResult:
    """Parses a digital Running Account Bill-C PDF."""
    src = io.BytesIO(pdf_source) if isinstance(pdf_source, (bytes, bytearray)) else pdf_source
    meta = EBillMetadata()
    items: List[EBillItem] = []
    warnings: List[str] = []
    texts: List[str] = []

    with pdfplumber.open(src) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            texts.append(page.extract_text() or "")
            for tbl in page.extract_tables() or []:
                for row in tbl:
                    if not row or not any(row):
                        continue
                    cells = [c.replace("\n", " ").strip() if c else "" for c in row]
                    if cells[0].isdigit():
                        k = 0
                    elif len(cells) > 1 and cells[1].isdigit() and not cells[0]:
                        k = 1
                    else:
                        continue
                    try:
                        nums = [clean_number(c) for c in cells[k + 4:] if c and re.search(r"\d", c)]
                        if len(nums) < 3:
                            raise ValueError(f"expected rate / up-to-date / since amounts, found {len(nums)} numbers")
                        items.append(EBillItem(
                            sl=int(cells[k]), unit=cells[k + 1], qty=clean_number(cells[k + 2]),
                            desc=cells[k + 3], rate=nums[0], upto_amt=nums[1], since_amt=nums[2]))
                    except Exception as ex:
                        warnings.append(f"Page {page_idx + 1}: could not read line {cells[:3]}: {ex}")

    full = "\n".join(texts)
    meta.contractor = _first(r"Contractor\s+or\s+supplier\s*:\s*([^\n]+)", full) or \
        _first(r"Name\s+of\s+Contractor\s*:\s*([^\n]+)", full)
    meta.contractor = re.sub(r"\s*\(\d+\)\s*$", "", meta.contractor)
    meta.scheme = _first(r"Name\s+of\s+work\s*[-:]\s*([^\n]+)", full)
    meta.scheme = re.sub(r"\s*\(\d+\)\s*$", "", meta.scheme)
    serial = _first(r"Serial\s+No\.?\s+of\s+this\s+Bill\s*:\s*([^\n]+)", full)
    m = re.match(r"(\d+)\s*(.*)", serial)
    meta.bill_serial, meta.bill_ref = (m.group(1), m.group(2).strip()) if m else (serial, "")
    meta.bill_date = _first(r"Dated\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4})", full) or _first(r"Bill\s+Date\s*:\s*([\d/]+)", full)
    meta.work_order_no = _first(r"Work\s+Order\s+(?:Number|No\.?)\s*:\s*([^\n]+)", full)
    meta.agreement_no = _first(r"Reference\s+to\s+Agreement\s*:\s*([^\n]+)", full)

    def money(pattern: str) -> Optional[float]:
        v = _first(pattern, full)
        return clean_number(v) if v else None

    meta.total_upto_date = money(r"\n\s*Total\s*₹?\s*(-?[\d,.]+)\s*/-")
    meta.total_rounded = money(r"Total\s+value\s+of\s+work.*?₹?\s*(-?[\d,.]+)\s*/-")
    meta.deduct_previous = money(r"Deduct\s+value\s+of\s+work.*?₹?\s*(-?[\d,.]+)\s*/-")
    meta.net_this_bill = money(r"Net\s+value\s+of\s+work.*?₹?\s*(-?[\d,.]+)\s*/-")
    meta.figure_in_words = _first(r"Figure\s+in\s+words\s*:\s*([^\n]+)", full)

    result = EBillParseResult(metadata=meta, items=items, warnings=warnings)
    result.checks = _self_checks(result)
    return result


def _self_checks(r: EBillParseResult) -> List[Check]:
    m = r.metadata
    checks = [Check("E-bill lines read", len(r.items) > 0 and not r.warnings,
                    f"{len(r.items)} lines read from the PDF text layer (exact values, no OCR)"
                    + ("" if not r.warnings else f"; {len(r.warnings)} line(s) unreadable: " + "; ".join(r.warnings[:3])))]
    if m.net_this_bill is not None:
        ok = money_eq(r.sum_since_amt, m.net_this_bill, 0.011)
        checks.append(Check("E-bill lines add up to its 'Net value since previous bill'", ok,
                            f"Sum {format_inr(r.sum_since_amt)} vs printed {format_inr(m.net_this_bill)}"))
    else:
        checks.append(Check("E-bill net value found", False, "'Net value of work supplies since previous bill' not found"))
    if m.total_upto_date is not None:
        ok = money_eq(r.sum_upto_amt, m.total_upto_date, 0.011)
        checks.append(Check("E-bill up-to-date amounts add up to its Total", ok,
                            f"Sum {format_inr(r.sum_upto_amt)} vs printed {format_inr(m.total_upto_date)}"))
    if None not in (m.total_upto_date, m.deduct_previous, m.net_this_bill):
        ok = money_eq(m.total_upto_date - m.deduct_previous, m.net_this_bill, 0.011)
        checks.append(Check("E-bill summary: Total - Previous = Net", ok,
                            f"{format_inr(m.total_upto_date)} - {format_inr(m.deduct_previous)} = "
                            f"{format_inr(m.total_upto_date - m.deduct_previous)} vs {format_inr(m.net_this_bill)}"))
    return checks
