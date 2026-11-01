"""
pipeline.py - One entry point from uploaded documents to a verified Abstract.

    bill PDF (digital e-bill or scanned abstract) [+ TS PDF] [+ previous abstract Excel]
        -> Abstract (+ TS result + TS benchmark)

Routing:
  * scanned bill PDF                -> scanned_abstract.parse_scanned_abstract (grid OCR + arithmetic checks)
  * digital e-bill + template Excel -> item_mapper.abstract_from_template (exact identity matching)
  * digital e-bill alone            -> item_mapper.abstract_from_ebill (optionally completed from the TS)
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from abstract_model import Abstract, Check
from ebill_parser import EBillParseResult, parse_ebill_pdf
from item_mapper import abstract_from_ebill, abstract_from_template, ts_benchmark
from scanned_abstract import parse_scanned_abstract
from table_extractor import is_scanned_pdf
from ts_ebill_builder import abstract_from_ts_and_ebill
from ts_parser import TSParseResult, parse_ts_pdf


@dataclass
class PipelineResult:
    abstract: Abstract
    ebill: Optional[EBillParseResult] = None
    ts: Optional[TSParseResult] = None
    benchmark_rows: List[Dict] = field(default_factory=list)
    benchmark_checks: List[Check] = field(default_factory=list)
    bill_is_scanned: bool = False


def _looks_like_ts(pdf_bytes: bytes) -> bool:
    """A TS abstract ('Abstract of cost ... Technical Sanction / Revised estimate') rather than a bill."""
    import re
    from table_extractor import open_pdf, quick_page_text
    pdf = open_pdf(pdf_bytes)
    if len(pdf) > 120:
        return False
    texts = " ".join(quick_page_text(pdf_bytes)).lower()
    if re.search(r"rate\s*at\s*which|running\s*account|measurement\s*book|bill\s*no", texts):
        return False
    return bool(re.search(r"abstra\w*\s*(of|for)\b.{0,60}(estimate|technical|sanction)", texts))


def parse_ts(ts_bytes: Optional[bytes], pages: Optional[List[int]] = None) -> Optional[TSParseResult]:
    return parse_ts_pdf(ts_bytes, pages) if ts_bytes else None


def build_abstract(bill_bytes: Optional[bytes], ts: Optional[TSParseResult] = None,
                   template_bytes: Optional[bytes] = None,
                   ts_pages: Optional[List[int]] = None, layout: str = "ts") -> PipelineResult:
    """layout: "ts" = the TS abstract copied row for row; "standard" = regrouped in the standard TPIA order."""
    if bill_bytes and ts is None and _looks_like_ts(bill_bytes):
        # the TS was put in the e-bill box: use it as the TS (ts_pages = the page list typed for the TS)
        found = parse_ts_pdf(bill_bytes, ts_pages)
        if found.priced_items:
            result = build_abstract(None, found, layout=layout)
            result.abstract.notes.append("The file in the 'Contractor e-bill' box is a Technical Sanction - "
                                         "it was read as the TS.")
            return result
    if not bill_bytes:
        if ts is None or not ts.priced_items:
            raise ValueError("Upload the contractor bill or the Technical Sanction document.")
        # TS only: blank abstract in the official layout (quantities entered on screen)
        return PipelineResult(abstract=abstract_from_ts_and_ebill(ts, None, layout), ts=ts)
    scanned = is_scanned_pdf(bill_bytes)
    ebill = None
    if scanned:
        try:
            abstract = parse_scanned_abstract(bill_bytes)
        except ValueError as ex:
            if "Rate at which payment" in str(ex):
                raise ValueError("The file in the 'Contractor e-bill' box is not a contractor bill / TPIA abstract "
                                 f"({ex}). If it is the TS, put it only in the 'Technical Sanction' box and leave "
                                 "the e-bill box empty.") from ex
            raise
        if template_bytes:
            abstract.notes.append("A template was attached, but the scanned abstract already contains every "
                                  "line - the abstract is rebuilt from the scan.")
    else:
        ebill = parse_ebill_pdf(bill_bytes)
        if ts is not None and ts.priced_items:
            # main route: items and rates from the TS, previous / this-bill amounts from the e-bill
            abstract = abstract_from_ts_and_ebill(ts, ebill, layout)
            if template_bytes:
                abstract.notes.append("An Excel template was also attached; the workbook is generated from the TS "
                                      "and the e-bill in the official layout (the template is not needed).")
        elif template_bytes:
            abstract = abstract_from_template(ebill, template_bytes)
        else:
            abstract = abstract_from_ebill(ebill, ts)
            abstract.source_checks.append(Check(
                "Technical Sanction attached", False,
                "Upload the TS document: the full item list, rates and milestone rows come from its abstract page",
                "warning"))
    result = PipelineResult(abstract=abstract, ebill=ebill, ts=ts, bill_is_scanned=scanned)
    if ts is not None and abstract.source_kind != "ts_ebill":
        result.benchmark_rows, result.benchmark_checks = ts_benchmark(abstract, ts)
    return result
