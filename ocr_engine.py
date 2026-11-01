"""
ocr_engine.py - Compatibility wrapper (older scripts import ocr_pdf from here).

All OCR now lives in table_extractor.py (RapidOCR with Windows OCR fallback, grid-aware reading).
This module only re-exposes a page/line view of that output.
"""

from dataclasses import dataclass, field
from typing import List

from table_extractor import OCRWord, PdfSource, detect_orientation, ocr_words, open_pdf, render_page_native


@dataclass
class OCRLine:
    text: str
    x: float
    y: float
    w: float
    h: float
    words: List[OCRWord] = field(default_factory=list)


@dataclass
class OCRPageResult:
    page_number: int
    rotation_applied: int
    char_count: int
    lines: List[OCRLine]
    full_text: str


def ocr_pdf(pdf_source: PdfSource, scale: float = 0.0) -> List[OCRPageResult]:
    """OCRs every page (upright orientation detected automatically). `scale` is ignored (native resolution)."""
    pdf = open_pdf(pdf_source)
    out: List[OCRPageResult] = []
    angle = None
    for i in range(len(pdf)):
        img = render_page_native(pdf[i])
        if angle is None:
            angle = detect_orientation(img)
        img = img.rotate(angle, expand=True) if angle else img
        words = sorted(ocr_words(img), key=lambda w: (round(w.cy / 20), w.x))
        lines = [OCRLine(w.text, w.x, w.y, w.w, w.h, [w]) for w in words]
        text = "\n".join(w.text for w in words)
        out.append(OCRPageResult(i + 1, angle, len(text), lines, text))
    return out
