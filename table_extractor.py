"""
table_extractor.py - Grid-based table reader for scanned (camera / CamScanner / OKEN) PDFs.

Pipeline per page:
  1. Render the page at the pixel density of its embedded scan (no detail thrown away).
  2. Pick the reading orientation (0/90/180/270) with a fast low-resolution OCR probe.
  3. Remove skew (shear search on the ruling lines).
  4. Detect the ruled grid: vertical rulings -> columns; horizontal rulings are found inside each
     column (immune to page curl) and tracked column-to-column into global table rows.
  5. OCR the page once (RapidOCR / PP-OCR, offline; Windows OCR as fallback) and drop every word
     into the exact (row, column) it sits in. Merged cells keep their span.

Every word keeps its recognition confidence so the parsers can treat low-confidence numbers as
"needs verification" instead of silently trusting them.
"""

import asyncio
import io
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import numpy as np
import pypdfium2
from PIL import Image, ImageFilter, ImageOps

logger = logging.getLogger(__name__)

PdfSource = Union[str, Path, bytes, io.BytesIO]


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class OCRWord:
    text: str
    x: float
    y: float
    w: float
    h: float
    score: float = 1.0
    alts: List[str] = field(default_factory=list)    # alternative readings (second opinions)

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2


@dataclass
class Cell:
    """A ruled cell of one column. Spans master rows row_start..row_end (merged cells span several)."""
    col: int
    row_start: int
    row_end: int


@dataclass
class ScannedTable:
    page_number: int
    width: int
    height: int
    rotation: int
    columns: List[Tuple[int, int]]                 # (x0, x1) per column, left to right
    n_rows: int                                    # master rows (between consecutive tracked rulings)
    cells: List[List[Cell]]                        # cells[col] top-to-bottom
    grid_words: Dict[Tuple[int, int], List[OCRWord]]   # (row, col) -> words in that exact row/column
    page_words: List[OCRWord]                      # every OCR word on the page (titles, footers...)
    top: float = 0.0                               # y of the first ruling (left edge)
    bottom: float = 0.0                            # y of the last ruling
    # the same cells read by a second, independent engine (Windows OCR): its mistakes are different,
    # so text the two engines disagree on can be decided word by word
    alt_grid_words: Dict[Tuple[int, int], List[OCRWord]] = field(default_factory=dict)

    @property
    def n_cols(self) -> int:
        return len(self.columns)

    def words(self, row: int, col: int) -> List[OCRWord]:
        return self.grid_words.get((row, col), [])

    def alt_words(self, row: int, col: int) -> List[OCRWord]:
        return self.alt_grid_words.get((row, col), [])

    def text(self, row: int, col: int) -> str:
        return _join_words(self.words(row, col))

    def span(self, row: int, col: int) -> Tuple[int, int]:
        for c in self.cells[col]:
            if c.row_start <= row <= c.row_end:
                return c.row_start, c.row_end
        return row, row

    def span_words(self, row: int, col: int) -> List[OCRWord]:
        a, b = self.span(row, col)
        out: List[OCRWord] = []
        for r in range(a, b + 1):
            out.extend(self.words(r, col))
        return out

    def span_text(self, row: int, col: int) -> str:
        return _join_words(self.span_words(row, col))

    def row_text(self, row: int) -> str:
        return " | ".join(self.text(row, c) for c in range(self.n_cols))


def _join_words(words: List[OCRWord]) -> str:
    if not words:
        return ""
    ws = sorted(words, key=lambda w: w.cy)
    lines: List[List[OCRWord]] = []
    for w in ws:
        if lines and abs(w.cy - float(np.mean([v.cy for v in lines[-1]]))) < max(w.h, 10) * 0.6:
            lines[-1].append(w)
        else:
            lines.append([w])
    return " ".join(" ".join(v.text for v in sorted(ln, key=lambda v: v.x)) for ln in lines).strip()


# ---------------------------------------------------------------------------
# OCR backends
# ---------------------------------------------------------------------------

_rapid_engine = None
_rapid_failed = False


def _get_rapid():
    """
    Lazily creates the RapidOCR engine and runs one detection + one recognition pass.
    The warm-up must happen before Windows OCR (WinRT) is used in the process: running onnxruntime
    for the first time after WinRT has initialised COM crashes the interpreter (access violation).
    """
    global _rapid_engine, _rapid_failed
    if _rapid_engine is None and not _rapid_failed:
        try:
            logging.getLogger("RapidOCR").setLevel(logging.ERROR)
            from rapidocr import RapidOCR
            eng = RapidOCR(params={"Global.log_level": "error"})
            warm = np.full((64, 256, 3), 255, dtype=np.uint8)
            warm[28:36, 20:236] = 0
            eng(warm, use_det=True, use_cls=False, use_rec=True)
            eng(warm, use_det=False, use_cls=False, use_rec=True)
            _rapid_engine = eng
        except Exception as ex:  # pragma: no cover - depends on installation
            logger.warning("RapidOCR unavailable (%s); falling back to Windows OCR.", ex)
            _rapid_failed = True
    return _rapid_engine


def ocr_backend_name() -> str:
    return "RapidOCR (PP-OCR, offline)" if _get_rapid() is not None else "Windows OCR"


def _ocr_windows(img: Image.Image) -> List[OCRWord]:
    _get_rapid()   # see _get_rapid(): onnxruntime must be initialised before WinRT
    try:
        import winocr
    except ImportError:
        return []          # winocr is Windows-only; on Linux just skip
    res = asyncio.run(winocr.recognize_pil(img.convert("RGB"), "en"))
    out: List[OCRWord] = []
    for line in res.lines:
        for w in line.words:
            r = w.bounding_rect
            out.append(OCRWord(w.text, r.x, r.y, r.width, r.height, score=0.85))
    return out


def ocr_words(img: Image.Image) -> List[OCRWord]:
    """Text boxes for a whole (upright) page image."""
    eng = _get_rapid()
    if eng is None:
        words = _ocr_windows(img)
        if not words:
            logger.warning("Neither RapidOCR nor Windows OCR available — no text will be read")
        return words
    # flags are passed explicitly every time: RapidOCR remembers the last call's flags
    res = eng(np.asarray(img.convert("RGB")), use_det=True, use_cls=False, use_rec=True)
    out: List[OCRWord] = []
    if res is None or res.boxes is None:
        return out
    for box, txt, sc in zip(res.boxes, res.txts, res.scores):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
        out.append(OCRWord(str(txt).strip(), float(x0), float(y0), float(x1 - x0), float(y1 - y0), float(sc)))
    return [w for w in out if w.text]


def read_crop(img: Image.Image, box: Tuple[float, float, float, float],
              scales: Tuple[float, ...] = (0.75, 1.4)) -> List[Tuple[str, float]]:
    """Independent readings of one text box: recognition-only passes at several scales + Windows OCR."""
    x0, y0, x1, y1 = box
    pad = 6
    crop = img.crop((max(0, int(x0) - pad), max(0, int(y0) - pad), int(x1) + pad, int(y1) + pad)).convert("RGB")
    out: List[Tuple[str, float]] = []
    eng = _get_rapid()
    if eng is not None:
        for scale in scales:
            c = crop if scale == 1.0 else crop.resize((max(8, int(crop.width * scale)), max(8, int(crop.height * scale))), Image.LANCZOS)
            try:
                r = eng(np.asarray(c), use_det=False, use_cls=False, use_rec=True)
                if r is not None and getattr(r, "txts", None):
                    out.append((str(r.txts[0]).strip(), float(r.scores[0])))
            except Exception:
                pass
    try:
        canvas = Image.new("RGB", (crop.width + 60, crop.height + 60), (255, 255, 255))
        canvas.paste(crop, (30, 30))
        out.append((" ".join(w.text for w in _ocr_windows(canvas)), 0.85))
    except Exception:
        pass
    return [(t, s) for t, s in out if t]


def reread(img: Image.Image, box: Tuple[float, float, float, float]) -> List[str]:
    """Second opinions (texts only) for a low-confidence box."""
    return [t for t, _ in read_crop(img, box)]


def _norm_reading(t: str) -> str:
    return re.sub(r"\s+", "", t).replace(",", "").translate(str.maketrans("oOlI|", "00111"))


def _split_across_columns(w: OCRWord, columns: List[Tuple[int, int]]) -> List[OCRWord]:
    """
    Text in two neighbouring cells that nearly touch is often recognised as one word
    ('5907.70RM' = rate | unit). Split such a word at the ruling it crosses: the cut goes to the
    character boundary nearest the ruling's position, preferring a digit/letter or space boundary.
    """
    inside = [(max(w.x, x0), min(w.x + w.w, x1)) for x0, x1 in columns]
    spans = [(a, b, ci) for ci, (a, b) in enumerate(inside) if b - a > 0.12 * w.w and b - a > 8]
    if len(spans) < 2 or len(w.text) < 2 or w.w > 3 * min(x1 - x0 for x0, x1 in columns[spans[0][2]:spans[-1][2] + 1]):
        return [w]
    text = w.text

    def is_num(ch: str) -> bool:
        return ch.isdigit() or ch in ".,"

    # only cut where the text itself changes kind (number <-> letters) or at a space:
    # a plain number such as '76' is never split
    allowed = [k for k in range(1, len(text))
               if text[k - 1] == " " or text[k] == " " or (is_num(text[k - 1]) != is_num(text[k])
                                                          and text[k - 1].isalnum() and text[k].isalnum())]
    if not allowed:
        return [w]
    cuts = []
    for a, b, ci in spans[:-1]:
        ideal = (b - w.x) / w.w * len(text)             # ruling position, in characters
        k = min(allowed, key=lambda kk: abs(kk - ideal))
        if abs(k - ideal) <= max(2.0, 0.25 * len(text)) and (not cuts or k > cuts[-1]):
            cuts.append(k)
    if not cuts:
        return [w]
    parts, prev = [], 0
    bounds = [w.x] + [b for a, b, _ in spans[:-1]] + [w.x + w.w]
    for i, k in enumerate(cuts + [len(text)]):
        piece = text[prev:k].strip()
        if piece:
            x0, x1 = bounds[i], bounds[i + 1]
            parts.append(OCRWord(piece, x0, w.y, max(1.0, x1 - x0), w.h, w.score))
        prev = k
    return parts or [w]


def _components(mask: np.ndarray) -> List[Tuple[np.ndarray, np.ndarray]]:
    """8-connected components of a small boolean mask (pure numpy/python; masks here are tiny)."""
    ys, xs = np.nonzero(mask)
    if len(ys) == 0:
        return []
    remaining = set(zip(ys.tolist(), xs.tolist()))
    comps = []
    while remaining:
        seed = remaining.pop()
        stack, comp = [seed], [seed]
        while stack:
            y, x = stack.pop()
            for dy in (-1, 0, 1):
                for dx in (-1, 0, 1):
                    p = (y + dy, x + dx)
                    if p in remaining:
                        remaining.discard(p)
                        stack.append(p)
                        comp.append(p)
        arr = np.array(comp)
        comps.append((arr[:, 0], arr[:, 1]))
    return comps


def _read_empty_cell(clean: Image.Image, dark: np.ndarray, x0: int, x1: int, y0: float, y1: float,
                     text_h: float) -> Optional[OCRWord]:
    """
    The detector occasionally skips a lone short token such as '1' or '0'. If a cell with no OCR
    word still holds glyph-shaped ink (not a ruling fragment, not dust), recognise the ink directly.
    text_h: typical text height on the page, used to tell glyphs from line fragments.
    """
    ya, yb = int(y0) + 12, int(y1) - 12
    xa, xb = int(x0) + 4, int(x1) - 4          # glyphs may sit right against a vertical ruling
    if yb - ya < 12 or xb - xa < 12:
        return None
    reg = dark[ya:yb, xa:xb]
    if reg.sum() < 60:
        return None
    # keep ink components that float inside the cell; ruling fragments run into its edges
    keep = np.zeros_like(reg)
    H_, W_ = reg.shape
    for ys_c, xs_c in _components(reg):
        if ys_c.min() <= 0 or ys_c.max() >= H_ - 1 or xs_c.min() <= 0 or xs_c.max() >= W_ - 1:
            continue
        if ys_c.max() - ys_c.min() + 1 > 1.6 * text_h or len(ys_c) < 4:
            continue
        keep[ys_c, xs_c] = True
    reg = keep
    if reg.sum() < 60:
        return None
    ys, xs = np.nonzero(reg)
    h, w = ys.max() - ys.min() + 1, xs.max() - xs.min() + 1
    # a glyph is roughly text-height tall, has some width, and does not run edge-to-edge
    if not (0.35 * text_h <= h <= 1.5 * text_h) or w < 0.06 * text_h:
        return None
    if ys.min() <= 1 or ys.max() >= reg.shape[0] - 2:
        return None
    bx0, by0, bx1, by1 = xa + xs.min(), ya + ys.min(), xa + xs.max(), ya + ys.max()
    reads = read_crop(clean, (bx0, by0, bx1, by1), scales=(1.0, 0.75, 1.4))
    if not reads:
        return None
    norm = [_norm_reading(t) for t, _ in reads]
    best = max(set(norm), key=norm.count)
    agree = norm.count(best)
    text = next(t for t, _ in reads if _norm_reading(t) == best)
    rapid_scores = [s for (t, s), n in zip(reads, norm) if n == best]
    score = max(rapid_scores) if agree >= 2 else min(0.6, max(rapid_scores))
    return OCRWord(text, float(bx0), float(by0), float(bx1 - bx0), float(by1 - by0), score=score,
                   alts=[t for t, _ in reads if _norm_reading(t) != best])


def detect_orientation(img: Image.Image) -> int:
    """Returns the PIL rotate() angle (counter-clockwise degrees) that makes the text upright."""
    small = img.copy()
    small.thumbnail((1400, 1400))
    small = ImageOps.autocontrast(ImageOps.grayscale(small), cutoff=1)
    good = re.compile(r"^(?:[A-Za-z][a-z]{2,}|[A-Z]{3,}|\d[\d,.]*\d)$")
    best_angle, best_score = 0, -1
    for angle in (0, 270, 90, 180):
        rot = small.rotate(angle, expand=True) if angle else small
        try:
            score = sum(len(w.text) for w in _ocr_windows(rot) if good.match(w.text))
        except Exception as ex:  # pragma: no cover
            logger.warning("Orientation probe failed at %s deg: %s", angle, ex)
            continue
        if score > best_score:
            best_angle, best_score = angle, score
    return best_angle


# ---------------------------------------------------------------------------
# Page loading
# ---------------------------------------------------------------------------

def open_pdf(pdf_source: PdfSource) -> pypdfium2.PdfDocument:
    if isinstance(pdf_source, (bytes, bytearray)):
        return pypdfium2.PdfDocument(bytes(pdf_source))
    if isinstance(pdf_source, io.BytesIO):
        return pypdfium2.PdfDocument(pdf_source.getvalue())
    return pypdfium2.PdfDocument(str(pdf_source))


def render_page_native(page) -> Image.Image:
    """Renders a page at the pixel density of its embedded scan (between ~216 and ~432 dpi)."""
    long_pt = max(page.get_size())
    native_long = 0
    for obj in page.get_objects():
        if obj.type == pypdfium2.raw.FPDF_PAGEOBJ_IMAGE:
            try:
                native_long = max(native_long, *obj.get_px_size())
            except Exception:
                pass
    scale = min(6.0, max(3.0, (native_long / long_pt) if native_long else 4.0))
    return page.render(scale=scale).to_pil().convert("RGB")


def is_scanned_pdf(pdf_source: PdfSource, min_chars: int = 100) -> bool:
    """True when the PDF has (almost) no extractable text layer."""
    pdf = open_pdf(pdf_source)
    chars = 0
    for i in range(len(pdf)):
        chars += len(pdf[i].get_textpage().get_text_range().strip())
        if chars > min_chars:
            return False
    return True


# ---------------------------------------------------------------------------
# Image analysis
# ---------------------------------------------------------------------------

def _dark_mask(gray: Image.Image) -> np.ndarray:
    """Ink mask that tolerates uneven phone-camera lighting (local background subtraction)."""
    w, h = gray.size
    bg = gray.resize((max(1, w // 24), max(1, h // 24)), Image.BOX).filter(ImageFilter.MaxFilter(3))
    bg = bg.resize((w, h), Image.BILINEAR)
    g = np.asarray(gray, dtype=np.int16)
    b = np.asarray(bg, dtype=np.int16)
    return (b - g) > 45


def _best_shear(ys: np.ndarray, xs: np.ndarray, size: int, max_deg: float = 3.0) -> float:
    """Slope t that best aligns rulings along y - t*x = const (maximises projection energy)."""
    best_t, best_score = 0.0, -1.0
    for deg in np.arange(-max_deg, max_deg + 1e-9, 0.05):
        t = np.tan(np.radians(deg))
        proj = np.round(ys - xs * t).astype(np.int64)
        proj -= proj.min()
        hist = np.bincount(proj, minlength=size)
        score = float((hist.astype(np.float64) ** 2).sum())
        if score > best_score:
            best_t, best_score = t, score
    return best_t


def deskew(gray: Image.Image) -> Image.Image:
    dark = _dark_mask(gray)
    f = 4
    h, w = dark.shape
    small = dark[: h - h % f, : w - w % f].reshape(h // f, f, w // f, f).mean(axis=(1, 3)) > 0.3
    ys, xs = np.nonzero(small)
    if len(ys) < 1000:
        return gray
    size = small.shape[0] + small.shape[1]
    t_h = _best_shear(ys.astype(np.float64), xs.astype(np.float64), size)
    t_v = _best_shear(xs.astype(np.float64), ys.astype(np.float64), size)
    if abs(t_h) < 1e-4 and abs(t_v) < 1e-4:
        return gray
    # output(x, y) samples input(x + t_v*y, y + t_h*x)
    return gray.transform(gray.size, Image.AFFINE, (1, t_v, 0, t_h, 1, 0),
                          resample=Image.BILINEAR, fillcolor=255)


def _peaks(profile: np.ndarray, thresh: float, merge_gap: int) -> List[int]:
    idx = np.nonzero(profile > thresh)[0]
    runs: List[List[int]] = []
    for i in idx:
        if runs and i - runs[-1][1] <= merge_gap:
            runs[-1][1] = i
        else:
            runs.append([int(i), int(i)])
    return [(a + b) // 2 for a, b in runs]


def _longest_runs(mask: np.ndarray, max_gap: int) -> np.ndarray:
    """Longest run of True along axis 0 for every column, bridging breaks of up to max_gap pixels."""
    run = np.zeros(mask.shape[1], dtype=np.int32)
    gap = np.zeros(mask.shape[1], dtype=np.int32)
    best = np.zeros(mask.shape[1], dtype=np.int32)
    for i in range(mask.shape[0]):
        row = mask[i]
        gap = np.where(row, 0, gap + 1)
        run = np.where(row, run + 1, np.where(gap <= max_gap, run, 0))
        np.maximum(best, run, out=best)
    return best


def _vertical_rulings(dark: np.ndarray, n_bands: Optional[int] = None, keep_frac: float = 0.5) -> List[int]:
    """
    x positions of vertical rulings.

    A photographed page curls, so a ruling is not one straight line from top to bottom. The page is
    cut into horizontal bands; inside a short band a ruling is a near-straight stroke covering most
    of the band (dashes bridged), while text never is. Band hits are chained downwards allowing a
    few pixels of drift per band, and a chain must span a large part of the table to count
    (keep_frac of the longest chain).
    """
    H = dark.shape[0]
    m = dark.copy()
    for s in (1, 2):
        m[:, s:] |= dark[:, :-s]
        m[:, :-s] |= dark[:, s:]
    n_bands = n_bands or max(4, H // 300)
    edges = np.linspace(0, H, n_bands + 1).astype(int)
    band_hits: List[List[int]] = []
    for a, b in zip(edges, edges[1:]):
        best = _longest_runs(m[a:b], max_gap=12)
        band_hits.append(_peaks(best.astype(np.float32), 0.6 * (b - a), 6))

    chains: List[List[Tuple[int, int]]] = []          # [(band, x), ...]
    for bi, xs in enumerate(band_hits):
        for x in xs:
            best_chain, best_d = None, None
            for ch in chains:
                lb, lx = ch[-1]
                if lb < bi and bi - lb <= 2 and abs(lx - x) <= 14 * (bi - lb):
                    d = abs(lx - x)
                    if best_d is None or d < best_d:
                        best_chain, best_d = ch, d
            if best_chain is not None and best_chain[-1][0] != bi:
                best_chain.append((bi, x))
            else:
                chains.append([(bi, x)])
    # the edge of the paper / photo frame is a full-height dark line hugging the image border: it is
    # not part of the table and must not set the length a table ruling has to reach
    W = dark.shape[1]
    margin = max(10, int(0.012 * W))
    chains = [c for c in chains if margin < np.median([x for _, x in c]) < W - margin]
    if not chains:
        return []
    longest = max(len(c) for c in chains)
    keep = [c for c in chains if len(c) >= max(2, keep_frac * longest)]
    xs = sorted(int(np.median([x for _, x in c])) for c in keep)
    merged: List[int] = []
    for x in xs:
        if merged and x - merged[-1] <= 8:
            continue
        merged.append(x)
    return merged


def _ruling_span(dark: np.ndarray, x: int, max_gap: int = 12) -> Tuple[int, int]:
    """(top, bottom) of the longest near-continuous ink run down the strip around x."""
    col = dark[:, max(0, x - 3): x + 4].any(axis=1)
    best, start, last = (0, 0), None, -10 ** 9
    for y in np.nonzero(col)[0]:
        if start is None or y - last > max_gap:
            start = y
        last = y
        if last - start > best[1] - best[0]:
            best = (int(start), int(last))
    return best


def _solid_line(dark: np.ndarray, x: int) -> bool:
    """A drawn line is inked along (nearly) all of its length; letters of centred bold text that happen
    to line up over a few rows are not (measured: lines >= 0.97, text <= 0.62)."""
    a, b = _ruling_span(dark, x)
    return max(dark[a:b + 1, max(0, x + dx - 1): x + dx + 2].any(axis=1).mean() for dx in range(-4, 5)) >= 0.8


def _leaning_rulings(dark: np.ndarray, xs: List[int]) -> List[int]:
    """
    Lines leaning by a degree or two (one side of a photographed page curls, the other does not)
    leave the narrow strip a page-sized band is tested in. Narrower bands keep them inside it, and a
    chain through many narrow bands is itself the proof of a drawn line (text never runs that long
    vertically); lines not found already are added.
    """
    if len(xs) < 3:
        return xs                     # no table on this page: nothing to complete
    H = dark.shape[0]
    fine = _vertical_rulings(dark, n_bands=max(4, H // 150))
    extra = [x for x in fine if all(abs(x - x0) > 15 for x0 in xs)]
    return sorted(xs + extra)


def _short_table_rulings(dark: np.ndarray, xs: List[int]) -> List[int]:
    """
    Second pass for a table that fills only part of the page (last page of an abstract: three items,
    then the Sub Total / contingency rows drawn without the inner column lines). The page-sized bands
    of the first pass are taller than those inner lines, so they are looked for again inside the
    table's own height with narrow bands, keeping any line that runs through at least two of them.
    """
    def solid(x: int) -> bool:
        return _solid_line(dark, x)

    spans = [_ruling_span(dark, x) for x in xs]
    if not spans:
        # nothing found at page scale: a table a few rows high (the last abstract page) spans only one
        # page-sized band - look again with narrow bands, keeping only solid lines of some length
        H = dark.shape[0]
        found = _vertical_rulings(dark, n_bands=max(4, H // 100), keep_frac=0.25)
        found = [x for x in found if solid(x) and (lambda s: s[1] - s[0])(_ruling_span(dark, x)) >= max(150, 0.04 * H)]
        return found if len(found) >= 3 else xs
    longest = max(b - a for a, b in spans)
    top = min(a for a, b in spans if b - a >= 0.5 * longest)
    bot = max(b for a, b in spans if b - a >= 0.5 * longest)
    if bot - top < 100:
        return xs
    found = _vertical_rulings(dark[top:bot + 1], n_bands=max(4, (bot - top) // 100), keep_frac=0.25)
    found = [x for x in found if any(abs(x - x0) <= 8 for x0 in xs) or solid(x)]
    return found if len(found) > len(xs) else xs


def _table_extent(dark: np.ndarray, xs: List[int]) -> Tuple[int, int]:
    """Rows where most vertical rulings are inked = the table body (robust to page curl)."""
    H = dark.shape[0]
    if len(xs) < 3:
        return 0, H
    hits = np.zeros(H, dtype=np.int32)
    for x in xs:
        hits += dark[:, max(0, x - 4): x + 5].any(axis=1)
    inside = hits >= max(2, int(len(xs) * 0.5))
    k = 25
    smooth = np.convolve(inside.astype(np.float32), np.ones(k) / k, mode="same") > 0.5
    rows = np.nonzero(smooth)[0]
    if len(rows) == 0:
        return 0, H
    return int(rows[0]), int(rows[-1])


def _horizontal_rulings(dark: np.ndarray, x0: int, x1: int, min_frac: float = 0.6) -> List[int]:
    """
    Rows inside [x0, x1) crossed by one long unbroken stroke. Text never qualifies because its
    glyphs are separated by gaps; a ruling that drifts ±3 px across the column (page curl) still does.
    """
    w = x1 - x0
    if w < 10:
        return []
    src = dark[:, x0:x1]
    # vertical tolerance for slightly slanted / curved rulings (wider columns drift more)
    m = src.copy()
    for s in range(1, max(2, int(w * 0.01)) + 1):
        m |= np.roll(src, s, axis=0) | np.roll(src, -s, axis=0)
    run = np.zeros(m.shape[0], dtype=np.int32)
    gap = np.zeros(m.shape[0], dtype=np.int32)
    best = np.zeros(m.shape[0], dtype=np.int32)
    for j in range(w):
        col = m[:, j]
        gap = np.where(col, 0, gap + 1)
        run = np.where(col, run + 1, np.where(gap <= 2, run, 0))   # bridge 1-2 px breaks in faint ink
        np.maximum(best, run, out=best)
    return _peaks(best.astype(np.float32), min_frac * w, 4)


def _track_rulings(col_lines: List[List[int]]) -> Tuple[List[Dict[int, float]], Dict[int, float]]:
    """
    Links each column's rulings to its neighbours' into global row-rulings ("tracks"), so a ruling
    that drifts a few pixels across a curled page is still recognised as the same table row border.
    Returns the tracks (col -> y) sorted top-to-bottom, and per-column y offsets relative to the
    reference column (used to place rulings that a column does not draw itself).
    """
    n = len(col_lines)
    if n == 0:
        return [], {}
    ref = int(np.argmax([len(ls) for ls in col_lines]))
    ref_lines = sorted(col_lines[ref])
    gaps = np.diff(ref_lines) if len(ref_lines) > 1 else np.array([100])
    tol = float(np.clip(np.median(gaps) * 0.4, 15, 45))

    tracks: List[Dict[int, float]] = [{ref: float(y)} for y in ref_lines]

    def nearest_known(t: Dict[int, float], c: int, direction: int) -> Optional[float]:
        # y of this track at the nearest processed column on the side we came from
        cc = c - direction
        while 0 <= cc < n:
            if cc in t:
                return t[cc]
            cc -= direction
        return None

    for direction, order in ((1, range(ref + 1, n)), (-1, range(ref - 1, -1, -1))):
        for c in order:
            ys = sorted(col_lines[c])
            pairs = []
            for ti, t in enumerate(tracks):
                est = nearest_known(t, c, direction)
                if est is None:
                    continue
                for yi, y in enumerate(ys):
                    d = abs(y - est)
                    if d <= tol:
                        pairs.append((d, ti, yi))
            used_t, used_y = set(), set()
            for d, ti, yi in sorted(pairs):
                if ti in used_t or yi in used_y:
                    continue
                tracks[ti][c] = float(ys[yi])
                used_t.add(ti)
                used_y.add(yi)
            for yi, y in enumerate(ys):
                if yi not in used_y:
                    tracks.append({c: float(y)})

    # per-column offset to the reference column (median over shared tracks)
    offsets: Dict[int, float] = {ref: 0.0}
    for c in range(n):
        if c == ref:
            continue
        diffs = [t[ref] - t[c] for t in tracks if ref in t and c in t]
        if diffs:
            offsets[c] = float(np.median(diffs))
    for c in range(n):  # columns without shared tracks inherit their neighbour's offset
        if c not in offsets:
            near = min(offsets.keys(), key=lambda k: abs(k - c))
            offsets[c] = offsets[near]

    def ref_y(t: Dict[int, float]) -> float:
        c = min(t.keys(), key=lambda k: abs(k - ref))
        return t[c] + offsets[c]

    # a "ruling" seen in a single column is a pen underline / strike-through, not a table row border
    if n > 2:
        tracks = [t for t in tracks if len(t) >= 2]
    tracks.sort(key=ref_y)
    # merge tracks that ended up describing the same ruling
    merged: List[Dict[int, float]] = []
    for t in tracks:
        if merged and abs(ref_y(t) - ref_y(merged[-1])) < tol * 0.5 and not (set(t) & set(merged[-1])):
            merged[-1].update(t)
        else:
            merged.append(t)
    return merged, offsets


def _track_y(t: Dict[int, float], c: int, offsets: Dict[int, float]) -> float:
    if c in t:
        return t[c]
    k = min(t.keys(), key=lambda kk: abs(kk - c))
    return t[k] + offsets.get(k, 0.0) - offsets.get(c, 0.0)


# ---------------------------------------------------------------------------
# Main entry points
# ---------------------------------------------------------------------------

def extract_table(img: Image.Image, page_number: int = 1, rotation: int = 0,
                  reread_below: float = 0.97) -> ScannedTable:
    gray = deskew(ImageOps.grayscale(img))
    dark = _dark_mask(gray)
    H, W = dark.shape

    xs = _vertical_rulings(dark)
    if len(xs) < 6:                   # a TS / abstract table has at least 6 column lines
        xs = _short_table_rulings(dark, xs)
    xs = _leaning_rulings(dark, xs)
    t_top, t_bot = _table_extent(dark, xs)
    columns = [(int(xs[i]), int(xs[i + 1])) for i in range(len(xs) - 1) if xs[i + 1] - xs[i] > 25]

    clean = ImageOps.autocontrast(gray, cutoff=1)
    page_words = ocr_words(clean)

    # Outer border cut off by the photo: short tokens lined up just outside the first/last ruling
    # (e.g. the Sl. No column) still form a column of the table.
    if columns:
        def margin_words(side: int) -> List[OCRWord]:
            edge = columns[0][0] if side < 0 else columns[-1][1]
            out = []
            for w in page_words:
                if not (t_top - 10 <= w.cy <= t_bot + 10) or len(w.text) > 6:
                    continue
                if side < 0 and edge - 260 <= w.x and w.x + w.w <= edge + 8:
                    out.append(w)
                if side > 0 and edge - 8 <= w.x and w.x + w.w <= edge + 260:
                    out.append(w)
            return out
        lw = margin_words(-1)
        if len(lw) >= 2:
            columns.insert(0, (max(0, int(min(w.x for w in lw)) - 25), columns[0][0]))
        rw = margin_words(1)
        if len(rw) >= 2:
            columns.append((columns[-1][1], min(W, int(max(w.x + w.w for w in rw)) + 25)))
        columns = [(a, b) for a, b in columns if b - a > 25]

    col_lines: List[List[int]] = []
    for x0, x1 in columns:
        m = max(6, (x1 - x0) // 12)
        col_lines.append([y for y in _horizontal_rulings(dark, x0 + m, x1 - m) if t_top - 8 <= y <= t_bot + 8])

    tracks, offsets = _track_rulings(col_lines)
    n_rows = max(0, len(tracks) - 1)

    # merged-cell spans per column from that column's own rulings
    cells: List[List[Cell]] = []
    for c in range(len(columns)):
        own = [ti for ti, t in enumerate(tracks) if c in t]
        col_cells = [Cell(col=c, row_start=a, row_end=b - 1) for a, b in zip(own, own[1:]) if b > a]
        cells.append(col_cells)

    def place(words: List[OCRWord]):
        """Each word with its (row, column)."""
        placed: List[OCRWord] = []
        for w in words:
            placed.extend(_split_across_columns(w, columns))
        for w in placed:
            overlaps = [(min(w.x + w.w, x1) - max(w.x, x0), ci) for ci, (x0, x1) in enumerate(columns)]
            best = max(overlaps, default=(0, -1))
            if best[0] <= 0:
                continue
            ci = best[1]
            if best[0] < 0.5 * w.w and w.w > (columns[ci][1] - columns[ci][0]) * 1.5:
                continue   # a title spanning many columns, not a cell value
            for ri in range(n_rows):
                if _track_y(tracks[ri], ci, offsets) <= w.cy < _track_y(tracks[ri + 1], ci, offsets):
                    yield ri, ci, w
                    break

    # place each OCR word in its (row, column)
    grid: Dict[Tuple[int, int], List[OCRWord]] = {}
    for row, ci, w in place(page_words):
        if w.score < reread_below and re.search(r"\d", w.text):
            w.alts = reread(clean, (w.x, w.y, w.x + w.w, w.y + w.h))
        grid.setdefault((row, ci), []).append(w)
    alt_grid: Dict[Tuple[int, int], List[OCRWord]] = {}
    if _get_rapid() is not None:           # (without RapidOCR, Windows OCR is already the main engine)
        try:
            for row, ci, w in place(_ocr_windows(clean)):
                alt_grid.setdefault((row, ci), []).append(w)
        except Exception as ex:               # Windows OCR unavailable: single-engine reading
            logger.info("second OCR engine unavailable: %s", ex)
    # low-confidence text outside the grid (titles: scheme, contractor ...) gets a second opinion too
    for w in page_words:
        if w.score < 0.85 and not w.alts and not re.fullmatch(r"[\d.,%\s-]*", w.text):
            w.alts = reread(clean, (w.x, w.y, w.x + w.w, w.y + w.h))

    # cells the detector skipped although they hold ink (lone '1' / '0' digits)
    text_h = float(np.median([w.h for w in page_words])) if page_words else 40.0
    for ci, (x0, x1) in enumerate(columns):
        for ri in range(n_rows):
            if (ri, ci) in grid:
                continue
            a, b = next(((c.row_start, c.row_end) for c in cells[ci] if c.row_start <= ri <= c.row_end), (ri, ri))
            if a != b:
                continue          # merged cell: its text sits in another row of the span
            y0 = _track_y(tracks[ri], ci, offsets)
            y1 = _track_y(tracks[ri + 1], ci, offsets)
            word = _read_empty_cell(clean, dark, x0, x1, y0, y1, text_h)
            if word is not None:
                grid[(ri, ci)] = [word]

    top =_track_y(tracks[0], 0, offsets) if tracks else 0.0
    bottom = _track_y(tracks[-1], 0, offsets) if tracks else float(H)
    return ScannedTable(page_number=page_number, width=W, height=H, rotation=rotation, columns=columns,
                        n_rows=n_rows, cells=cells, grid_words=grid, page_words=page_words,
                        top=top, bottom=bottom, alt_grid_words=alt_grid)


def quick_page_text(pdf_source: PdfSource, scale: float = 1.6) -> List[str]:
    """Fast, low-resolution text of every page (for locating pages; not for reading values)."""
    pdf = open_pdf(pdf_source)
    out = []
    for i in range(len(pdf)):
        tp = pdf[i].get_textpage().get_text_range().strip()
        if len(tp) > 50:
            out.append(tp)
            continue
        img = ImageOps.grayscale(pdf[i].render(scale=scale).to_pil())
        text = _quick_ocr(img)
        if _readability(text) < 3:
            # upside-down scans and sideways pages (comparative statements) read as gibberish
            for angle in (180, 90, 270):
                alt = _quick_ocr(img.rotate(angle, expand=True))
                if _readability(alt) > max(_readability(text), 2):
                    text = alt
                    break
        out.append(text)
    return out


_COMMON_WORDS = re.compile(r"\b(of|the|and|for|work|works|construction|total|estimate|installation|pipe|job|nos|"
                           r"rate|amount|water|mission|name|supply|supplying|sanction|scheme|dia|engineer)\b", re.I)


def _readability(text: str) -> int:
    return len(_COMMON_WORDS.findall(text or ""))


def _quick_ocr(img: Image.Image) -> str:
    try:
        return " ".join(w.text for w in _ocr_windows(img))
    except Exception:
        return " ".join(w.text for w in ocr_words(img))


def extract_pdf_tables(pdf_source: PdfSource, pages: Optional[List[int]] = None) -> List[ScannedTable]:
    """Reads the pages of a scanned PDF (all, or the given 0-based page indexes) into ScannedTable objects."""
    return list(iter_pdf_tables(pdf_source, pages))


def iter_pdf_tables(pdf_source: PdfSource, pages: Optional[List[int]] = None):
    """extract_pdf_tables one page at a time (the caller may stop early)."""
    pdf = open_pdf(pdf_source)
    orientation: Optional[int] = None
    for i in (pages if pages is not None else range(len(pdf))):
        img = render_page_native(pdf[i])
        # pages of one scan session share an orientation; re-probe only if the grid looks wrong
        angle = detect_orientation(img) if orientation is None else orientation
        tbl = extract_table(img.rotate(angle, expand=True) if angle else img, i + 1, angle)
        if orientation is not None and tbl.n_cols < 3:
            angle = detect_orientation(img)
            tbl = extract_table(img.rotate(angle, expand=True) if angle else img, i + 1, angle)
        orientation = angle
        yield tbl
