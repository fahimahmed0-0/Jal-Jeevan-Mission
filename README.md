# JJM PHED Assam — TPIA Abstract of Bill (Payment Recommendation)

A desktop tool for Third Party Inspection Agencies working under PHED Assam (Jal Jeevan Mission).
It turns a contractor's bill into the official **Abstract of Bill (Payment Recommendation)** Excel
workbook, and **proves** every figure before the file can be downloaded.

## Quick start (Windows)

1. Double-click `run.bat` (installs dependencies on first run).
2. The app opens at `http://localhost:8501`.
3. Load your documents (or a sample from the sidebar) and click **Process**.

## What goes in

| Slot | Required | What it is |
|---|---|---|
| Contractor bill (PDF) | yes | The digital portal e-bill (Running Account Bill-C from jjmassam.in); a scanned TPIA abstract is also accepted |
| Technical Sanction document (PDF) | yes, for the official layout | The whole TS document (scanned or digital). Its **Abstract of Cost Estimate for Revised Technical Sanction** pages are found automatically (or type the page numbers, e.g. `9-10`) |
| Previous TPIA abstract (Excel) | optional | Only used when no TS is given: the new bill is written into that workbook |

## How the workbook is built (TS + e-bill)

* **From the TS abstract:** every item with its quantity, rate and amount (columns A–F), arranged in
  the standard TPIA order (DTW, pumping mains, pumps, TP, UGR, FHTC, ESR, distribution pipes, valve
  chambers, then OTHER WORKS), numbered serially. Items with zero amount are left out. When the printed
  quantity disagrees with Amount ÷ Rate (or is blank) the quantity is Amount ÷ Rate. The **Sub Total** is
  the contract total and the Work Order Amount used for progress.
* **Milestones (column G = rate × %):** pumping main 30/60/10 · pumps 40/60 · UGR 40/60 ·
  ESR 15/25/40/20 · HDPE distribution 20/50/20/10 · pump house 40/60 · power connection 80/20 ·
  everything else 100%.
* **From the e-bill:** each line is one milestone (rate = TS rate; milestone % = Up-to-date ÷ (Qty × Rate)).
  Previous bill = Up-to-date − Since previous bill (columns H/K), this bill = Since previous bill (I/L).
  Recoveries (negative this-bill lines) are written like the engineer's abstract: previous amount as paid,
  this bill `=F−K` (whole item) or `=D×G−K` (one milestone).
* **Checks:** contract total = TS Sub Total, previous = e-bill "Deduct value … previous bill",
  this bill = e-bill net, up to date = e-bill total — then the Excel file is recalculated and compared.

## How accuracy is guaranteed

No value is assumed, defaulted or "healed" from a sample scheme. Every number either comes straight
from a document **and** is confirmed by that document's own arithmetic, or it is highlighted for review.

**Digital e-bill** — read from the PDF text layer (exact). Checked: lines add up to the bill's
*Net value since previous bill* and *Total*, and *Total − Previous = Net*.

**Digital e-bill + previous abstract** — every e-bill line is placed on the template row(s) where all
of these hold exactly: rate = contract rate, amount this bill = L, amount up-to-date = K + L, and
quantity up-to-date = H + I. A line that fits nowhere is reported, never guessed. The amounts shown
are produced by evaluating the template's own formulas.

**Scanned abstract / TS** (`table_extractor.py`, `scanned_abstract.py`, `ts_parser.py`):
1. pages are rendered at the scan's native resolution, orientation and skew are corrected;
2. the ruled grid is detected (curled photos handled) and every cell is read with RapidOCR
   (offline PP-OCR models; Windows OCR as a second opinion / fallback);
3. each line is resolved with *Prev + This = Upto*, *This = Qty × Rate paid*,
   *Rate paid = Rate × milestone %*, each item with *Qty × Rate = Amount*, and the whole bill with the
   printed **Total** row. Repairs are only accepted when another identity confirms them.

**Excel output** (`excel_writer.py`) — the workbook is recalculated by `formula_eval.py` and compared
with the reviewed figures line by line (L, K, M), in total, the ROUNDUP payment, the amount in words
and the progress %. The download button only appears when all of that matches.

## Review screen

* ✅ / ❌ / ⚠️ checks with the actual numbers (nothing is a fixed message).
* Every line with its status: 🟢 read & confirmed, 🟡 corrected with evidence, 🔴 needs review,
  ✏️ edited by the engineer. Uncertain lines are highlighted in the table and in the Excel file.
* Edit quantities/amounts (and, for scanned abstracts, rates); totals, checks and the workbook update.
* If critical checks fail, the engineer must confirm the flagged lines before downloading.

## Project layout

```
app.py              Streamlit UI
pipeline.py         Routing: bill (+TS) (+template) -> Abstract
ts_ebill_builder.py Main route: TS abstract layout + e-bill amounts -> Abstract
abstract_model.py   The single data model, totals and checks
ebill_parser.py     Digital Running Account Bill-C parser
scanned_abstract.py Scanned TPIA abstract -> Abstract (grid OCR + arithmetic resolution)
ts_parser.py        Technical Sanction parser (digital or scanned)
item_mapper.py      E-bill -> template rows (exact), e-bill-only abstract, TS benchmark
table_extractor.py  Scanned table reader (grid detection + OCR)
excel_writer.py     Official-format writer, template filler, workbook verification
formula_eval.py     Excel formula evaluator used for verification
utils.py            Indian currency format, amount in words, number parsing
tests/              pytest suite with ground truth for the sample documents
```

## Tests

```bash
python -m pytest          # or: python test_system.py
```

The suite OCRs the scanned samples (1–2 minutes), rebuilds the Bengsatra abstract from the 96-page TS
document + e-bill and compares it row by row with the engineer's workbook (`Doc/BISONDOI PT-II BANG SATRA.xlsx`),
compares all 62 payment lines of the Tolli abstract with a hand-checked transcription, and recalculates
every generated workbook.
