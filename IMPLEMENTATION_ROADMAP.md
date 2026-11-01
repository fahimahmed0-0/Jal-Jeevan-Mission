# JJM PHED Assam: E-Bill to TPIA Abstract System
## Phase-by-Phase Technical Implementation Roadmap

> **Historical planning document.** The current design is described in `README.md`. Two points below
> were superseded: e-bill lines are now mapped by exact amount identities (not keywords), and e-bill
> line 23 (APDCL, 100%) covers both milestone rows 62 and 63 of the BISONDOI template.

---

### Executive Overview

This document provides the definitive, step-by-step technical plan to build the automated civil engineering billing tool for **Public Health Engineering Department (PHED), Assam (Jal Jeevan Mission)**.

The system automates the transition from a contractor's digital **E-Bill PDF** (Running Account Bill-C) to a fully compliant **TPIA Abstract of Bill (Payment Recommendation)** Excel document, preserving all existing Excel formulas, milestone percentages, and formatting.

---

## Architecture Flow

```
┌─────────────────────────────────────────────────────────────┐
│                   STREAMLIT WEB UI (app.py)                 │
│                                                             │
│  [ Upload E-Bill PDF ]     [ Upload Prev/Template Excel ]   │
│                             │                               │
│                             ▼                               │
│                  [ Process & Match Items ]                  │
│                             │                               │
│  ┌──────────────────────────▼────────────────────────────┐  │
│  │         Interactive Review & Verification Table        │  │
│  │   • Matched Items (26/26)    • Quantity Overrides     │  │
│  │   • Confidence Indicators    • Negative Alert Badges  │  │
│  └──────────────────────────┬────────────────────────────┘  │
│                             │                               │
│                  [ Generate & Download ]                    │
└─────────────────────────────┼───────────────────────────────┘
                              │
       ┌──────────────────────┼──────────────────────┐
       ▼                      ▼                      ▼
┌──────────────┐      ┌──────────────┐      ┌────────────────┐
│ ebill_parser │      │ item_mapper  │      │  excel_writer  │
│              │      │              │      │                │
│ • pdfplumber │─────▶│ • 3-Layer    │─────▶│ • openpyxl     │
│ • 26 items   │      │   matching   │      │ • Writes Col   │
│ • Metadata   │      │ • Quantities │      │   H & I only   │
│ • Summary    │      │ • Overrides  │      │ • Formulas intact│
└──────────────┘      └──────────────┘      └────────────────┘
```

---

## Detailed Implementation Phases

### Phase 1: Core Parsing Engine (`ebill_parser.py`)
**Objective**: Build a robust, production-grade parser for PHED Assam Running Account Bill-C PDFs.

- **Step 1.1: Document Structure & Metadata Extraction**
  - Extract Scheme Name, Work Order Number, Contractor Name, Bill Serial Number, and Date.
  - Extract portal summary totals:
    - Total value of work done to date (`total_upto_date`)
    - Deduct value of work shown on previous bill (`deduct_previous`)
    - Net value of work since previous bill (`net_this_bill`)
- **Step 1.2: Multi-Page Table Extraction**
  - Iterate through all pages of the E-Bill using `pdfplumber`.
  - Handle table boundaries, page headers/footers, and multi-line item descriptions cleanly.
- **Step 1.3: Data Normalization & Cleaning**
  - Parse Indian currency strings, commas, whitespace, and negative values (`-30.00`, `-32536.00`).
  - Extract 26 line items with: `sl`, `unit`, `qty`, `desc`, `rate`, `upto_amt`, `since_amt`.
- **Step 1.4: Output Schema & Self-Validation**
  - Return strongly-typed data structures (`dict` / `dataclass`).
  - Add internal validation assert: verify $\sum \text{since\_amt} == \text{net\_this\_bill}$ to confirm zero data loss.

---

### Phase 2: Intelligent 3-Layer Item Mapping Engine (`item_mapper.py`)
**Objective**: Reliably map each E-Bill item to its corresponding row in the Excel template without manual row lookup.

- **Step 2.1: Excel Template Ingestion**
  - Load the active Excel template sheet (`Sheet1`) using `openpyxl`.
  - Read Rows 12 to 74 to build an in-memory catalog of parent items and milestone sub-items:
    - Row index, Item Serial No (Col A), Description (Col B), Unit (Col C), Contract Qty (Col D), Contract Rate (Col E), Milestone Rate (Col G), Milestone % (Col O), Remarks (Col N).
- **Step 2.2: Layer 1 — Contract Rate Matching**
  - For each E-Bill item, query the template catalog for matching Contract Rates (`Col E`):
    - Absolute difference tolerance: $|Rate_{ebill} - Rate_{excel}| < 1.0$.
    - Single match &rarr; Direct match confirmed (e.g., Sl#1 DTW, Sl#6 TP, Sl#18 Boundary Wall).
- **Step 2.3: Layer 2 — Milestone Ratio Disambiguation**
  - If multiple candidate rows share the same contract rate (e.g., UGR milestones, Pump House milestones, HDPE pipe diameters):
    - Compute the ratio: $\text{ratio} = \frac{\text{upto\_amt}}{\text{contract\_qty} \times \text{contract\_rate}}$.
    - Match this ratio against `Col O` (Milestone Percentage).
    - Resolves 100% of ambiguous sub-items (e.g., UGR 40% vs 60%, Pipe Supply 20% vs Laying 50%).
- **Step 2.4: Layer 3 — Keyword Safety Net**
  - Fallback logic using regular expressions and keywords (pipe diameter `110mm`, `90mm`; structure type `UGR`, `ESR`, `Pump House`).
- **Step 2.5: Quantity Computation & Reconciliation**
  - Calculate `qty_this_bill` for each mapped row:
    - Simple items: $\frac{\text{since\_amt}}{\text{contract\_rate}}$
    - Milestone items: $\frac{\text{since\_amt}}{\text{milestone\_rate}}$ (where $\text{milestone\_rate} = \text{Col G}$)
  - Handle negative corrections gracefully by setting negative executed quantities or flagging them for engineer confirmation.

---

### Phase 3: Excel Generation & Formula Preservation (`excel_writer.py`)
**Objective**: Populate the engineer's Excel template with surgical precision while preserving all styles, merges, and mathematical formulas.

- **Step 3.1: Template Cloning**
  - Open the source template workbook with `openpyxl.load_workbook(template_path, data_only=False)` to prevent stripping Excel formulas.
- **Step 3.2: Cell Value Injection**
  - **Write to Column H (Upto Previous Bill Qty)**: Populate from previous bill's Col J or carry forward existing values.
  - **Write to Column I (This Bill Qty)**: Populate calculated `qty_this_bill` for all matched rows.
  - All non-billed rows remain at `0` in Column I.
- **Step 3.3: Header & Metadata Synchronization**
  - Update Scheme Name (Cell C4), Contractor Name (Cell C5), Work Order No (Cell C6), Completion Date (Cell C9).
  - Update Bill Recommendation summary rows if required (Row 77, Row 78 "In Words").
- **Step 3.4: Formula Integrity Check**
  - Ensure zero edits to:
    - Column F (`=D*E`)
    - Column G (`=E*pct`)
    - Column J (`=I+H`)
    - Column K (`=G*H`)
    - Column L (`=I*G` or `=F-K` for reconciliation items)
    - Column M (`=K+L`)
    - Row 75 Totals (`=SUM(F12:F74)`, `=SUM(K12:K74)`, `=SUM(L12:L74)`, `=SUM(M12:M74)`)
    - Rows 80 & 81 Progress (`=M75/F76`)

---

### Phase 4: Utilities & Formatting Helpers (`utils.py`)
**Objective**: Provide utility functions for data formatting and calculations.

- **Step 4.1: Indian Currency Formatting**
  - Convert standard floats to Indian comma notation (e.g., `25,79,768.66`).
- **Step 4.2: Amount to Words Converter**
  - Convert total recommended bill amount into Indian numbering text (e.g., *"Twenty Five Lakh Seventy Nine Thousand Seven Hundred Sixty Nine Only"*).
- **Step 4.3: Sanity & Audit Loggers**
  - Helper to generate structured JSON summary logs of each run for auditing.

---

### Phase 5: Streamlit Web UI (`app.py`)
**Objective**: Build a clean, professional, and intuitive user interface tailored for civil engineers.

- **Step 5.1: File Upload & Configuration Area**
  - E-Bill PDF drag-and-drop uploader.
  - Previous Bill / Blank Template Excel uploader.
  - Single-click action button: `[ Process & Map Bill ]`.
- **Step 5.2: Mapping Review & Verification Table**
  - Display parsed E-Bill items alongside matched Excel rows.
  - Color-coded status badges:
    - 🟢 **Clean Match**: Exact rate & milestone matched.
    - 🟡 **Warning**: Quantity exceeds contract limit.
    - 🔴 **Negative Adjustment**: Over-payment deduction flagged.
  - Editable columns allowing the engineer to tweak `This Bill Qty` on the fly before exporting.
- **Step 5.3: Executive Summary & Financial Card**
  - Cards showing Total Work Order, Previous Bill, Net This Bill, and Estimated Physical Progress %.
- **Step 5.4: Export & Download Action**
  - One-click `[ Download Generated Excel ]` generating the validated, finalized `.xlsx` file.

---

### Phase 6: End-to-End Validation & Verification
**Objective**: Rigorously test the entire workflow against real project data.

- **Step 6.1: Run 2nd RA Bill Test (Bengsatra)**
  - Test input: `E-Bill__ PHED, ASSAM bengsatra.pdf` + `BISONDOI PT-II BANG SATRA.xlsx`.
  - Verify all 26 items mapped to Rows 12 to 74.
  - Verify total net recommended matches `₹ 25,79,768.66`.
- **Step 6.2: Run 1st RA Bill Simulation**
  - Run with blank previous quantities (Col H = 0).
  - Verify Col I populates correctly and totals match first bill expectations.
- **Step 6.3: Excel Open & Recalculation Check**
  - Open generated Excel in Microsoft Excel to confirm no repair prompts or broken formulas appear.

---

### Phase 7: Packaging & Delivery
**Objective**: Make the tool zero-friction to run for the engineer.

- **Step 7.1: Dependencies File (`requirements.txt`)**
  - `pdfplumber>=0.11.0`
  - `openpyxl>=3.1.2`
  - `pandas>=2.0.0`
  - `streamlit>=1.35.0`
- **Step 7.2: Windows Batch Launcher (`run.bat`)**
  - Double-click script that launches the Streamlit app in the user's default web browser automatically.
- **Step 7.3: User Operation Guide (`README.md`)**
  - Concise step-by-step user instructions with screenshots and operational guidelines.

---

---

## Verified Forensic Mapping Table (26 E-Bill Items &rarr; Excel Rows)

This table cross-references each line item from the contractor's E-Bill with the exact target row in the TPIA Abstract Excel template:

| E-Bill # | Item Description | E-Bill Rate (₹) | Target Excel Row | Excel Parent Item & Milestone | Milestone % (Col O) | Match Method | Verified Status |
|:---:|:---|:---:|:---:|:---|:---:|:---:|:---:|
| **1** | TP, 4SQM | 4,05,640.00 | **Row 23** | Treatment Plant (Sl#6) | 100% | Layer 1: Rate | 🟢 Match |
| **2** | UGR 30KL (40%) | 3,78,000.00 | **Row 24** | UGR 30 KL - Sump Wall (Sl#7) | 40% | Layer 2: Ratio 0.40 | 🟢 Match |
| **3** | UGR 30KL (60%) | 3,78,000.00 | **Row 25** | UGR 30 KL - Sump Roof (Sl#7) | 60% | Layer 2: Ratio 0.60 | 🟢 Match |
| **4** | Sub-Soil Investigation | 1,34,435.00 | **Row 65** | Sub-Soil Investigation (Sl#25) | 100% | Layer 1: Rate | 🟢 Match |
| **5** | ESR 50KL (15%) | 30,25,575.00 | **Row 27** | ESR 50 KL - Foundation (Sl#9) | 15% | Layer 2: Ratio 0.15 | 🟢 Match |
| **6** | Site Development | 342.00 | **Row 61** | Site Development (Sl#22) | 100% | Layer 1: Rate | 🟢 Match |
| **7** | Sign Board | 42,900.00 | **Row 56** | Sign Board (Sl#16) | 100% | Layer 1: Rate | 🟢 Match |
| **8** | Site Clearance | 1,50,000.00 | **Row 73** | Site Clearance / Camp (Sl#34) | 100% | Layer 1: Rate | 🟢 Match |
| **9** | Brick Drain | 1,100.00 | **Row 71** | Construction of Drains (Sl#31) | 100% | Layer 1: Rate | 🟢 Match |
| **10** | Soak Pit | 5,576.00 | **Row 74** | Community Soak Pit (Sl#35) | 100% | Layer 1: Rate | 🟢 Match |
| **11** | DTW 100m | 5,907.70 | **Row 12** | Installation of DTW (Sl#1) | 100% | Layer 1: Rate | 🔴 Negative Adj (-₹30) |
| **12** | Pump House (40%) | 16,950.00 | **Row 54** | Pump House - Lintel (Sl#14) | 40% | Layer 2: Ratio 0.40 | 🔴 Negative Adj (-₹32,536) |
| **13** | Pump House (60%) | 16,950.00 | **Row 55** | Pump House - Remaining (Sl#14) | 60% | Layer 2: Ratio 0.60 | 🟢 Match (+₹2,44,080) |
| **14** | Heavy Duty Iron Gate | 9,500.00 | **Row 60** | Iron Gate (Sl#21) | 100% | Layer 1: Rate | 🔴 Negative Adj (-₹12,825) |
| **15** | 110mm HDPE Pipe (Supply) | 610.00 | **Row 32** | 110mm HDPE Pipe - Supply | 20% | Layer 2: Ratio 0.20 | 🟢 Match (+₹21,960) |
| **16** | 90mm HDPE Pipe (Supply) | 400.00 | **Row 36** | 90mm HDPE Pipe - Supply | 20% | Layer 2: Ratio 0.20 | 🟢 Match (+₹38,400) |
| **17** | 75mm HDPE Pipe (Supply) | 330.00 | **Row 40** | 75mm HDPE Pipe - Supply | 20% | Layer 2: Ratio 0.20 | 🟢 Match (+₹94,710) |
| **18** | 63mm HDPE Pipe (Supply) | 275.00 | **Row 44** | 63mm HDPE Pipe - Supply | 20% | Layer 2: Ratio 0.20 | 🟢 Match (+₹1,95,250) |
| **19** | 110mm HDPE Pipe (Laying) | 610.00 | **Row 33** | 110mm HDPE Pipe - Laying | 50% | Layer 2: Ratio 0.50 | 🟢 Match (+₹54,900) |
| **20** | 90mm HDPE Pipe (Laying) | 400.00 | **Row 37** | 90mm HDPE Pipe - Laying | 50% | Layer 2: Ratio 0.50 | 🟢 Match (+₹96,000) |
| **21** | 75mm HDPE Pipe (Laying) | 330.00 | **Row 41** | 75mm HDPE Pipe - Laying | 50% | Layer 2: Ratio 0.50 | 🟢 Match (+₹2,36,775) |
| **22** | 63mm HDPE Pipe (Laying) | 275.00 | **Row 45** | 63mm HDPE Pipe - Laying | 50% | Layer 2: Ratio 0.50 | 🟢 Match (+₹4,88,125) |
| **23** | APDCL Power Connection | 9,07,591.26 | **Row 62** | Power Connection (Sl#23) | 100% | Layer 1: Rate | 🟢 Match (+₹9,07,591.26) |
| **24** | Boundary Wall | 5,399.00 | **Row 58** | Boundary Wall (Sl#18) | 100% | Layer 1: Rate | 🔴 Negative Adj (-₹63,018.60) |
| **25** | Toilet Construction | 44,570.00 | **Row 66** | Toilet Construction (Sl#26) | 100% | Layer 1: Rate | 🔴 Negative Adj (-₹7,107) |
| **26** | Electrification Pump House | 35,418.00 | **Row 57** | Electrification (Sl#15) | 100% | Layer 1: Rate | 🟢 Match (+₹35,418) |

---

## Bill Cycle Data Handling Matrix

| Column / Field | Changes Each Bill Cycle? | Source | How Handled by System |
|:---|:---|:---|:---|
| Item Details (Col A-C), Contract Qty (D), Rate (E) | **Never** | TS Estimate | Read-only from template |
| Contract Amount (F) | **Never** | `=D*E` | Preserved formula |
| Milestone Rate (G), Remarks (N), Milestone % (O) | **Never** | Guidelines | Read-only from template |
| **Qty Previous Bill (H)** | **Each Cycle** | Previous Bill Col J | **Populated by tool** (0 if 1st RA) |
| **Qty This Bill (I)** | **Each Cycle** | E-Bill / Site MB | **Populated by tool** from calculated amounts |
| Qty Upto Date (J) | Auto | `=H+I` | Preserved formula |
| Amount Previous (K) | Auto | `=G*H` (or hardcoded for over-billing) | Preserved formula / carryover |
| Amount This Bill (L) | Auto | `=I*G` (or `=F-K` for reconciliation) | Preserved formula |
| Amount Upto Date (M) | Auto | `=K+L` | Preserved formula |
| Totals (Row 75) & Progress % (Row 80-81) | Auto | `=SUM(...)`, `=M75/F76` | Preserved formula |

---

## Negative Reconciliation Handling Rules

For items where the contractor was overpaid in a prior billing cycle (e.g., E-Bill items 11, 12, 14, 24, 25):
1. **Physical Qty (Col I)** remains `0` because the physical work was already executed in full.
2. **Financial Adjustment (Col L)** is achieved either through:
   - Setting `Col L = =F - K` (reconciling contract amount against prior payment), or
   - Setting a negative incremental quantity in Col I if the template uses pure `=I*G`.
3. In the UI, these 5 items are clearly badged with 🔴 **Negative Adjustment** so the engineer sees both the previous overpayment and the current deduction.

---

## Deliverables Summary

| File | Purpose |
|:---|:---|
| `ebill_parser.py` | Extracts 26 items + metadata from E-Bill PDF |
| `item_mapper.py` | 3-layer matching algorithm (Rate &rarr; Ratio &rarr; Keyword) |
| `excel_writer.py` | Populates Columns H & I while preserving all formulas |
| `utils.py` | Currency format, number-to-words, validation |
| `app.py` | Interactive Streamlit web interface |
| `run.bat` | 1-click Windows desktop launcher |
| `requirements.txt` | Complete python dependencies list |
