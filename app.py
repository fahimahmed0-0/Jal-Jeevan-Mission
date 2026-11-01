"""
app.py - Streamlit interface for the JJM PHED Assam TPIA Abstract of Bill (Payment Recommendation).

Flow: upload documents -> Process -> review (every check shown with its real numbers, uncertain
values highlighted, lines editable) -> download an Excel file that has been recalculated and
verified against the reviewed figures.
"""

import copy
import datetime
import hashlib
import os
import re
import time
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from abstract_model import Abstract, Check
from excel_writer import generate_workbook, verify_workbook
from item_mapper import rebuild_template_abstract
from pipeline import PipelineResult, build_abstract, parse_ts
from table_extractor import ocr_backend_name
from ts_ebill_builder import LAYOUTS
from ts_parser import parse_page_spec
from utils import amount_to_words, format_inr

APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "output"

st.set_page_config(page_title="JJM PHED Assam | TPIA Abstract of Bill", page_icon="💧", layout="wide")

st.markdown("""
<style>
  .metric-card {border:1px solid rgba(128,128,128,.35); border-radius:8px; padding:.8rem 1rem;}
  .metric-label {font-size:.78rem; font-weight:600; opacity:.75; text-transform:uppercase; letter-spacing:.04em;}
  .metric-value {font-size:1.45rem; font-weight:700;}
</style>
""", unsafe_allow_html=True)

st.title("💧 JJM PHED Assam — TPIA Abstract of Bill")
st.caption("Contractor bill ➜ verified Payment Recommendation workbook. Every figure is checked against the "
           "documents' own arithmetic before the Excel file can be downloaded.")


# ---------------------------------------------------------------------------
# Processing (cached by file content, so re-runs and reviews never re-OCR)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False, max_entries=8)
def _process(bill: Optional[bytes], ts: Optional[bytes], template: Optional[bytes], ts_pages: str = "",
             layout: str = "ts") -> PipelineResult:
    if bill is not None and bill == ts:
        bill = None                   # the TS was put in both boxes: there is no e-bill
    pages = parse_page_spec(ts_pages)
    return build_abstract(bill, parse_ts(ts, pages), template, ts_pages=pages, layout=layout)


def _sig(*parts: Optional[bytes]) -> str:
    h = hashlib.sha256()
    for p in parts:
        h.update(p or b"-")
        h.update(b"|")
    return h.hexdigest()


ss = st.session_state
ss.setdefault("uploader_gen", 0)
ss.setdefault("inputs", {"bill": None, "ts": None, "template": None, "names": {}})
ss.setdefault("result", None)
ss.setdefault("result_sig", None)


def _load_sample(bill: str, ts: Optional[str], template: Optional[str]):
    ss.inputs = {
        "bill": (APP_DIR / bill).read_bytes(),
        "ts": (APP_DIR / ts).read_bytes() if ts else None,
        "template": (APP_DIR / template).read_bytes() if template else None,
        "names": {"bill": Path(bill).name, "ts": Path(ts).name if ts else None,
                  "template": Path(template).name if template else None},
    }
    ss.uploader_gen += 1          # clears the upload widgets so they cannot override the sample
    ss.run_now = True


with st.sidebar:
    st.subheader("Sample documents")
    ts_docs = sorted((APP_DIR / "Doc").glob("DocScanner*.pdf")) if (APP_DIR / "Doc").is_dir() else []
    samples = {
        "Bengsatra — e-bill + TS document":
            ("Doc/E-Bill__ PHED, ASSAM bengsatra.pdf",
             str(ts_docs[0].relative_to(APP_DIR)) if ts_docs else "Doc/missing-ts.pdf", None),
        "Bengsatra — e-bill + previous abstract (Excel)":
            ("Doc/E-Bill__ PHED, ASSAM bengsatra.pdf", None, "Doc/BISONDOI PT-II BANG SATRA.xlsx"),
        "Tolli PT-I — scanned abstract + scanned TS":
            ("Doc/test/E-bill.pdf", "Doc/test/ts.pdf", None),
    }
    for label, files in samples.items():
        if all((APP_DIR / f).exists() for f in files if f):
            st.button(label, on_click=_load_sample, args=files, width="stretch")
    st.divider()
    st.caption(f"OCR engine: {ocr_backend_name()}")
    st.caption("Scanned PDFs are read cell-by-cell from the ruled table; every number is cross-checked with "
               "the document's own totals. Values that cannot be confirmed are highlighted for review.")


# ---------------------------------------------------------------------------
# Intake
# ---------------------------------------------------------------------------

st.subheader("1. Documents")
c1, c2, c3 = st.columns(3)
gen = ss.uploader_gen
with c1:
    up_bill = st.file_uploader("Contractor e-bill (PDF) — optional with a TS", type=["pdf"], key=f"bill_{gen}",
                               help="Digital portal e-bill (Running Account Bill-C) or a scanned TPIA abstract. "
                                    "Without it, the abstract is laid out from the TS with quantities at zero.")
with c2:
    up_ts = st.file_uploader("Technical Sanction document (PDF) — needed for the official layout", type=["pdf"],
                             key=f"ts_{gen}",
                             help="The whole TS document may be uploaded: the 'Abstract of Cost Estimate for Revised "
                                  "Technical Sanction' pages are found automatically. Items, quantities, rates and "
                                  "the Sub Total come from that table.")
    ts_pages = st.text_input("TS abstract pages (optional)", value=ss.get("ts_pages", ""), key=f"tsp_{gen}",
                             placeholder="auto - e.g. 9-10",
                             help="Leave empty to find the abstract table automatically.")
    ss.ts_pages = ts_pages
    layout_names = list(LAYOUTS)
    ss.layout = st.radio("Abstract layout", layout_names, format_func=LAYOUTS.get,
                         index=layout_names.index(ss.get("layout", "ts")), key=f"layout_{gen}",
                         help="As in the TS: every TS line copied in its own order with its own Sl numbers "
                              "and wording. Standard: items regrouped in the standard TPIA order.")
with c3:
    up_tpl = st.file_uploader("Previous TPIA abstract (Excel) — optional", type=["xlsx"], key=f"tpl_{gen}",
                              help="Only used when no TS is given: the new bill is then written into this workbook, "
                                   "keeping its formulas and formatting.")
uploaded = ss.inputs.setdefault("uploaded", set())
for key, up in (("bill", up_bill), ("ts", up_ts), ("template", up_tpl)):
    if up is not None:
        ss.inputs[key] = up.getvalue()
        ss.inputs["names"][key] = up.name
        uploaded.add(key)
    elif key in uploaded:             # file removed from the upload box (x): forget it
        ss.inputs[key] = None
        ss.inputs["names"][key] = None
        uploaded.discard(key)

names = ss.inputs["names"]
st.caption("Loaded: " + " · ".join(f"**{k}**: {v}" for k, v in names.items() if v and ss.inputs.get(k)) if any(
    ss.inputs.get(k) for k in ("bill", "ts", "template")) else "No documents loaded yet.")

try:
    parse_page_spec(ss.ts_pages)
except ValueError as ex:
    st.error(str(ex))
    st.stop()
cur_sig = _sig(ss.inputs["bill"], ss.inputs["ts"], ss.inputs["template"], ss.ts_pages.encode(), ss.layout.encode())
cols = st.columns([1, 1, 4])
run = cols[0].button("🚀 Process", type="primary", width="stretch",
                     disabled=ss.inputs["bill"] is None and ss.inputs["ts"] is None)
if cols[1].button("Clear", width="stretch"):
    ss.inputs = {"bill": None, "ts": None, "template": None, "names": {}}
    ss.result, ss.result_sig = None, None
    ss.uploader_gen += 1
    st.rerun()

if run or ss.pop("run_now", False):
    t0 = time.time()
    with st.spinner("Reading documents (scanned pages take ~10 s each)…"):
        try:
            res = _process(ss.inputs["bill"], ss.inputs["ts"], ss.inputs["template"], ss.ts_pages, ss.layout)
            ss.result = copy.deepcopy(res)          # edits must never touch the cached copy
            ss.result_sig = cur_sig
            ss.elapsed = time.time() - t0
        except Exception as ex:
            ss.result, ss.result_sig = None, None
            st.error(f"Could not process the documents: {ex}")

if ss.result is not None and ss.result_sig != cur_sig:
    st.warning("The documents changed since the last run — click **Process** to update the results.")
    st.stop()
if ss.result is None:
    st.stop()

res: PipelineResult = ss.result
A: Abstract = res.abstract
mode = A.source_kind
mode_label = {"ts_ebill": "TS abstract (items, rates) + contractor e-bill (previous / this bill)",
              "ts_only": "TS abstract only — enter previous / this-bill quantities in the table",
              "scanned_abstract": "Scanned TPIA abstract (grid OCR)",
              "ebill_template": "Digital e-bill written into the previous abstract (template)",
              "digital_ebill": "Digital e-bill (official layout generated)"}[mode]
st.caption(f"Mode: **{mode_label}** · processed in {ss.get('elapsed', 0):.1f} s")


# ---------------------------------------------------------------------------
# Header data (editable)
# ---------------------------------------------------------------------------

st.subheader("2. Scheme details")
with st.expander("Scheme / contractor / work order", expanded=False):
    m = A.meta
    a1, a2 = st.columns(2)
    m.scheme = a1.text_input("Name of scheme", m.scheme)
    m.contractor = a2.text_input("Name of contractor", m.contractor)
    m.work_order_no = a1.text_input("Work order no.", m.work_order_no)
    m.contract_value_text = a2.text_input("Contract value (as printed)", m.contract_value_text)
    m.division = a1.text_input("Division", m.division)
    m.sub_division = a2.text_input("Sub-division", m.sub_division)
    m.status_text = a1.text_input("Completion date / status", m.status_text)
    m.bill_ref = a2.text_input("Bill reference", m.bill_ref)
    wo = st.number_input("Work order amount (₹) — used for physical / financial progress",
                         min_value=0.0, value=float(A.work_order_amount or 0.0), step=1.0, format="%.2f")
    A.work_order_amount = wo or None


# ---------------------------------------------------------------------------
# Status + metrics
# ---------------------------------------------------------------------------

checks = A.run_checks() + res.benchmark_checks
blocking = [c for c in checks if not c.passed and c.level == "critical"]
warnings = [c for c in checks if not c.passed and c.level == "warning"]

st.subheader("3. Review")
if not blocking:
    st.success(f"✅ **All critical checks passed** ({sum(c.passed for c in checks)}/{len(checks)} checks OK"
               + (f", {len(warnings)} warning(s)" if warnings else "") + ").")
else:
    st.error("❌ **Not verified** — " + "; ".join(c.name for c in blocking) + ". See the Checks tab and the highlighted lines.")

def card(col, label, value, sub=""):
    col.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div>'
                 f'<div class="metric-value">{value}</div><div style="font-size:.8rem;opacity:.75">{sub}</div></div>',
                 unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
card(k1, "Contract amount", format_inr(A.total_contract))
card(k2, "Paid upto previous bill", format_inr(A.total_prev))
card(k3, "This bill", format_inr(A.total_this),
     ("claimed " + format_inr(A.stated["claim_net"])) if "claim_net" in A.stated else "")
card(k4, "Payment recommended", format_inr(A.recommended), "ROUNDUP(this bill)")
card(k5, "Progress (upto date)", f"{A.progress:.2%}" if A.progress is not None else "—",
     "of work order amount" if A.progress is not None else "enter work order amount")
st.info(f"**In words:** {amount_to_words(A.recommended)}")

tab_lines, tab_checks, tab_ts, tab_notes = st.tabs(["📊 Abstract lines", "✅ Checks", "📋 Technical Sanction", "📝 Reading log"])

STATUS_ICON = {"ok": "🟢 ok", "corrected": "🟡 corrected", "review": "🔴 review", "edited": "✏️ edited"}

with tab_lines:
    editable = {"scanned_abstract": ["Qty", "Rate", "Rate paid", "Prev qty", "This qty", "Prev amt", "This amt"],
                "digital_ebill": ["This qty", "This amt"],
                "ts_ebill": ["This qty", "This amt"],
                "ts_only": ["Prev qty", "This qty"],
                "ebill_template": ["This qty"]}[mode]
    text_editable = ["Description", "Unit"] if mode in ("ts_ebill", "ts_only") else []
    st.caption("Editable columns: " + ", ".join(text_editable + editable) + ". Upto-date values recompute "
               "automatically. 🔴 lines could not be confirmed from the document — check them against the "
               "original. A note 'Check the wording' marks words the two OCR readings disagreed on.")
    show_only_review = st.toggle("Show only lines that need attention", value=False)
    recs = []
    for rid, it, r in A.row_ids():
        if show_only_review and r.status not in ("review", "corrected", "edited"):
            continue
        first = it.rows[0] is r
        recs.append({
            "_id": rid, "Sl": it.sl if first else "", "Description": it.desc if first else "",
            "Unit": it.unit if first else "", "Qty": it.qty, "Rate": it.rate, "Milestone": r.remarks,
            "%": f"{r.pct * 100:g}%" if r.pct is not None else "", "Rate paid": r.rate_paid,
            "Prev qty": r.prev_qty, "This qty": r.this_qty, "Upto qty": r.upto_qty,
            "Prev amt": round(r.prev_amt, 2), "This amt": round(r.this_amt, 2), "Upto amt": round(r.upto_amt, 2),
            "Status": STATUS_ICON.get(r.status, r.status),
            "E-bill line": r.ebill_ref if r.ebill_ref is not None else "",
            "Notes": " | ".join((it.notes if first else []) + r.notes),
        })
    df = pd.DataFrame(recs)
    if df.empty:
        st.write("No lines to show.")
    else:
        money = st.column_config.NumberColumn(format="%.2f")
        qty = st.column_config.NumberColumn(format="%.3f")
        edited_df = st.data_editor(
            df, hide_index=True, width="stretch", height=520, key=f"editor_{ss.result_sig}_{show_only_review}",
            column_order=[c for c in df.columns if c != "_id"],
            disabled=[c for c in df.columns if c not in editable + text_editable],
            column_config={"Description": st.column_config.TextColumn(width="large"),
                           "Milestone": st.column_config.TextColumn(width="medium"),
                           "Notes": st.column_config.TextColumn(width="large"),
                           "Qty": qty, "Prev qty": qty, "This qty": qty, "Upto qty": qty,
                           "Rate": money, "Rate paid": money, "Prev amt": money, "This amt": money, "Upto amt": money})
        field_of = {"Qty": "qty", "Rate": "rate", "Rate paid": "rate_paid", "Prev qty": "prev_qty",
                    "This qty": "this_qty", "Prev amt": "prev_amt", "This amt": "this_amt"}
        changes = []
        text_changes = []
        for (_, old), (_, new) in zip(df.iterrows(), edited_df.iterrows()):
            for col in text_editable:
                if old["Description"] == "" and old["Unit"] == "":
                    continue                      # milestone row under an item: the text lives on its first row
                if not pd.isna(new[col]) and str(new[col]).strip() != str(old[col]).strip():
                    text_changes.append((old["_id"], {"Description": "desc", "Unit": "unit"}[col], str(new[col])))
            for col in editable:
                nv = new[col]
                if pd.isna(nv):
                    continue                      # a cleared cell keeps its value
                if abs(float(nv) - float(old[col])) > 1e-9:
                    changes.append((old["_id"], field_of[col], float(nv)))
        for rid, f, text in text_changes:
            A.apply_text_edit(rid, f, text)
        if text_changes and not changes:
            st.rerun()
        if changes:
            if mode == "ebill_template":
                edits = {}
                for rid, f, v in changes:
                    _, r = A.row_by_id(rid)
                    if r.l_fixed:
                        st.warning(f"Row {r.excel_row}: this line's amount is fixed by a reconciliation formula in "
                                   "the template — its quantity does not change the payment.")
                    else:
                        edits[r.excel_row] = v
                if edits:
                    res.abstract = rebuild_template_abstract(A, res.ebill, edits)
            else:
                for rid, f, v in changes:
                    A.apply_edit(rid, f, v)
            st.rerun()

with tab_checks:
    for c in checks:
        st.markdown(f"{c.icon} **{c.name}** — {c.detail}")
    for n in A.notes:
        st.caption(n)

with tab_ts:
    if res.ts is None:
        st.info("No Technical Sanction uploaded.")
    else:
        t = res.ts
        st.markdown(f"**{t.metadata.scheme or 'Technical Sanction'}** — pages "
                    f"{', '.join(str(p + 1) for p in t.pages) or 'all'} · Sub Total "
                    f"{format_inr(t.metadata.sub_total or t.total_sanctioned_amount)}")
        for c in t.checks:
            st.markdown(f"{c.icon} **{c.name}** — {c.detail}")
        if res.benchmark_rows:
            st.markdown("**Abstract vs TS**")
            st.dataframe(pd.DataFrame(res.benchmark_rows), hide_index=True, width="stretch")
        with st.expander("TS items as read", expanded=mode == "ts_ebill"):
            st.dataframe(pd.DataFrame([{"Sl": i.sl, "Item": i.desc, "Qty": i.qty, "Rate": i.rate, "Unit": i.unit,
                                        "Amount": i.amount, "Status": STATUS_ICON.get(i.status, i.status),
                                        "Notes": " | ".join(i.notes)} for i in t.items if not i.heading]),
                         hide_index=True, width="stretch")

with tab_notes:
    logged = [(rid, it, r) for rid, it, r in A.row_ids() if r.notes or r.status != "ok"]
    if not logged:
        st.write("Every value was read directly and confirmed — nothing to report.")
    for rid, it, r in logged:
        st.markdown(f"{STATUS_ICON.get(r.status, r.status)} **Sl {it.sl or '—'} {it.desc[:60]}** — {r.remarks[:50]} "
                    f"({r.source}): " + "; ".join(r.notes))


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

st.subheader("4. Payment Recommendation workbook")
try:
    xlsx, rowmap = generate_workbook(A)
    xl_checks = verify_workbook(xlsx, A, rowmap)
except Exception as ex:
    xlsx, xl_checks = None, [Check("Workbook generated", False, f"{ex}")]
xl_failed = [c for c in xl_checks if not c.passed]

with st.expander("Excel recalculation check", expanded=bool(xl_failed)):
    for c in xl_checks:
        st.markdown(f"{c.icon} **{c.name}** — {c.detail}")

if xlsx is None or xl_failed:
    st.error("The generated workbook does not recalculate to the reviewed figures, so it is not offered for "
             "download. Please report this with the documents used.")
    st.stop()

allowed = not blocking
if blocking:
    allowed = st.checkbox("I have checked the flagged lines against the original documents and accept the figures "
                          "shown above.", value=False)

scheme_clean = re.sub(r"[^A-Za-z0-9]+", "_", A.meta.scheme or "Scheme").strip("_")[:40] or "Scheme"
bill_clean = re.sub(r"[^A-Za-z0-9]+", "_", A.meta.bill_ref or "").strip("_")
file_name = f"TPIA_Abstract_{scheme_clean}{'_' + bill_clean if bill_clean else ''}.xlsx"

d1, d2, d3 = st.columns(3)
d1.download_button("📥 Download Excel", data=xlsx, file_name=file_name, type="primary", disabled=not allowed,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
if d2.button("💾 Save to project 'output' folder", disabled=not allowed, width="stretch"):
    OUTPUT_DIR.mkdir(exist_ok=True)
    target = OUTPUT_DIR / file_name
    try:
        target.write_bytes(xlsx)
    except PermissionError:                     # the file is open in Excel
        target = OUTPUT_DIR / f"{target.stem}_{datetime.datetime.now():%H%M%S}.xlsx"
        target.write_bytes(xlsx)
    ss.saved_path = str(target)
if ss.get("saved_path"):
    d2.caption(f"Saved: `{ss.saved_path}`")
    if os.name == "nt" and d3.button("📊 Open saved file in Excel", width="stretch"):
        os.startfile(ss.saved_path)
st.caption(f"Recommended in this bill: **{format_inr(A.recommended)}** — {amount_to_words(A.recommended)}")
