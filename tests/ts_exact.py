"""
ts_exact.py - Scores a TS reading against the typed answer key (tests/data/ts_truth.json).

    python tests/ts_exact.py            # all documents in the key
    python tests/ts_exact.py Nalia      # documents whose name contains "Nalia"

Every printed cell is compared exactly: Sl, Name of Work (spaces collapsed), Unit, Quantity, Rate,
Amount, plus the Sub Total / platform / contingencies / Grand Total lines. Rows are aligned in order,
so a missing or extra row is reported once instead of shifting every later row.
"""

import json
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional

ROOT = Path(__file__).resolve().parents[1]
TRUTH = ROOT / "tests" / "data" / "ts_truth.json"


def load_truth() -> Dict[str, dict]:
    data = json.loads(TRUTH.read_text(encoding="utf-8"))
    return {k: v for k, v in data.items() if not k.startswith("_")}


def norm_text(s: Optional[str]) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _num_eq(truth, got) -> bool:
    if truth is None:
        return not got
    return got is not None and abs(float(truth) - float(got)) < 0.005


def _variants(desc) -> List[str]:
    return [norm_text(d) for d in (desc if isinstance(desc, list) else [desc])]


def _sim(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def align(truth_rows: List[list], got: List[dict]):
    """Needleman-Wunsch on (text similarity + equal amount). Returns [(truth_index|None, got_index|None)]."""
    n, m = len(truth_rows), len(got)

    def score(i, j):
        t, g = truth_rows[i], got[j]
        s = max(_sim(v.lower(), norm_text(g["desc"]).lower()) for v in _variants(t[1]))
        if t[5] is not None and g["amount"] and _num_eq(t[5], g["amount"]):
            s += 0.6
        return s - 0.45                       # unrelated rows score below a gap
    gap = -0.2
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        dp[i][0] = i * gap
    for j in range(1, m + 1):
        dp[0][j] = j * gap
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            dp[i][j] = max(dp[i - 1][j - 1] + score(i - 1, j - 1), dp[i - 1][j] + gap, dp[i][j - 1] + gap)
    out, i, j = [], n, m
    while i > 0 or j > 0:
        if i > 0 and j > 0 and abs(dp[i][j] - (dp[i - 1][j - 1] + score(i - 1, j - 1))) < 1e-9:
            out.append((i - 1, j - 1)); i -= 1; j -= 1
        elif i > 0 and abs(dp[i][j] - (dp[i - 1][j] + gap)) < 1e-9:
            out.append((i - 1, None)); i -= 1
        else:
            out.append((None, j - 1)); j -= 1
    return out[::-1]


def compare(ts, truth: dict) -> dict:
    """ts: TSParseResult. Returns counts and a list of differences."""
    got = [dict(sl=it.sl, desc=it.desc, unit=it.unit, qty=it.qty, rate=it.rate, amount=it.amount)
           for it in ts.items]
    rows = truth["rows"]
    pairs = align(rows, got)
    c = dict(rows=len(rows), rows_found=0, numbers=0, numbers_ok=0, units=0, units_ok=0, sl=0, sl_ok=0,
             desc=0, desc_exact=0, chars=0, char_errors=0, extra_rows=0)
    diffs: List[str] = []
    for ti, gi in pairs:
        if ti is None:
            c["extra_rows"] += 1
            diffs.append(f"EXTRA row read: {got[gi]['desc'][:70]!r}")
            continue
        t = rows[ti]
        label = f"Sl {t[0] or '-'} {_variants(t[1])[0][:40]!r}"
        if gi is None:
            diffs.append(f"MISSING row: {label}")
            c["numbers"] += sum(v is not None for v in t[3:6])
            c["desc"] += 1
            c["chars"] += len(_variants(t[1])[0])
            c["char_errors"] += len(_variants(t[1])[0])
            continue
        g = got[gi]
        c["rows_found"] += 1
        for k, tv in zip(("qty", "rate", "amount"), t[3:6]):
            if tv is None and not g[k]:
                continue
            c["numbers"] += 1
            if _num_eq(tv, g[k]):
                c["numbers_ok"] += 1
            else:
                diffs.append(f"{label}: {k} {g[k]!r} should be {tv!r}")
        if t[2]:
            c["units"] += 1
            if norm_text(g["unit"]) == t[2]:
                c["units_ok"] += 1
            else:
                diffs.append(f"{label}: unit {g['unit']!r} should be {t[2]!r}")
        if t[0]:
            c["sl"] += 1
            if g["sl"] == t[0]:
                c["sl_ok"] += 1
            else:
                diffs.append(f"{label}: Sl {g['sl']!r} should be {t[0]!r}")
        c["desc"] += 1
        text = norm_text(g["desc"])
        best = max(_variants(t[1]), key=lambda v: _sim(v, text))
        c["chars"] += len(best)
        if text in _variants(t[1]):
            c["desc_exact"] += 1
        else:
            sm = SequenceMatcher(None, best, text)
            errs = sum(max(i2 - i1, j2 - j1) for op, i1, i2, j1, j2 in sm.get_opcodes() if op != "equal")
            c["char_errors"] += errs
            diffs.append(f"{label}: text {text!r}")
    s = truth["summary"]
    meta = ts.metadata
    for k in ("sub_total", "beneficiary_share", "contingencies", "grand_total"):
        c["numbers"] += 1
        if _num_eq(s[k], getattr(meta, k)):
            c["numbers_ok"] += 1
        else:
            diffs.append(f"{k}: {getattr(meta, k)!r} should be {s[k]!r}")
    c["pages_ok"] = [p + 1 for p in ts.pages] == truth["pages"]
    if not c["pages_ok"]:
        diffs.insert(0, f"pages read {[p + 1 for p in ts.pages]} should be {truth['pages']}")
    c["scheme_ok"] = norm_text(meta.scheme).upper() == truth["scheme"].upper()
    if not c["scheme_ok"]:
        diffs.insert(0, f"scheme {meta.scheme!r} should be {truth['scheme']!r}")
    c["diffs"] = diffs
    return c


def summary_line(name: str, c: dict) -> str:
    return (f"{name[:34]:34} pages {'ok' if c['pages_ok'] else 'X '} scheme {'ok' if c['scheme_ok'] else 'X '} "
            f"rows {c['rows_found']}/{c['rows']} (+{c['extra_rows']}) numbers {c['numbers_ok']}/{c['numbers']} "
            f"units {c['units_ok']}/{c['units']} Sl {c['sl_ok']}/{c['sl']} text exact {c['desc_exact']}/{c['desc']} "
            f"chars {100 * (1 - c['char_errors'] / max(1, c['chars'])):.1f}%")


def main(argv: List[str]) -> int:
    sys.path.insert(0, str(ROOT))
    sys.stdout.reconfigure(encoding="utf-8")
    import table_extractor
    table_extractor._get_rapid()
    from ts_parser import parse_ts_pdf
    pattern = argv[0].lower() if argv else ""
    total = dict(numbers=0, numbers_ok=0, desc=0, desc_exact=0, chars=0, char_errors=0)
    lines = []
    for name, truth in load_truth().items():
        if pattern not in name.lower():
            continue
        ts = parse_ts_pdf((ROOT / truth["path"]).read_bytes())
        c = compare(ts, truth)
        for k in total:
            total[k] += c[k]
        lines.append(summary_line(name, c))
        print(lines[-1])
        for d in c["diffs"]:
            print("    " + d)
    print("\n".join(["", *lines]))
    print(f"TOTAL numbers {total['numbers_ok']}/{total['numbers']}  text exact {total['desc_exact']}/{total['desc']}  "
          f"chars {100 * (1 - total['char_errors'] / max(1, total['chars'])):.2f}%")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
