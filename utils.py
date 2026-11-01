"""
utils.py - Number helpers for Jal Jeevan Mission (JJM) TPIA billing.

  * Indian currency formatting (lakh / crore grouping)
  * Amount-to-words in the Indian numbering system
  * Robust number parsing for PDF text and OCR output, including the alternative readings an OCR
    engine typically produces (dropped decimal point, O/0 and l/1 confusions, space as decimal).
"""

import math
import re
from typing import List, Optional, Tuple, Union

Number = Union[int, float]


def format_inr(amount: Optional[Number], include_symbol: bool = True) -> str:
    """
    Formats a number into the Indian numbering system (Lakhs & Crores).
        2579768.66  -> '₹ 25,79,768.66'
        10353628.38 -> '₹ 1,03,53,628.38'
        -115516.60  -> '-₹ 1,15,516.60'
    """
    if amount is None:
        amount = 0.0
    is_negative = amount < 0 and round(abs(amount), 2) != 0
    integer_part, decimal_part = f"{abs(amount):.2f}".split(".")
    if len(integer_part) > 3:
        last_three = integer_part[-3:]
        remaining = integer_part[:-3]
        groups = []
        while len(remaining) > 2:
            groups.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            groups.insert(0, remaining)
        integer_part = ",".join(groups) + "," + last_three
    prefix = "₹ " if include_symbol else ""
    text = f"{prefix}{integer_part}.{decimal_part}"
    return f"-{text}" if is_negative else text


_ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen"
]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _two_digits_to_words(n: int) -> str:
    if n < 20:
        return _ONES[n]
    tens, ones = _TENS[n // 10], _ONES[n % 10]
    return f"{tens} {ones}" if ones else tens


def _three_digits_to_words(n: int) -> str:
    hundreds, rest = divmod(n, 100)
    parts = []
    if hundreds:
        parts.append(f"{_ONES[hundreds]} Hundred")
    if rest:
        parts.append(_two_digits_to_words(rest))
    return " ".join(parts)


def _integer_to_words(n: int) -> str:
    """Indian system words for a non-negative integer (crore groups recurse for very large values)."""
    if n == 0:
        return "Zero"
    crore, rem = divmod(n, 10_000_000)
    lakh, rem = divmod(rem, 100_000)
    thousand, hundreds = divmod(rem, 1000)
    parts = []
    if crore:
        parts.append(f"{_integer_to_words(crore)} Crore")
    if lakh:
        parts.append(f"{_two_digits_to_words(lakh)} Lakh")
    if thousand:
        parts.append(f"{_two_digits_to_words(thousand)} Thousand")
    if hundreds:
        parts.append(_three_digits_to_words(hundreds))
    return " ".join(parts)


def amount_to_words(amount: Optional[Number]) -> str:
    """
    Converts a currency amount into Indian English words (upper case), e.g.
        4573421     -> 'RUPEES FORTY FIVE LAKH SEVENTY THREE THOUSAND FOUR HUNDRED TWENTY ONE ONLY'
        2579768.66  -> 'RUPEES TWENTY FIVE LAKH ... SIXTY EIGHT AND SIXTY SIX PAISE ONLY'
    """
    if amount is None:
        amount = 0.0
    is_negative = amount < 0
    total_paise = int(round(abs(amount) * 100))          # rounding carries 0.995 -> 1 rupee correctly
    rupees, paise = divmod(total_paise, 100)
    result = f"Rupees {_integer_to_words(rupees)}"
    if paise:
        result += f" and {_two_digits_to_words(paise)} Paise"
    result += " Only"
    if is_negative and total_paise:
        result = f"Minus {result}"
    return result.upper()


# ---------------------------------------------------------------------------
# Number parsing
# ---------------------------------------------------------------------------

_OCR_DIGIT_MAP = str.maketrans({
    "O": "0", "o": "0", "Q": "0", "D": "0", "U": "0", "c": "0", "C": "0",
    "l": "1", "I": "1", "|": "1", "i": "1", "!": "1", "J": "1",
    "S": "5", "s": "5", "B": "8", "Z": "2", "z": "2", "G": "6", "b": "6", "g": "9", "q": "9",
    "m": "00", "n": "0",
})


def clean_number(val: Optional[Union[str, Number]]) -> float:
    """
    Parses a number printed in a digital PDF / spreadsheet: '₹ -32536' -> -32536.0,
    '3,015,488.97' -> 3015488.97, '5,67,139.20' -> 567139.2, '₹ 5595258/-' -> 5595258.0.
    Returns 0.0 for empty / dash cells.
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("−", "-")
    if s.endswith("/-"):
        s = s[:-2]
    neg = s.startswith("-") or bool(re.match(r"^[₹Rr]s?\.?\s*-", s)) or (s.startswith("(") and s.endswith(")"))
    digits = re.sub(r"[^\d.]", "", s)
    if not digits or digits == ".":
        return 0.0
    if digits.count(".") > 1:                       # 5.67.139.20 (dots as group separators)
        head, tail = digits.rsplit(".", 1)
        digits = head.replace(".", "") + "." + tail
    try:
        v = float(digits)
    except ValueError:
        return 0.0
    return -v if neg else v


def number_candidates(text: Optional[str], decimals_expected: Optional[int] = None) -> List[Tuple[float, float]]:
    """
    Plausible numeric readings of an OCR token, as (value, cost) pairs sorted by cost.
    cost 0 = the literal reading; higher costs are increasingly speculative repairs:
        '567168 00' -> 567168.00 (space read instead of the decimal point)
        '4750800'   -> 4750800 and 47508.00 (decimal point dropped)
        'ooo'/'moo' -> 0.00
        '0 00%'     -> 0.00
    decimals_expected: 2 for amounts printed with paise - makes the dropped-decimal repair cheaper.
    """
    if text is None:
        return []
    raw = str(text).strip()
    if not raw:
        return []
    out: List[Tuple[float, float]] = []

    def add(v: float, cost: float):
        if not math.isfinite(v):
            return
        for i, (ov, oc) in enumerate(out):
            if abs(ov - v) < 1e-9:
                if cost < oc:
                    out[i] = (v, cost)
                return
        out.append((v, cost))

    s = raw.replace("₹", "").replace("Rs.", "").replace("Rs", "").strip()
    s = s.rstrip("%").strip().rstrip("/-").strip()
    neg = s.startswith("-")
    s = s.lstrip("-").strip()
    sign = -1.0 if neg else 1.0

    # literal reading (an inner space usually stands for a lost decimal point - see below)
    lit = re.sub(r"[,\s]", "", s)
    if re.fullmatch(r"\d+(\.\d+)?", lit):
        add(sign * float(lit), 0.5 if re.search(r"\d\s+\d", s) else 0.0)
    elif re.fullmatch(r"\d+(\.\d+){2,}", lit):         # 5.67.139.20
        head, tail = lit.rsplit(".", 1)
        add(sign * float(head.replace(".", "") + "." + tail), 0.1)

    # space (or comma) used as the decimal separator: '674549 50', '0 00'
    m = re.fullmatch(r"(\d[\d,]*)[\s,](\d{2})", s)
    if m:
        add(sign * float(m.group(1).replace(",", "") + "." + m.group(2)), 0.1 if " " in s else 0.2)

    # OCR letter/digit confusions
    mapped = s.translate(_OCR_DIGIT_MAP)
    mapped = re.sub(r"[,\s]", "", mapped)
    if mapped != lit and re.fullmatch(r"\d+(\.\d+)?", mapped):
        add(sign * float(mapped), 0.6)
    m2 = re.fullmatch(r"(\d+)\s*(\d{2})", s.translate(_OCR_DIGIT_MAP).replace(",", ""))
    if m2 and " " in s:
        add(sign * float(m2.group(1) + "." + m2.group(2)), 0.6)

    # dropped decimal point: '4750800' -> 47508.00 (only for digit strings ending in two digits)
    for base in {lit, mapped}:
        if re.fullmatch(r"\d{3,}", base or ""):
            add(sign * float(base[:-2] + "." + base[-2:]), 0.3 if decimals_expected == 2 else 0.9)
            if base.endswith("0") and len(base) >= 3:
                add(sign * float(base[:-1] + "." + base[-1:]), 1.2)   # one-decimal style '405640.0'

    # stray symbols from a neighbouring cell ('30130.00 =', ': 4,05,640.00'): read the number itself
    nums = re.findall(r"\d[\d,]*(?:\.\d+)?", s)
    if len(nums) == 1 and nums[0] != s and re.fullmatch(r"[^\w]*", s.replace(nums[0], "", 1)):
        add(sign * float(nums[0].replace(",", "")), 0.05)     # only punctuation around it ('1o' stays 10)

    # all-zero garbage ('ooo', 'coo', 'moo', 'o.co')
    if re.fullmatch(r"[0oOcCmQ.\s]+", s) and not out:
        add(0.0, 0.4)

    out.sort(key=lambda p: p[1])
    return out


def best_number(text: Optional[str], default: float = 0.0, decimals_expected: Optional[int] = None) -> float:
    c = number_candidates(text, decimals_expected)
    return c[0][0] if c else default


def money_eq(a: Optional[float], b: Optional[float], tol: float = 0.005) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def round2(x: float) -> float:
    return float(f"{x + 0.0:.2f}")


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    for v in (2579768.66, 10353628.38, 5595258.00, -115516.60, 405640.00, 50.00, 0.995, 4573421):
        print(f"  {v:>15.3f} -> {format_inr(v, include_symbol=False):>18s} | {amount_to_words(v)}")
