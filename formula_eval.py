"""
formula_eval.py - Minimal Excel formula evaluator used to *prove* generated workbooks.

openpyxl writes formulas but cannot calculate them. Before a workbook is offered for download the
app recalculates it here and compares every total with the verified numbers it was built from.
Supports the subset used by TPIA abstracts: arithmetic, percent literals, cell references, ranges
and SUM / ROUND / ROUNDUP / ROUNDDOWN / ABS / MIN / MAX / IF / IFERROR / AND / OR.
"""

import math
import re
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.utils import column_index_from_string, get_column_letter


class FormulaError(Exception):
    pass


_TOKEN = re.compile(r"""
    (?P<ws>\s+)
  | (?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<range>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)
  | (?P<ref>\$?[A-Z]{1,3}\$?\d+)
  | (?P<func>[A-Z][A-Z0-9.]*(?=\())
  | (?P<bool>TRUE|FALSE)
  | (?P<op><>|<=|>=|[-+*/^&=<>%(),:])
""", re.VERBOSE)


def _tokenize(src: str) -> List[Tuple[str, str]]:
    out, pos = [], 0
    while pos < len(src):
        m = _TOKEN.match(src, pos)
        if not m:
            raise FormulaError(f"Unexpected text at {src[pos:pos + 15]!r}")
        pos = m.end()
        kind = m.lastgroup
        if kind != "ws":
            out.append((kind, m.group(kind)))
    return out


def _excel15(x: float) -> float:
    """Excel works with 15 significant digits; apply that before rounding functions."""
    return float(f"{x:.15g}") if isinstance(x, float) and math.isfinite(x) else x


def _roundup(x: float, n: int) -> float:
    f = 10 ** n
    v = abs(_excel15(x)) * f
    r = math.ceil(_excel15(v)) / f
    return -r if x < 0 else r


def _rounddown(x: float, n: int) -> float:
    f = 10 ** n
    v = abs(_excel15(x)) * f
    r = math.floor(_excel15(v)) / f
    return -r if x < 0 else r


def _round(x: float, n: int) -> float:
    f = 10 ** n
    v = abs(_excel15(x)) * f
    r = math.floor(_excel15(v) + 0.5) / f
    return -r if x < 0 else r


class WorksheetEvaluator:
    """Evaluates cells of one openpyxl worksheet loaded with data_only=False."""

    def __init__(self, ws):
        self.ws = ws
        self._cache: Dict[Tuple[int, int], Any] = {}
        self._stack: set = set()

    # -- public ---------------------------------------------------------
    def value(self, row: int, col: int) -> Any:
        key = (row, col)
        if key in self._cache:
            return self._cache[key]
        if key in self._stack:
            raise FormulaError(f"Circular reference at {get_column_letter(col)}{row}")
        self._stack.add(key)
        try:
            raw = self.ws.cell(row, col).value
            if isinstance(raw, str) and raw.startswith("="):
                val = self._eval_formula(raw[1:])
            else:
                val = raw
        finally:
            self._stack.discard(key)
        self._cache[key] = val
        return val

    def number(self, row: int, col: int) -> float:
        v = self.value(row, col)
        if v is None or v == "":
            return 0.0
        if isinstance(v, bool):
            return float(v)
        if isinstance(v, (int, float)):
            return float(v)
        raise FormulaError(f"{get_column_letter(col)}{row} is text ({v!r}), not a number")

    def coord(self, ref: str) -> Any:
        col, row = self._split(ref)
        return self.value(row, col)

    # -- parser ---------------------------------------------------------
    def _eval_formula(self, src: str) -> Any:
        # a referenced cell may itself hold a formula: keep the outer formula's parse position
        saved = (getattr(self, "_toks", None), getattr(self, "_i", 0))
        try:
            self._toks = _tokenize(src)
            self._i = 0
            val = self._comparison()
            if self._i != len(self._toks):
                raise FormulaError(f"Unparsed formula tail in ={src}")
            return val
        finally:
            self._toks, self._i = saved

    def _peek(self) -> Optional[Tuple[str, str]]:
        return self._toks[self._i] if self._i < len(self._toks) else None

    def _take(self, expected: Optional[str] = None) -> Tuple[str, str]:
        tok = self._peek()
        if tok is None or (expected is not None and tok[1] != expected):
            raise FormulaError(f"Expected {expected!r}")
        self._i += 1
        return tok

    def _comparison(self):
        left = self._concat()
        while self._peek() and self._peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self._take()[1]
            right = self._concat()
            l, r = self._num(left, True), self._num(right, True)
            left = {"=": l == r, "<>": l != r, "<": l < r, ">": l > r, "<=": l <= r, ">=": l >= r}[op]
        return left

    def _concat(self):
        left = self._additive()
        while self._peek() and self._peek()[1] == "&":
            self._take()
            left = f"{'' if left is None else left}{'' if (r := self._additive()) is None else r}"
        return left

    def _additive(self):
        left = self._term()
        while self._peek() and self._peek()[1] in ("+", "-"):
            op = self._take()[1]
            right = self._term()
            left = self._num(left) + self._num(right) if op == "+" else self._num(left) - self._num(right)
        return left

    def _term(self):
        left = self._power()
        while self._peek() and self._peek()[1] in ("*", "/"):
            op = self._take()[1]
            right = self._power()
            if op == "*":
                left = self._num(left) * self._num(right)
            else:
                d = self._num(right)
                if d == 0:
                    raise FormulaError("#DIV/0!")
                left = self._num(left) / d
        return left

    def _power(self):
        left = self._unary()
        while self._peek() and self._peek()[1] == "^":
            self._take()
            left = self._num(left) ** self._num(self._unary())
        return left

    def _unary(self):
        tok = self._peek()
        if tok and tok[1] in ("-", "+"):
            self._take()
            v = self._num(self._unary())
            return -v if tok[1] == "-" else v
        return self._postfix()

    def _postfix(self):
        v = self._primary()
        while self._peek() and self._peek()[1] == "%":
            self._take()
            v = self._num(v) / 100.0
        return v

    def _primary(self):
        tok = self._take()
        kind, text = tok
        if kind == "num":
            return float(text)
        if kind == "str":
            return text[1:-1].replace('""', '"')
        if kind == "bool":
            return text == "TRUE"
        if kind == "ref":
            return self.coord(text)
        if kind == "range":
            return self._range(text)
        if kind == "func":
            return self._call(text)
        if text == "(":
            v = self._comparison()
            self._take(")")
            return v
        raise FormulaError(f"Unexpected token {text!r}")

    def _args(self) -> List[Any]:
        self._take("(")
        args: List[Any] = []
        if self._peek() and self._peek()[1] == ")":
            self._take(")")
            return args
        while True:
            args.append(self._comparison())
            t = self._take()
            if t[1] == ")":
                return args
            if t[1] != ",":
                raise FormulaError("Expected ',' or ')'")

    def _skip_arg(self) -> None:
        """Advances past one argument without evaluating it (stops before ',' or ')' at depth 0)."""
        depth = 0
        while True:
            tok = self._peek()
            if tok is None:
                raise FormulaError("Unterminated function call")
            if depth == 0 and tok[1] in (",", ")"):
                return
            if tok[1] == "(":
                depth += 1
            elif tok[1] == ")":
                depth -= 1
            self._i += 1

    def _call(self, name: str):
        if name == "IF":
            # Excel evaluates only the branch it returns: IF(F76>0, M75/F76, "") must not divide by zero
            self._take("(")
            cond = self._comparison()
            truthy = bool(cond) if not isinstance(cond, str) else cond != ""
            if self._take()[1] == ")":                       # IF(cond)
                return truthy
            if truthy:
                then_val = self._comparison()
            else:
                self._skip_arg()
                then_val = None
            if self._take()[1] == ")":                       # IF(cond, then) -> FALSE when not met
                return then_val if truthy else False
            if truthy:
                self._skip_arg()
                else_val = None
            else:
                else_val = self._comparison()
            self._take(")")
            return then_val if truthy else else_val
        if name == "IFERROR":
            # evaluate lazily so the error branch is honoured
            self._take("(")
            start = self._i
            try:
                v = self._comparison()
                self._take(",")
                depth = 0
                while True:            # skip the fallback expression
                    t = self._take()
                    if t[1] == "(":
                        depth += 1
                    elif t[1] == ")":
                        if depth == 0:
                            return v
                        depth -= 1
            except FormulaError:
                self._i = start
                depth = 0
                while True:            # skip the failing expression
                    t = self._take()
                    if t[1] == "(":
                        depth += 1
                    elif t[1] == ")":
                        depth -= 1
                    elif t[1] == "," and depth == 0:
                        break
                v = self._comparison()
                self._take(")")
                return v
        args = self._args()
        flat = lambda: [x for a in args for x in (a if isinstance(a, list) else [a])]
        if name == "SUM":
            return sum(self._num(x, True) for x in flat() if not isinstance(x, str))
        if name == "MIN":
            return min(self._num(x, True) for x in flat())
        if name == "MAX":
            return max(self._num(x, True) for x in flat())
        if name == "ABS":
            return abs(self._num(args[0]))
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            n = int(self._num(args[1])) if len(args) > 1 else 0
            fn = {"ROUND": _round, "ROUNDUP": _roundup, "ROUNDDOWN": _rounddown}[name]
            return fn(self._num(args[0]), n)
        if name == "AND":
            return all(bool(x) for x in flat())
        if name == "OR":
            return any(bool(x) for x in flat())
        raise FormulaError(f"Unsupported function {name}()")

    # -- helpers ----------------------------------------------------------
    @staticmethod
    def _split(ref: str) -> Tuple[int, int]:
        m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", ref)
        if not m:
            raise FormulaError(f"Bad reference {ref}")
        return column_index_from_string(m.group(1)), int(m.group(2))

    def _range(self, text: str) -> List[Any]:
        a, b = text.split(":")
        c1, r1 = self._split(a)
        c2, r2 = self._split(b)
        return [self.value(r, c) for r in range(min(r1, r2), max(r1, r2) + 1)
                for c in range(min(c1, c2), max(c1, c2) + 1)]

    @staticmethod
    def _num(v: Any, lenient: bool = False) -> float:
        if v is None or v == "":
            return 0.0
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, list):
            raise FormulaError("Range used where a single value is needed")
        if lenient:
            return 0.0
        try:
            return float(v)
        except (TypeError, ValueError):
            raise FormulaError(f"#VALUE! ({v!r} is not a number)")


def formula_refs(formula: Optional[str]) -> List[str]:
    """Cell references used in a formula (ranges expanded to their end points)."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return []
    return [r.replace("$", "") for r in re.findall(r"\$?[A-Z]{1,3}\$?\d+", formula[1:])]
