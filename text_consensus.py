"""
text_consensus.py - One exact text from two independent OCR readings of the same cell.

RapidOCR and Windows OCR make different mistakes on the same photo ('fttig' / 'fitting',
'Consü-uction' / 'Construction'), so their readings are aligned word by word:

  * words both engines read alike are kept;
  * where they differ, the reading that is a known TS word wins. Known words come from the
    vocabulary learnt from TS documents (data/ts_vocabulary.txt, see tools/build_vocabulary.py),
    so the documents' own spellings ('Borad', 'submersibble', 'Disel', 'ofClear') are kept and
    never "corrected" into dictionary English;
  * a difference that neither reading resolves is reported as uncertain (shown for review).

Accented letters are folded to plain ASCII first: TS text is English, an 'í' is always a misread 'i'.
"""

import re
import unicodedata
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

VOCAB_FILE = Path(__file__).resolve().parent / "data" / "ts_vocabulary.txt"
KNOWN_MIN_DOCS = 2

_PUNCT_MAP = str.maketrans({"‘": "'", "’": "'", "‛": "'", "`": "'", "´": "'", "“": '"', "”": '"', "„": '"',
                            "–": "-", "—": "-", "−": "-", "‐": "-", "|": " "})


def fold(text: str) -> str:
    """Plain-ASCII form: accents dropped, typographic quotes / dashes made ASCII, other symbols removed."""
    t = unicodedata.normalize("NFKD", (text or "").translate(_PUNCT_MAP))
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = "".join(ch for ch in t if ch == "₹" or 32 <= ord(ch) < 127)
    return re.sub(r"\s+", " ", t).strip()


def key(token: str) -> str:
    return re.sub(r"[^a-z0-9]", "", fold(token).lower())


class Vocabulary:
    def __init__(self, counts: Optional[Dict[str, int]] = None):
        self.counts = dict(counts or {})

    def known(self, k: str) -> bool:
        # single letters ('s', 'm' of 'Sqm' read apart) are never evidence of a word
        return len(k) >= 2 and self.counts.get(k, 0) >= KNOWN_MIN_DOCS

    def add(self, keys: Iterable[str], n: int = 1):
        for k in keys:
            if k:
                self.counts[k] = self.counts.get(k, 0) + n

    def save(self, path: Path = VOCAB_FILE):
        path.parent.mkdir(parents=True, exist_ok=True)
        lines = [f"{k}\t{n}" for k, n in sorted(self.counts.items(), key=lambda kv: (-kv[1], kv[0]))]
        path.write_text("# word key <TAB> number of TS documents it was read in by both engines\n"
                        + "\n".join(lines) + "\n", encoding="utf-8")


@lru_cache(maxsize=1)
def load_vocabulary(path: str = str(VOCAB_FILE)) -> Vocabulary:
    counts: Dict[str, int] = {}
    p = Path(path)
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            if line.startswith("#") or "\t" not in line:
                continue
            k, n = line.split("\t", 1)
            counts[k] = int(n)
    return Vocabulary(counts)


def _score(tok: str, vocab: Vocabulary) -> int:
    """+1 a known TS word, 0 punctuation / figures, -1 anything else (also stray 1-2 letter pieces)."""
    k = key(tok)
    if not k or any(c.isdigit() for c in k):
        return 0
    return 1 if vocab.known(k) else -1


def _punct(tokens: List[str]) -> int:
    return sum(len(re.sub(r"[A-Za-z0-9\s]", "", t)) for t in tokens)


def _strip_glued(tokens: List[str], drop: Set[str]) -> List[str]:
    """'(2Job' / '30000Job' / 'ofJob' -> '(2' / '30000' / 'of' when 'Job' belongs to the unit cell."""
    out = []
    for t in tokens:
        for d in drop:
            if d and len(t) > len(d) and t[-len(d):].lower() == d and key(t) != d:
                t = t[:-len(d)]
                break
        out.append(t)
    return out


def merge(primary: str, secondary: str, vocab: Optional[Vocabulary] = None,
          drop: Set[str] = frozenset()) -> Tuple[str, List[str]]:
    """
    primary: RapidOCR reading (better on digits); secondary: Windows OCR reading (better word gaps).
    drop: word keys that belong to a neighbouring cell (a unit printed against the description).
    Returns (text, uncertain words).
    """
    vocab = vocab or load_vocabulary()
    r = _strip_glued(fold(primary).split(), drop)
    w = _strip_glued(fold(secondary).split(), drop)
    if not r:
        # the second engine only refines what the main one placed in this cell: its words alone are
        # usually a neighbouring row's text whose line centre fell here
        return "", []
    if not w:
        return " ".join(t for t in r if key(t) not in drop or not key(t)), []
    kr, kw = [key(t) or t for t in r], [key(t) or t for t in w]
    out: List[str] = []
    uncertain: List[str] = []
    for op, i1, i2, j1, j2 in SequenceMatcher(None, kr, kw, autojunk=False).get_opcodes():
        R, W = r[i1:i2], w[j1:j2]
        if op == "equal":
            for a, b in zip(R, W):
                out.append(b if _punct([b]) > _punct([a]) else a)    # a faint comma is dropped, not invented
        elif op == "delete":
            out.extend(t for t in R if not (key(t) in drop and key(t)) and (key(t) or t in "&/-()"))
        elif op == "insert":
            out.extend(t for t in W if _score(t, vocab) == 1 and key(t) not in drop)
        else:
            # the main engine's reading stands unless the second one is the better-attested TS text
            if "".join(key(t) for t in R) == "".join(key(t) for t in W):
                # same letters, different spacing / punctuation ('4Nos' / '4 Nos'): word gaps from
                # Windows OCR, unless that splits known words into unknown pieces or drops punctuation
                bad_split = any(_score(t, vocab) < 0 for t in W) and all(_score(t, vocab) >= 0 for t in R)
                choice = R if (bad_split or _punct(R) > _punct(W)) else W
            elif len(R) == len(W):
                choice = []
                for a, b in zip(R, W):
                    sa, sb = _score(a, vocab), _score(b, vocab)
                    choice.append(b if sb > sa else a)
                    if sa == sb == -1:
                        uncertain.append(a)
            else:
                sR, sW = sum(_score(t, vocab) for t in R), sum(_score(t, vocab) for t in W)
                choice = W if sW > sR else R
                if sR == sW:
                    uncertain.extend(t for t in choice if _score(t, vocab) < 0)
            out.extend(t for t in choice if not (key(t) in drop and key(t)))
    return re.sub(r"\s+", " ", " ".join(out)).strip(), uncertain
