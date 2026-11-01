import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

BENGSATRA_BILL = ROOT / "Doc" / "E-Bill__ PHED, ASSAM bengsatra.pdf"
BENGSATRA_TEMPLATE = ROOT / "Doc" / "BISONDOI PT-II BANG SATRA.xlsx"
TOLLI_BILL = ROOT / "Doc" / "test" / "E-bill.pdf"
TOLLI_TS = ROOT / "Doc" / "test" / "ts.pdf"
FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _need(*paths):
    missing = [p for p in paths if not p.exists()]
    if missing:
        pytest.skip(f"sample document missing: {missing[0]}")


@pytest.fixture(scope="session")
def bengsatra_ebill():
    _need(BENGSATRA_BILL)
    from ebill_parser import parse_ebill_pdf
    return parse_ebill_pdf(BENGSATRA_BILL.read_bytes())


@pytest.fixture(scope="session")
def bengsatra_template_bytes():
    _need(BENGSATRA_TEMPLATE)
    return BENGSATRA_TEMPLATE.read_bytes()


@pytest.fixture(scope="session")
def tolli_abstract():
    """Scanned abstract read with OCR (~25 s, done once per test session)."""
    _need(TOLLI_BILL)
    from scanned_abstract import parse_scanned_abstract
    return parse_scanned_abstract(TOLLI_BILL.read_bytes())


@pytest.fixture(scope="session")
def tolli_ts():
    _need(TOLLI_TS)
    from ts_parser import parse_ts_pdf
    return parse_ts_pdf(TOLLI_TS.read_bytes())
