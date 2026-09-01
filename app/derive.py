"""Pure, code-side derivations computed from extracted fields (no model call).

Kept separate from extractor.py because this is plain computation, not
something Claude is asked to infer - deriving year/month from a DD/MM/YYYY
string is unambiguous and shouldn't be left to a model that (per the known
limitation documented in CLAUDE.md) sometimes guesses instead of leaving a
field blank.
"""
import re
from datetime import date as _date
from typing import Optional, Tuple

from .fields import HEBREW_MONTHS

# Accepts both DD/MM/YYYY (the requested format) and DD/MM/YY, since the
# model doesn't always follow the requested format exactly.
_DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})$")


def derive_year_month(date_str: str) -> Tuple[str, str]:
    """Derives (year, hebrew_month_name) from a DD/MM/YYYY (or DD/MM/YY) date string.

    Returns ("", "") if date_str is empty or doesn't match a recognizable
    date format - same "leave blank rather than guess" policy as every other
    field in this pipeline.
    """
    match = _DATE_RE.match((date_str or "").strip())
    if not match:
        return "", ""
    day, month, year = (int(part) for part in match.groups())
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return "", ""
    if year < 100:
        year += 2000 if year <= 79 else 1900
    return str(year), HEBREW_MONTHS[month - 1]


def parse_date(date_str: str) -> Optional[_date]:
    """Parses a DD/MM/YYYY (or DD/MM/YY) date string into a real `date`
    object, for callers that need actual day-level arithmetic (date-
    proximity matching - see pipeline._cross_check_bill_of_lading_estimates
    and duplicate_check.flag_potential_duplicates) rather than just the
    year/month derive_year_month() returns.

    Deliberately a separate function reusing the same _DATE_RE, not a
    refactor of derive_year_month() itself: date(year, month, day) raises on
    a numerically-in-range but calendar-invalid combination (e.g. 31/04 -
    April has 30 days), which derive_year_month() has never actually
    checked for (it only validates day is 1-31 and month is 1-12
    generically) - changing that function's own behavior now, however
    arguably more correct, risks the exact kind of "the date logic changed
    and now something upstream disagrees" report this project has already
    chased down as a false alarm once (see CLAUDE.md's 2026-08-30
    follow-up). Same "None over guessing" policy either way: unparseable or
    calendar-invalid input returns None, never raises.
    """
    match = _DATE_RE.match((date_str or "").strip())
    if not match:
        return None
    day, month, year = (int(part) for part in match.groups())
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    if year < 100:
        year += 2000 if year <= 79 else 1900
    try:
        return _date(year, month, day)
    except ValueError:
        return None
