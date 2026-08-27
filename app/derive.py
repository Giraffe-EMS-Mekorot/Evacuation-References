"""Pure, code-side derivations computed from extracted fields (no model call).

Kept separate from extractor.py because this is plain computation, not
something Claude is asked to infer - deriving year/month from a DD/MM/YYYY
string is unambiguous and shouldn't be left to a model that (per the known
limitation documented in CLAUDE.md) sometimes guesses instead of leaving a
field blank.
"""
import re
from typing import Tuple

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
