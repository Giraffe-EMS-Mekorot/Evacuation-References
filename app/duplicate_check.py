"""Flags potential duplicate certificates: the same real-world shipment
reported twice through two different channels - typically a manually-typed
Excel row entered ahead of time (see app/excel_input.py), and the actual
scanned certificate for the same shipment arriving later and getting run
through the normal AI-extraction pipeline (see app/extractor.py).

**Detection only - this never auto-merges or auto-deletes a row.** Collapsing
two rows automatically risks silently discarding a real difference between
them (a genuinely different certificate that just happens to look similar),
which is too dangerous to do without a human looking at it - per an explicit
requirement. The note this leaves on both rows is for a reviewer to resolve
by hand.
"""
from typing import List, Optional

from .derive import parse_date
from .excel_writer import parse_quantity

# "תאריך קרוב" - a few days' slack, not exact-day-only, since the manual
# entry and the later scan may honestly disagree slightly on which day a
# shipment happened (e.g. logged on pickup date vs. arrival date).
_DATE_PROXIMITY_DAYS = 3
# "כמות קרובה/זהה" - within 5% of each other counts as "the same quantity,
# possibly rounded/re-entered slightly differently", not a coincidence.
_QUANTITY_PROXIMITY_RATIO = 0.05

_DUPLICATE_NOTE_TEMPLATE = (
    '⚠ כפילות אפשרית - ראה גם שורה ממקור "{source}"{date_part}, ייתכן שזה אותו '
    "משלוח שדווח פעמיים (ידני+סרוק)"
)


def _is_manual_excel_record(record: dict) -> bool:
    """A record's origin (hand-typed Excel vs. AI-extracted PDF/image) isn't
    tracked as its own field - it's inferred from source_file's own
    extension, which is already a persisted, always-available column (see
    fields.EXCEL_COLUMNS) that survives a round-trip through
    excel_writer.read_existing_records(). That means this check works
    identically on records built this session and on rows read back from a
    previously-saved project file, with no extra bookkeeping.
    """
    return (record.get("source_file") or "").lower().endswith((".xlsx", ".xls"))


def _quantities_close(a: Optional[float], b: Optional[float]) -> bool:
    if a is None or b is None:
        return False
    if a == b:
        return True
    return abs(a - b) <= max(abs(a), abs(b)) * _QUANTITY_PROXIMITY_RATIO


def flag_potential_duplicates(records: List[dict]) -> None:
    """Compares every (manual-Excel, AI-extracted) pair in `records` and
    appends a note to BOTH sides of any pair whose date, site, waste_type,
    and quantity all look like the same real-world shipment.

    Only ever compares across the two different origins (see
    _is_manual_excel_record) - two rows from the SAME channel that happen to
    look similar isn't this specific failure mode (a supplier may genuinely
    deliver the same waste type from the same site on two nearby dates) and
    comparing same-origin pairs too would produce far more false positives
    than this feature is meant to catch.

    The note references the OTHER row by source file name (and date, when
    known) rather than a spreadsheet row NUMBER - deliberately: row position
    shifts every time the sheet is re-sorted by confidence (see
    excel_writer.sort_by_confidence, which write_records() always applies)
    or a later batch adds more rows, so a numeric reference would go stale
    the moment either happens. A source-file+date reference stays correct
    and is at least as easy for a reviewer to search/filter for in Excel.

    Idempotent by design (safe to call again on the same records, e.g. on
    every Streamlit rerun that re-writes the project file): the "note
    already present" check below matches on the note's own stable source-
    file text, not a volatile row number, so re-running this never
    re-appends the same flag or drifts out of sync with a stale position.
    """
    for i, a in enumerate(records):
        a_is_manual = _is_manual_excel_record(a)
        a_date = parse_date(a.get("date"))
        a_site = (a.get("site") or "").strip().casefold()
        a_waste = (a.get("waste_type") or "").strip()
        a_qty = parse_quantity(a.get("quantity"))
        if not a_site or not a_waste or a_qty is None:
            continue
        for j in range(i + 1, len(records)):
            b = records[j]
            if _is_manual_excel_record(b) == a_is_manual:
                continue  # only cross-origin pairs are this specific failure mode
            b_date = parse_date(b.get("date"))
            b_site = (b.get("site") or "").strip().casefold()
            b_waste = (b.get("waste_type") or "").strip()
            b_qty = parse_quantity(b.get("quantity"))
            if b_site != a_site or b_waste != a_waste:
                continue
            if not _quantities_close(a_qty, b_qty):
                continue
            if a_date and b_date:
                if abs((a_date - b_date).days) > _DATE_PROXIMITY_DAYS:
                    continue
            elif a_date != b_date:  # exactly one side has a date - too uncertain to flag
                continue

            _append_duplicate_note(a, b)
            _append_duplicate_note(b, a)


def _append_duplicate_note(record: dict, other: dict) -> None:
    other_source = other.get("source_file", "")
    other_date = other.get("date", "")
    date_part = f" (תאריך {other_date})" if other_date else ""
    note = _DUPLICATE_NOTE_TEMPLATE.format(source=other_source, date_part=date_part)
    stable_prefix = f'ראה גם שורה ממקור "{other_source}"'
    if stable_prefix in (record.get("notes") or ""):
        return
    record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note
