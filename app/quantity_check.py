"""Flags an extracted quantity that's wildly outside the historical norm for
the same (waste_type, unit) pair - a cheap, code-side second check that
doesn't need another API call, layered on top of (not instead of) the
prompt-level "double-check every digit" instruction in extractor.py's
FIELD_DEFS/_SYSTEM_PROMPT. Quantity is the single field this pipeline cares
most about getting right, so it gets two independent lines of defense: the
model being told to be careful, and this system-level sanity check that
doesn't rely on the model noticing its own mistake.

Historical reference comes from output/ריכוז_תעודות.xlsx, the same file
app/normalize.py already bootstraps its known-names list from (see
app/pipeline.py, which reads it once and hands the same records to both).
"""
from statistics import median
from typing import Dict, List, Tuple

# Fewer than this many historical (waste_type, unit) samples and there's not
# enough of a track record to call anything "typical" - skip the check
# quietly for that group rather than flag against a shaky reference point.
_MIN_HISTORY_SAMPLES = 3
# How many times above (or below, as a fraction) the historical median counts
# as "wildly outside the norm" - an order of magnitude, per the spec ("פי 10
# ומעלה, או קרובה לאפס"); "close to zero" is exactly the mirror image of
# that on the low side, so one ratio covers both directions.
_OUTLIER_RATIO = 10

QuantityHistory = Dict[Tuple[str, str], float]


def build_quantity_history(records: List[dict]) -> QuantityHistory:
    """Groups historical numeric quantities by (waste_type, unit), returning
    the median for groups with at least _MIN_HISTORY_SAMPLES - too few
    samples for a group means it's simply absent from the result, which
    flag_quantity_outlier() treats as "nothing to check against", per the
    spec's "otherwise just skip the check quietly".
    """
    buckets: Dict[Tuple[str, str], List[float]] = {}
    for record in records:
        quantity = record.get("quantity")
        waste_type, unit = record.get("waste_type"), record.get("unit")
        if isinstance(quantity, (int, float)) and waste_type and unit:
            buckets.setdefault((waste_type, unit), []).append(quantity)
    return {key: median(values) for key, values in buckets.items() if len(values) >= _MIN_HISTORY_SAMPLES}


def flag_quantity_outlier(record: dict, quantity: float, history: QuantityHistory) -> None:
    """Flags `record` (confidence -> נמוכה, a note appended) if `quantity` is
    at least _OUTLIER_RATIO times above or below the historical median for
    the same (waste_type, unit) - a strong signal of a misread digit (a
    misplaced decimal point, or a confused 0/6/8) rather than genuine
    real-world variation.

    Takes the already-parsed numeric quantity as its own argument rather than
    reading record["quantity"] directly, because at the point this runs
    (app.pipeline.process_files, right after extraction) that's still the
    raw string extract_certificate_pages() returned - parsing it into the
    float that actually lands in the Excel cell happens later, in
    excel_writer.write_records() (CLI) or streamlit_app.py (UI). The caller
    is expected to have already run it through excel_writer.parse_quantity().
    Does nothing if there's no history for this exact (waste_type, unit) pair.
    """
    if quantity <= 0:
        return
    typical = history.get((record.get("waste_type"), record.get("unit")))
    if not typical or typical <= 0:
        return
    if quantity >= typical * _OUTLIER_RATIO or quantity <= typical / _OUTLIER_RATIO:
        note = "כמות חריגה ביחס להיסטוריה - לבדוק ידנית"
        record["confidence"] = "נמוכה"
        record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note
