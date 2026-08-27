"""Fuzzy-matches free-text names (site/reference_type) extracted from a
certificate against a running list of names already seen, so near-duplicate
spellings of the same site or supplier ("VERIDIS" vs "veridis ltd" vs
"וריידיס בע\"מ" vs "וריידיס בעמ") collapse to one canonical value instead of
silently fragmenting the data - e.g. the region breakdown in the summary
sheet counts by exact string match (see excel_writer.py's note on that), so
two spellings of the same site are otherwise invisible to each other there.

Library choice: rapidfuzz, not stdlib difflib. Two reasons, both checked
empirically before picking this, not assumed:
  1. Speed: difflib.SequenceMatcher is pure Python; rapidfuzz is a C++
     extension built specifically for exactly this shape of workload
     (fuzzy-matching one new name against a list of known ones, repeatedly,
     once per certificate in a batch).
  2. Scoring quality: plain character-ratio scoring (rapidfuzz's own
     fuzz.ratio, or difflib's default) is case-sensitive and penalizes an
     added/missing word heavily - "veridis ltd" vs "VERIDIS" scores 0 with
     it, which is useless here. fuzz.WRatio with rapidfuzz's own
     utils.default_process (lowercases, strips punctuation/extra
     whitespace) scores that same pair 90 - actually usable. This
     combination isn't difflib's default behavior and would need
     hand-rolled preprocessing to match; rapidfuzz ships it built in.

MIT-licensed, ships prebuilt wheels (no C compiler needed on Windows or on
Streamlit Community Cloud's Linux containers).
"""
from typing import Dict, List

from rapidfuzz import fuzz, process, utils

# Below this score (0-100, from fuzz.WRatio), two names are treated as
# genuinely different - not spelling/formatting variants of the same one.
# 85 is the threshold suggested in the spec for this feature; matches like
# "veridis ltd"/"VERIDIS" (90) or a missing gershayim in a Hebrew ח.פ. suffix
# (95+) clear it comfortably, while genuinely different names score well
# below it (see the module docstring's tests, or app/pipeline.py's caller).
_SIMILARITY_THRESHOLD = 85

# Which extracted fields get normalized against a running "known names"
# list. Both are free text with no closed list (unlike waste_type/region -
# see fields.py) - exactly the kind of field where the same real-world name
# comes back spelled differently certificate to certificate.
NORMALIZED_FIELDS = ["site", "reference_type"]


class NameNormalizer:
    """Holds one growing "known names" list per normalized field for the
    duration of a batch. Bootstrap it from a project's already-processed
    records via build_from_records() so names from earlier runs are
    recognized too, not just ones seen since the normalizer was created -
    this project has no separate persisted "known names" file, so the
    existing output/ריכוז_תעודות.xlsx *is* that list.
    """

    def __init__(self) -> None:
        self._known: Dict[str, List[str]] = {field: [] for field in NORMALIZED_FIELDS}

    @classmethod
    def build_from_records(cls, records: List[dict]) -> "NameNormalizer":
        normalizer = cls()
        for record in records:
            for field in NORMALIZED_FIELDS:
                value = (record.get(field) or "").strip()
                if value:
                    normalizer._remember(field, value)
        return normalizer

    def _remember(self, field: str, value: str) -> None:
        known = self._known[field]
        if value not in known:
            known.append(value)

    def normalize(self, record: dict) -> None:
        """Normalizes record's free-text name fields in place. Whenever a
        value is close enough to an already-known one, it's rewritten to
        that canonical spelling and a "נורמל מ-X" note is appended (X being
        what the model actually returned) so the correction stays visible,
        not silent. A value with no close match is kept as its own spelling
        and remembered, so later records in the same batch that are close to
        *it* normalize to it in turn.
        """
        for field in NORMALIZED_FIELDS:
            value = (record.get(field) or "").strip()
            if not value:
                continue
            known = self._known[field]
            match = (
                process.extractOne(value, known, scorer=fuzz.WRatio, processor=utils.default_process)
                if known
                else None
            )
            if match is not None and match[1] >= _SIMILARITY_THRESHOLD and match[0] != value:
                canonical = match[0]
                note = f'נורמל מ-"{value}"'
                record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note
                record[field] = canonical
                value = canonical  # don't also remember the raw spelling as a new "known" name
            self._remember(field, value)
