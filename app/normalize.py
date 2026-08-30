"""Fuzzy-matches free-text names (site/supplier_or_carrier) extracted from a
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

**Reversed-Hebrew detection (2026-08-30)** - a distinct concern from the
fuzzy-matching above, triggered by a real (if not yet directly observed on
this project's own certificates) vision-model failure mode: RTL Hebrew read
off an image in left-to-right pixel/glyph order comes out character-order
-reversed. looks_reversed_hebrew() is the shared, dictionary-free detector
(also used by extractor.py, which has no "known names" list to correct
against - it can only flag); NameNormalizer.normalize() below uses the same
idea in its own else-branch to actively *correct* a reversed value, since it
does have a known-names list to check the reversal against.
"""
from typing import Dict, List

from rapidfuzz import fuzz, process, utils

# Letters with a distinct final ("sofit") form, valid ONLY as the last
# character of a Hebrew word - and the regular-form counterpart essentially
# never legitimately ends one either. Either violation (a final form
# mid-word, or a should-be-final regular form at the end) is a strong,
# dictionary-free signal that a word's characters came out in reversed (or
# otherwise scrambled) order - see looks_reversed_hebrew().
_FINAL_FORM_TO_REGULAR = {"ך": "כ", "ם": "מ", "ן": "נ", "ף": "פ", "ץ": "צ"}
_FINAL_FORM_LETTERS = set(_FINAL_FORM_TO_REGULAR)
_REGULAR_LETTERS_WITH_FINAL_FORM = set(_FINAL_FORM_TO_REGULAR.values())


def _is_hebrew_letter(ch: str) -> bool:
    return "א" <= ch <= "ת"


def _hebrew_word_looks_reversed(word: str) -> bool:
    if len(word) < 2:
        return False
    if any(ch in _FINAL_FORM_LETTERS for ch in word[:-1]):
        return True
    return word[-1] in _REGULAR_LETTERS_WITH_FINAL_FORM


def looks_reversed_hebrew(text: str) -> bool:
    """True if any whitespace-separated, Hebrew-letters-only word in `text`
    violates final-letter-form placement (see _hebrew_word_looks_reversed) -
    a plain Hebrew-orthography rule, not a dictionary lookup, so it needs no
    word list and works on a name never seen before. Skips any word that
    mixes in a digit/Latin letter/punctuation (e.g. "מ.עבד", "VERIDIS",
    "ח.פ. 12-345") - the rule only holds for plain Hebrew spelling, and a
    mixed token would misfire it in both directions.
    """
    for word in (text or "").split():
        if word and all(_is_hebrew_letter(c) for c in word) and _hebrew_word_looks_reversed(word):
            return True
    return False


# At or above this score (0-100, from fuzz.WRatio), a name is treated as an
# OCR/handwriting variant of an already-known one and silently normalized to
# it - matches like "veridis ltd"/"VERIDIS" (90) or a missing gershayim in a
# Hebrew ח.פ. suffix (95+) clear it comfortably. Below it, a name is treated
# as a genuinely new supplier/site rather than illegible handwriting of a
# known one - see normalize()'s two distinct notes for these cases. 70, not a
# stricter value, because the whole point of this lower band is to still
# catch real OCR noise, just noisier than the once-85-only cases were.
_SIMILARITY_THRESHOLD = 70

# Which extracted fields get normalized against a running "known names"
# list. Both are free text with no closed list (unlike waste_type/region -
# see fields.py) - exactly the kind of field where the same real-world name
# comes back spelled differently certificate to certificate.
NORMALIZED_FIELDS = ["site", "supplier_or_carrier"]


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
        """Normalizes record's free-text name fields in place. Four cases,
        each noted distinctly rather than lumped together:

          - Close match, different spelling (score >= _SIMILARITY_THRESHOLD,
            not identical): treated as the same real-world site/supplier read
            with OCR/handwriting noise - rewritten to the known spelling,
            with a "נורמל מ-X" note (X being what the model actually
            returned) so the correction stays visible, not silent.
          - Close match, identical spelling: already exactly a known value -
            nothing to note.
          - No forward match, but the value's REVERSED characters match a
            known name well: treated as reversed-character-order text (see
            looks_reversed_hebrew's docstring for the failure mode this
            guards against) rather than a genuinely new name - corrected to
            the known spelling with its own distinct note, same as the
            forward-match case above but naming the actual root cause.
          - No match either way: treated as a genuinely new site/supplier,
            *not* illegible handwriting of a known one - noted as such
            explicitly (distinct from "כתב יד לא קריא", which is the model's
            own call about legibility, not this system-level "is this even
            in our history" check) and remembered, so a later record in the
            same batch that's close to *this* new spelling normalizes to it
            in turn instead of also being flagged as new.
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
            if match is not None and match[1] >= _SIMILARITY_THRESHOLD:
                canonical = match[0]
                if canonical != value:
                    note = f'נורמל מ-"{value}"'
                    self._append_note(record, note)
                    record[field] = canonical
                    value = canonical  # don't also remember the raw spelling as a new "known" name
            else:
                reversed_match = (
                    process.extractOne(value[::-1], known, scorer=fuzz.WRatio, processor=utils.default_process)
                    if known
                    else None
                )
                if reversed_match is not None and reversed_match[1] >= _SIMILARITY_THRESHOLD:
                    canonical = reversed_match[0]
                    note = f'תוקן מטקסט שנקרא הפוך (RTL): "{value}"'
                    self._append_note(record, note)
                    record[field] = canonical
                    value = canonical
                else:
                    self._append_note(record, "ספק/אתר חדש - לא קיים ברשימת הייחוס")
            self._remember(field, value)

    @staticmethod
    def _append_note(record: dict, note: str) -> None:
        # site and supplier_or_carrier are checked independently and can easily
        # carry the exact same value (many certificates only ever name the
        # supplier once, reused for both fields) - without this guard, one
        # normalize() call could append the identical note twice.
        existing = record.get("notes") or ""
        if note in existing:
            return
        record["notes"] = f"{existing} | {note}" if existing else note
