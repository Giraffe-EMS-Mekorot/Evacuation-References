"""Parses a "דו\"ח סיכום חודשי ללקוח" - the consolidated monthly summary
report ק.מ.מ. מפעלי מיחזור בע"מ issues per customer - into one record per
data row.

**Deterministic text parsing, no model call.** Unlike every other document
this project handles, this one is a digitally-generated PDF with a real text
layer: no handwriting, no photograph, no OCR. Sending it to a vision model
would be slower, cost money per page, and be *less* accurate than reading the
text that is already in the file. It is parsed with pypdfium2, which is
already a dependency (used by extractor.py to rasterize pages), so this adds
no new package.

That choice has a consequence worth stating: **none of the pipeline's
vision-error safety nets apply to these records** - no NameNormalizer, no
quantity-history outlier check, no gross/tare arithmetic, no date-plausibility
flag. Those all exist to catch a vision model misreading something, and
nothing here was read by a model. Same reasoning app/excel_input.py already
applies to a hand-typed Excel row; see pipeline.process_files.

**One file becomes many rows.** The report is structured as
customer-center -> station -> month -> material, and one real file
(input/דוח כמויות מקורות 1-6.26.pdf, 2 pages) contains 16 stations and 68
data rows. A row here is one (station, month, material) triple with its own
weight.

## The two parsing traps in this format

**1. Extracted text is in visual order, not logical order.** The PDF stores
RTL Hebrew laid out left-to-right, so a line comes out with its words
reversed and punctuation attached to the wrong side: the header
`מס. לקוח: KM103527` extracts as `,KM103527 :לקוח .מס`, and `מ-8` as `8-מ`.
Every line is therefore token-reversed before being read (see `_logical`),
and fields are located by anchored markers rather than by column position.

**2. A date is printed only when it changes.** Consecutive rows in the same
month leave the date cell empty - 7 of the 68 rows in the real file do. The
date is forward-filled from the last row that had one (`_DATE_RE`), reset at
each new station. Without that, those 7 rows would silently get a blank
תאריך.

Because both traps are silent-corruption risks rather than crashes, any data
line that does not parse is turned into a flagged record with the raw line in
its notes (see `_unparsed_record`) rather than skipped - a dropped row here
would be invisible.
"""
import re
from pathlib import Path
from typing import List, Optional

import pypdfium2 as pdfium

from .fields import DOCUMENT_TYPE_KMM_SUMMARY_REPORT, empty_record

# Text markers that identify this report. Checked against the *logical*
# (token-reversed) text of the first page - see looks_like_kmm_report(). Two
# independent signals are required rather than one, so an ordinary
# certificate that merely mentions ק.מ.מ as a supplier is not misdetected
# (that exact company already appears in this project's history as an
# ordinary invoice issuer - see CLAUDE.md's invoice-filtering note).
_ISSUER_MARKERS = ("ק.מ.מ", "k.m.m")
_REPORT_TITLE_MARKERS = ("דו\"ח סיכום חודשי", "סיכום חודשי ללקוח")
# The per-station block header, and the customer-center header above it.
_STATION_MARKER = "שם תחנה"
_CUSTOMER_NO_RE = re.compile(r"KM\d+")
_ADDRESS_MARKER = "כתובת"

# A month cell, as printed: "ינו-2026". Kept as free text in the record's
# date field per spec ("כפי שמופיע במקור ... לא להמיר לפורמט תאריך מדויק"),
# which also means derive.derive_year_month() will not parse it - deliberate,
# see parse_kmm_report().
_DATE_RE = re.compile(r"^\d{4}-[֐-׿]{3}$")
_WEIGHT_RE = re.compile(r"^[\d,]+$")
_SKU_RE = re.compile(r"^\d+$")

# The unit is fixed for the whole report - it is stated once in the title
# ("משקל בק\"ג") and never repeated per row.
KMM_UNIT = 'ק"ג'

# The issuing company, for the "אתר קולט" column. Taken as a constant rather
# than scraped from the header line: the header is the most bidi-mangled line
# on the page (mixed Latin "k.m.m" and Hebrew with embedded quotes), and this
# is the one value in the report that is identical in every file it issues.
KMM_SUPPLIER = 'ק.מ.מ. מפעלי מיחזור בע"מ'

# The report's own material names -> this project's closed WASTE_TYPES list.
# Both real values map cleanly; an unmapped material is kept verbatim and
# flagged rather than guessed (see _map_waste_type).
MATERIAL_TO_WASTE_TYPE = {
    "קרטון": "קרטונים",
    "נייר לבן": "נייר",
    "נייר מסווג": "נייר",
    "נייר": "נייר",
}

_BASE_NOTE = 'מתוך דוח קמ"מ מרוכז'


# A token like "8-מ", which is how bidi extraction renders "מ-8" (digits and
# Hebrew swap sides around the hyphen). Only the digits+Hebrew case is
# corrected; a Hebrew-Hebrew hyphenation like "חודש-אשכול" is left alone,
# since there is no way to tell which side was originally first.
_SWAPPED_HYPHEN_RE = re.compile(r"^(\d+)-([֐-׿]+)$")


def _logical(line: str) -> str:
    """Reverses a line's token order, turning the PDF's visual-order RTL text
    into readable logical order - see this module's docstring, trap #1.

    Token-internal punctuation is NOT fixed here: bidi extraction leaves a
    colon on the wrong side of its word, so "לקוח:" arrives as ":לקוח" and a
    two-word label like "שם תחנה:" arrives as "שם :תחנה" - **with the colon
    wedged between the two words.** That is why marker matching must go
    through _normalized() below and not this function; matching "שם תחנה"
    against this output silently never fires, which is exactly what happened
    on the first run against the real file (all 68 rows were flagged
    unparsed).
    """
    return " ".join(reversed(line.split()))


def _normalized(line: str) -> str:
    """_logical() with the stray colons/commas removed and whitespace
    collapsed - the form to match field markers and slice values out of.
    """
    text = _logical(line)
    text = re.sub(r"[:,]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _fix_swapped_hyphens(text: str) -> str:
    """Restores "מ-8" from the "8-מ" that bidi extraction produces, so a
    station name reads as printed ("...אחיסמך מ-8 עד 15").
    """
    return " ".join(
        _SWAPPED_HYPHEN_RE.sub(lambda m: f"{m.group(2)}-{m.group(1)}", tok)
        for tok in text.split()
    )


def _looks_like_data_row(tokens: List[str]) -> bool:
    """A data row, in VISUAL order, starts with the weight and contains a
    separate integer מק"ט somewhere after it. Checked on raw (unreversed)
    tokens because that is the order the weight arrives first in.
    """
    if not tokens or not _WEIGHT_RE.match(tokens[0]):
        return False
    return any(_SKU_RE.match(t) for t in tokens[1:])


def _station_name(normalized_line: str) -> str:
    """The full station name as printed next to "שם תחנה:", e.g.
    "מקורות מרחב מרכז אחיסמך מ-8 עד 15". Bounded by the next field marker
    ("כתובת") rather than by a comma, since commas inside the name are
    indistinguishable from the field separator once bidi-mangled.
    """
    if _STATION_MARKER not in normalized_line:
        return ""
    after = normalized_line.split(_STATION_MARKER, 1)[1]
    name = after.split(_ADDRESS_MARKER, 1)[0]
    return _fix_swapped_hyphens(name.strip(" ,:\u05f4"))


def _station_address(normalized_line: str) -> str:
    """The station's address, for the notes column (per spec, "הכתובת של
    התחנה אם קיימת בדו"ח, לצורך הקשר"). Bounded by the next field marker
    after it, which is the customer-type code.
    """
    if _ADDRESS_MARKER not in normalized_line:
        return ""
    after = normalized_line.split(_ADDRESS_MARKER, 1)[1]
    for stop in ("קוד סוג", "תאור סוג", "פרמטר"):
        after = after.split(stop, 1)[0]
    return _fix_swapped_hyphens(after.strip(" ,:"))


def _map_waste_type(material: str) -> tuple:
    """(waste_type, note). Maps the report's own material name onto the closed
    WASTE_TYPES list. An unrecognized material is returned verbatim with a
    note, never silently coerced into a neighbouring category - same
    "flag, don't guess" policy the rest of this pipeline follows.
    """
    text = (material or "").strip()
    if text in MATERIAL_TO_WASTE_TYPE:
        return MATERIAL_TO_WASTE_TYPE[text], ""
    for name, mapped in MATERIAL_TO_WASTE_TYPE.items():
        if name and name in text:
            return mapped, ""
    return text, f'סוג חומר לא ממופה בדוח קמ"מ: {text}'


def _parse_data_row(tokens: List[str]) -> Optional[dict]:
    """Splits one visual-order data row into its five printed columns.

    Visual order is `weight | <product description> | מק"ט | <material> | date?`
    which is the logical `date? | material | מק"ט | product description | weight`
    read backwards. Anchoring on the two unambiguous ends - the weight is
    always token 0, the date (when present) is always the last token - and on
    the integer מק"ט between them is what makes this robust without needing
    per-column x-coordinates.
    """
    if not _looks_like_data_row(tokens):
        return None
    weight = tokens[0]
    rest = tokens[1:]

    date = ""
    if rest and _DATE_RE.match(rest[-1]):
        date = rest[-1]
        rest = rest[:-1]

    sku_index = next((i for i, t in enumerate(rest) if _SKU_RE.match(t)), None)
    if sku_index is None:
        return None

    product = _logical(" ".join(rest[:sku_index]))
    material = _logical(" ".join(rest[sku_index + 1:]))
    return {
        "weight": weight,
        "sku": rest[sku_index],
        "product": product.strip(),
        "material": material.strip(),
        "date": date,
    }


_VISUAL_MONTH_RE = re.compile(r"^(\d{4})-([֐-׿]{3})$")


def _month_label(raw_date: str) -> str:
    """The printed month cell as text, e.g. "ינו 2026".

    The token arrives as "2026-ינו" - the bidi rendering of "ינו-2026" - so
    the two halves are swapped back. Deliberately NOT converted to a real
    date: per spec the תאריך column keeps the report's own free-text month
    ("כפי שמופיע במקור ... לא להמיר לפורמט תאריך מדויק"), which also means
    derive.derive_year_month() will not parse it and year/month stay blank on
    these records.
    """
    if not raw_date:
        return ""
    match = _VISUAL_MONTH_RE.match(raw_date.strip())
    if match:
        return f"{match.group(2)} {match.group(1)}"
    return raw_date.replace("-", " ").strip()


def looks_like_kmm_report(path: Path) -> bool:
    """Whether `path` is a ק.מ.מ consolidated monthly summary report.

    Requires BOTH an issuer marker and a report-title marker on the first
    page, so an ordinary certificate or invoice that merely names ק.מ.מ as
    its supplier is not misrouted into this parser. Returns False for
    anything that is not a readable PDF - a non-PDF, an image, or a scanned
    PDF with no text layer all fall through to the normal vision path rather
    than erroring.
    """
    if path.suffix.lower() != ".pdf":
        return False
    try:
        doc = pdfium.PdfDocument(str(path))
    except Exception:
        return False
    try:
        if len(doc) == 0:
            return False
        raw = doc[0].get_textpage().get_text_range()
    except Exception:
        return False
    finally:
        doc.close()

    text = raw + "\n" + "\n".join(_normalized(line) for line in raw.split("\n"))
    has_issuer = any(m in text for m in _ISSUER_MARKERS) or "k.m.m" in text.lower()
    has_title = any(m in text for m in _REPORT_TITLE_MARKERS)
    return has_issuer and has_title


def _unparsed_record(filename: str, page_number: int, line: str) -> dict:
    """A data-looking line this parser could not split. Recorded as a flagged
    row rather than dropped: a silently missing row in a 68-row report is
    invisible, while a "נמוכה" row with the raw text in its notes is not.
    """
    record = empty_record(filename)
    record["document_type"] = DOCUMENT_TYPE_KMM_SUMMARY_REPORT
    record["page_number"] = page_number
    record["confidence"] = "נמוכה"
    record["notes"] = f'{_BASE_NOTE} | שורה לא פורסרה: {line.strip()[:120]}'
    return record


def parse_kmm_report(path: Path) -> List[dict]:
    """One record per data row in the report, in document order.

    Field mapping (per spec):
      site                     <- the station's full "שם תחנה" text
      waste_type               <- the row's material, mapped to WASTE_TYPES
      quantity                 <- the row's משקל
      unit                     <- KMM_UNIT (fixed for the whole report)
      supplier_or_carrier      <- KMM_SUPPLIER (the issuing company)
      date                     <- the month as printed ("ינו 2026"), free text
      certificate_or_reference <- the station's "מס. לקוח" (e.g. KM103527)
      notes                    <- _BASE_NOTE + the station's address
      source_file/page_number  <- file name + the page the row is on

    `region` is deliberately NOT set here - it comes from the user's per-batch
    selection (see pipeline.apply_kmm_selection), which is the one field of
    the three that this document does not state per row.

    vehicle_number/driver_name/gross_weight/tare_weight/reference_number stay
    empty: they do not exist in this format. reference_number in particular is
    left empty *on purpose* even though certificate_or_reference is filled -
    that keeps these records out of
    pipeline._cross_check_weighing_certificates, which indexes candidates by
    reference_number, so a customer number can never be mistaken for a
    תעודת שקילה's printed number.
    """
    doc = pdfium.PdfDocument(str(path))
    records: List[dict] = []
    # Station state lives OUTSIDE the page loop on purpose: a station's block
    # continues across a page break, and its rows on the next page have no
    # repeated "שם תחנה" header. Resetting per page silently orphaned the
    # first 6 rows of page 2 of the real reference file (1,305 of its 18,670
    # ק"ג) - they were flagged unparsed rather than lost, which is how this was
    # caught, but they still belonged to the last station of page 1.
    station_name = ""
    station_address = ""
    customer_no = ""
    current_date = ""
    try:
        for page_index in range(len(doc)):
            raw = doc[page_index].get_textpage().get_text_range()
            page_number = page_index + 1

            for line in raw.split("\n"):
                line = line.strip()
                if not line:
                    continue

                if _STATION_MARKER in _normalized(line):
                    normalized = _normalized(line)
                    station_name = _station_name(normalized)
                    station_address = _station_address(normalized)
                    match = _CUSTOMER_NO_RE.search(line)
                    customer_no = match.group(0) if match else ""
                    # A new station restarts the month forward-fill - the
                    # first row of a station always prints its own date, and
                    # inheriting across a station boundary would be wrong.
                    current_date = ""
                    continue

                tokens = line.split()
                if not _looks_like_data_row(tokens):
                    continue
                if not station_name:
                    # A data row before any station header - structurally
                    # impossible in this format, so flag rather than guess
                    # which station it belongs to.
                    records.append(_unparsed_record(path.name, page_number, line))
                    continue

                parsed = _parse_data_row(tokens)
                if parsed is None:
                    records.append(_unparsed_record(path.name, page_number, line))
                    continue

                if parsed["date"]:
                    current_date = parsed["date"]

                waste_type, map_note = _map_waste_type(parsed["material"])

                record = empty_record(path.name)
                record["document_type"] = DOCUMENT_TYPE_KMM_SUMMARY_REPORT
                record["site"] = station_name
                record["waste_type"] = waste_type
                record["quantity"] = parsed["weight"]
                record["unit"] = KMM_UNIT
                record["supplier_or_carrier"] = KMM_SUPPLIER
                record["date"] = _month_label(current_date)
                record["certificate_or_reference"] = customer_no
                record["page_number"] = page_number
                record["confidence"] = "גבוהה"

                notes = [_BASE_NOTE]
                if station_address:
                    notes.append(f"כתובת: {station_address}")
                if parsed["product"]:
                    notes.append(parsed["product"])
                if map_note:
                    notes.append(map_note)
                    record["confidence"] = "בינונית"
                if not current_date:
                    notes.append("תאריך חסר בדוח")
                    record["confidence"] = "בינונית"
                record["notes"] = " | ".join(notes)

                records.append(record)
    finally:
        doc.close()
    return records
