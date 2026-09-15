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

## RTL text extraction: line breaks from pdfium, character ORDER from geometry

A PDF stores glyphs positioned on the page, not logical text. For Hebrew that
means the extracted character stream is in neither logical nor reliably visual
order, and punctuation lands beside the wrong neighbour: `בע"מ` came out as
`מ"בע`, `ו' סגור` as `ו סגור'`, `ת"א` as `א"ת`.

An earlier version reversed each line's *tokens* to compensate. That fixed
word order but could not fix anything inside or between tokens, and it also
made marker matching fragile (a two-word label like `שם תחנה:` arrives with
the colon wedged in the middle, `שם :תחנה`, so matching `"שם תחנה"` silently
never fired - on the first real run that flagged all 68 rows as unparsed).

`_logical_lines()` replaces it with real geometry:

  1. **Line segmentation comes from pdfium's own text stream** - its `\r\n`
     boundaries are reliable and give exactly the document's 58/40 lines.
     Reconstructing lines by clustering characters on their y coordinate was
     tried first and is *not* reliable here: the characters of one visual line
     spread over more than 5 points, which split data rows into fragments and
     dropped the spaces between words.
  2. **Character order within a line comes from each character's x
     coordinate** (`get_charbox`), which is the true visual left-to-right
     order.
  3. **Visual -> logical** by reversing the line, then re-reversing each
     left-to-right run (`_LTR_RUN`: Latin, digits, and the separators inside
     them). Without step 3, `KM103527`, `1,140` and `20/08/26` would come out
     backwards; with it, RTL and LTR runs both read correctly.

The result is genuine logical text: a data row reads
`ינו-2026 קרטון 1022 איסוף קרטון לפי קוב 115` - the document's own column
order (תאריך, סוג החומר, מק"ט, תאור מוצר, משקל) - so fields are read by
position from the ends inward rather than by unreversing anything.

This is the only place in the project that extracts PDF *text*; every other
document type is rasterized and read by the vision model (see extractor.py),
so no other field is affected by any of this.

## The other parsing trap: a date is printed only when it changes

Consecutive rows in the same month leave the date cell empty - 7 of the 68
rows in the real file do. The month is forward-filled from the last row that
had one, reset at each new station. Without that, those 7 rows would silently
get a blank תאריך.

A station block also continues across a page break with no repeated header,
so station state deliberately lives outside the page loop in
parse_kmm_report() - see the comment there.

Because every one of these is a silent-corruption risk rather than a crash,
any data line that fails to parse becomes a flagged record carrying its raw
text (`_unparsed_record`) instead of being skipped - a dropped row in a
68-row report is invisible.
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
# A month cell in logical order: "ינו-2026". (Before the geometry-based
# extraction below it arrived as "2026-ינו" and had to be un-swapped.)
_DATE_RE = re.compile(r"^[֐-׿]{3}-\d{4}$")
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


# Runs that read left-to-right even inside an RTL line: Latin letters, digits,
# and the separators that appear *inside* such a run ("KM103527", "1,140",
# "20/08/26", "18:24", "מ-7"'s digit part). Reversing a line flips these too,
# so each one is flipped back - see _visual_to_logical().
_LTR_RUN = re.compile(r"[A-Za-z0-9]+(?:[./:,\-][A-Za-z0-9]+)*")


def _visual_to_logical(visual: str) -> str:
    """Visual (left-to-right as painted) -> logical reading order.

    Reverses the line, which is correct for a base-RTL line, then restores
    each left-to-right run so embedded numbers and Latin ids don't come out
    backwards.
    """
    return _LTR_RUN.sub(lambda m: m.group(0)[::-1], visual[::-1])


def _logical_lines(page) -> List[str]:
    """The page's lines in logical reading order - see this module's docstring.

    Line breaks come from pdfium's own text stream; the order of characters
    within each line comes from their x coordinates. A character whose box
    can't be read (never seen on this format, but possible for an odd
    embedded font) leaves that line in stream order rather than dropping it.
    """
    textpage = page.get_textpage()
    lines: List[List[tuple]] = []
    current: List[tuple] = []
    for index in range(textpage.count_chars()):
        char = textpage.get_text_range(index, 1)
        if char in ("\r", "\n"):
            if current:
                lines.append(current)
                current = []
            continue
        if not char:
            continue
        try:
            x = textpage.get_charbox(index)[0]
        except Exception:
            x = None
        current.append((x, char))
    if current:
        lines.append(current)

    out = []
    for line in lines:
        if any(x is None for x, _c in line):
            visual = "".join(c for _x, c in line)
        else:
            visual = "".join(c for _x, c in sorted(line, key=lambda t: t[0]))
        out.append(_visual_to_logical(visual).strip())
    return out


def _looks_like_data_row(tokens: List[str]) -> bool:
    """A data row, in logical order, ENDS with the weight and contains a
    separate integer מק"ט before it.

    Both conditions are needed: the report's own column header
    ("תאריך סוג החומר מק\"ט תאור מוצר משקל") and its parameter line both
    contain integers, but neither ends in one.
    """
    if len(tokens) < 3 or not _WEIGHT_RE.fullmatch(tokens[-1]):
        return False
    return any(_SKU_RE.fullmatch(t) for t in tokens[:-1])


def _field_between(line: str, start_marker: str, stop_markers: tuple) -> str:
    """The text after `start_marker` up to whichever of `stop_markers` comes
    first. Field boundaries are matched on the labels, not on commas: the
    values themselves contain commas, and in this format the labels are the
    only dependable delimiter.
    """
    if start_marker not in line:
        return ""
    text = line.split(start_marker, 1)[1]
    for stop in stop_markers:
        text = text.split(stop, 1)[0]
    return text.strip(" ,:")


def _station_name(line: str) -> str:
    """The station's full name as printed, e.g.
    "מקורות מרחב מרכז אחיסמך מ-8 עד 15".
    """
    return _field_between(line, _STATION_MARKER, ("כתובת", "קוד סוג"))


def _station_address(line: str) -> str:
    """The station's address, for the notes column (per spec, "הכתובת של
    התחנה אם קיימת בדו"ח, לצורך הקשר"). Empty for a station that has none -
    one of the 16 in the real file genuinely doesn't.
    """
    return _field_between(line, _ADDRESS_MARKER, ("קוד סוג", "תאור סוג", "פרמטר"))


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
    """Splits one logical-order data row into the report's five columns.

    Logical order is `date? | material | מק"ט | product description | weight`.
    Anchored from the ends inward - the weight is the last token, the date
    (when present) the first, the מק"ט the first bare integer between them -
    which needs no per-column x coordinates and tolerates the multi-word
    material ("נייר לבן") and product ("איסוף קרטון לפי קוב") names.
    """
    if not _looks_like_data_row(tokens):
        return None
    weight = tokens[-1]
    rest = list(tokens[:-1])

    date = ""
    if rest and _DATE_RE.fullmatch(rest[0]):
        date = rest[0]
        rest = rest[1:]

    sku_index = next((i for i, t in enumerate(rest) if _SKU_RE.fullmatch(t)), None)
    if sku_index is None:
        return None

    return {
        "weight": weight,
        "sku": rest[sku_index],
        "material": " ".join(rest[:sku_index]).strip(),
        "product": " ".join(rest[sku_index + 1:]).strip(),
        "date": date,
    }


def _month_label(raw_date: str) -> str:
    """The printed month cell as display text: "ינו-2026" -> "ינו 2026".

    Deliberately NOT converted to a real date: per spec the תאריך column
    keeps the report's own free-text month ("כפי שמופיע במקור ... לא להמיר
    לפורמט תאריך מדויק"), which also means derive.derive_year_month() will
    not parse it and year/month stay blank on these records.
    """
    return raw_date.replace("-", " ").strip() if raw_date else ""


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
        logical = _logical_lines(doc[0])
    except Exception:
        return False
    finally:
        doc.close()

    text = raw + "\n" + "\n".join(logical)
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
            page_number = page_index + 1

            for line in _logical_lines(doc[page_index]):
                if not line:
                    continue

                if _STATION_MARKER in line:
                    station_name = _station_name(line)
                    station_address = _station_address(line)
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
