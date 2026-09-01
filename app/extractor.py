"""Sends certificate page images to Claude and returns their extracted fields.

A single-page image file (.jpg/.jpeg/.png) is one page. A PDF is split into
one image per page and each is extracted separately - never sent as one
multi-page request - for two reasons: Anthropic's API rejects an
oversized request outright (413 request_too_large, seen in practice against
a 105MB multi-page scan), and a multi-page PDF may contain several distinct
certificates that each need their own row, not one row for the whole file.

Trade-off worth knowing about: this always rasterizes PDFs to JPEG images,
even a clean single-page, computer-generated PDF with a real text layer that
Claude could otherwise read natively as a `document` block. That native path
is gone - every page now goes through the same downscale-and-compress step,
per the requirement that *every* image sent to the API be capped in size,
not just the multi-page case that motivated this. If that native-PDF fidelity
turns out to matter in practice, the fix is special-casing single-page PDFs
back to a `document` block, not re-litigating the multi-page split.
"""
from datetime import date
from pathlib import Path
from typing import List, Optional

import anthropic
import pypdfium2 as pdfium
from PIL import Image

from . import config, examples_library
from .derive import derive_year_month
from .excel_writer import parse_quantity
from .image_utils import MAX_IMAGE_BYTES, PDF_RENDER_SCALE, downscale_and_encode, encode_image_block
from .fields import (
    CERT_ROLE_BILL_OF_LADING_ZERO,
    CERT_ROLE_COMPLETION,
    CERT_ROLE_WEIGHING,
    CONFIDENCE_LEVELS,
    DOCUMENT_TYPE_OTHER,
    FIELD_DEFS,
    HEBREW_MONTHS,
    REGIONS,
    TOOL_NAME,
    UNCLASSIFIED_WASTE_TYPE,
    WASTE_TYPES,
    empty_record,
)
from .normalize import looks_reversed_hebrew

_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}

# Below this year, a certificate's date is treated as an OCR/garbled-digit
# misread rather than a real date - generous margin (10+ years) so it only
# catches clearly-wrong values, not just old certificates.
_MIN_PLAUSIBLE_YEAR = 2015

# Core fields (per the original spec) whose absence should always get a
# record flagged for manual review, regardless of what the model itself
# reports for רמת_ביטחון. year/month are derived together from the same
# "date" value (see derive_year_month) and aren't their own EXCEL_COLUMNS
# entry any more (the raw "date" field is, instead) - the check still keys
# off "year" internally, just labeled here to match the column a reviewer
# actually sees.
_CORE_FIELD_LABELS = [
    ("waste_type", "סוג הפסולת"),
    ("quantity", "כמות"),
    ("unit", "יחידת מידה"),
    ("site", "אתר/יחידה"),
    ("region", "מרחב"),
    ("year", "תאריך"),
]

# Tolerance (in the certificate's own unit, typically ק"ג) for the
# ברוטו-טרה=נטו arithmetic sanity check - see _flag_gross_tare_mismatch.
# Covers rounding on the certificate itself, not a real discrepancy.
_GROSS_TARE_TOLERANCE = 1.0
# Tolerance for cross-document quantity matching between a weighing
# certificate and its paired work-completion approval - see
# _cross_check_document_pair. Same rounding-margin reasoning as above.
_PAIR_QUANTITY_TOLERANCE = 1.0

_SYSTEM_PROMPT = (
    "אתה עוזר שמחלץ מידע מובנה מתעודות פינוי פסולת (PDF או תמונה סרוקה/מצולמת, "
    "חלקן ממוחשבות וחלקן כתובות בכתב יד). קרא את התעודה המצורפת בעיון והפעל את הכלי "
    f"{TOOL_NAME} עם הנתונים שחילצת. "
    "קבע קודם כל את סוג_המסמך (ראו הוראות השדה) - אם זה לא תעודת פינוי בפועל, אין "
    "טעם לחפש בו שדות של כמות/סוג פסולת, ומותר להשאיר אותם ריקים. "
    "אין פורמט קבוע אחד לתעודות האלה - ספקים שונים מנסחים ומסדרים שדות באופן שונה "
    "לגמרי זה מזה, ושדות עשויים להופיע בכל מקום בעמוד (ראו גם הוראות השדות "
    "הספציפיים). חפשו לפי המשמעות והתפקיד של כל שדה, לא לפי ניסוח או מיקום קבוע "
    "מראש - ואם התעודה בפורמט שלא נראה מוכר, התייחסו אליה באותה רצינות כמו לפורמט "
    "מוכר, לא כאל מקרה שאפשר לוותר עליו. "
    "מלא שדה רק אם הערך שלו כתוב או מופיע בבירור בתעודה עצמה - לא בהסקה מהקשר, "
    "לא משם הקובץ, ולא מדמיון לתעודות אחרות. "
    "בתעודה כתובה בכתב יד: התייחסו לכל שדה בנפרד לפי מידת הבהירות שלו - אם שדה "
    "מסוים כתוב בבירור (למשל תאריך קריא היטב) מלאו אותו בביטחון גבוה גם אם שדות "
    "אחרים באותה תעודה מטושטשים או לא ברורים; אל תוותרו על שדה שכן ניתן לקרוא רק "
    "בגלל ששדה אחר לא ניתן. ההפך גם נכון: אל תוותרו מראש על ניסיון לקרוא שדה רק כי "
    "הכתב יד קשה - קראו כמיטב יכולתכם ומלאו את מה שאתם כן מזהים בביטחון סביר. "
    "אם שדה כלשהו אינו ניתן לזיהוי ודאי, או שמילויו מצריך הסקה/ניחוש - השאר אותו "
    "כמחרוזת ריקה. שדה ריק תמיד עדיף על ניחוש שגוי. יוצא מן הכלל היחיד הוא "
    "סוג_הפסולת - ראו את הוראות השדה עצמו: שם אסור להשאיר ריק, ויש להחזיר "
    f"'{UNCLASSIFIED_WASTE_TYPE}' כשהפריט לא תואם לרשימה הסגורה. "
    "כלל אצבע לשאר השדות: אם אתה עומד לכתוב בשדה הערות שהערך שמילאת הוא 'הנחה', "
    "'הערכה' או 'ניחוש' - זהו סימן שהיה עליך להשאיר את השדה עצמו ריק ולתאר את חוסר "
    "הוודאות רק בהערות, לא למלא אותו. "
    "לפני שאתה קובע את שדה הכמות הסופי - זהו השדה הכי קריטי לדיוק בתעודה - עבור "
    "פעם נוספת על כל ספרה שקראת ווודא שלא התבלבלת בין ספרות דומות בכתב יד "
    "(0/6/8, 1/7, 3/8, 4/9) ושמיקום הנקודה העשרונית נכון; אין צורך לכתוב את תהליך "
    "הבדיקה הזה בשום שדה - רק לוודא אותו לפני שאתה עונה. אותה זהירות חלה גם על "
    "שדות ברוטו/טרה כשהם קיימים בתעודה - הם נבדקים בקוד מול הכמות (נטו) שדווחה. "
    "בחלק מהספקים (לא כולם) מופיעים זוגות של שני מסמכים סמוכים המתעדים את אותה "
    "פעולת פינוי: 'תעודת שקילה/משלוח' ממוחשבת עם ברוטו/טרה/נטו ומספר תעודה, "
    "ולידה 'אישור ביצוע עבודה/הטמנה' כתוב-יד עם מספר אסמכתא (המפנה למספר "
    "התעודה של תעודת השקילה), אתר/מיקום ביצוע, וחתימה. סמן בשדה cert_role את "
    "המבנה של העמוד הזה לפי המבנה/הכותרות שלו בפועל בלבד - לא לפי שם הספק, "
    "ובלי קשר לשאלה אם יש עמוד סמוך מתאים (זו בדיקה נפרדת שנעשית בקוד): גם אם "
    "אינך יודע אם קיים עמוד מקושר, אם התעודה הזו לבדה מציגה מבנה מובהק של אחד "
    "משני הסוגים - סמן אותו. ברוב התעודות אין מבנה כזה כלל - אז השאר ריק. "
    "תבנית שלישית, בלתי קשורה לזוג הזה: 'שטר מטען' (למשל ממערכות כמו טיקטראק) "
    "שמלווה משלוח *לפני* שקילה בפועל - שדה המשקל/הנפח שלו בטבלה מכיל במפורש "
    f"'0' (למשל '0 טון'), כי השקילה טרם בוצעה - לא כי הכמות באמת אפס. אם מזוהה "
    f"מסמך כזה - סמנו בשדה cert_role את הערך '{CERT_ROLE_BILL_OF_LADING_ZERO}', "
    "ואל תמלאו בשדה הכמות את ה-'0' מהטבלה. במקום זאת חפשו בשאר התעודה (למשל "
    "בשדה 'פרטים' או תיאור חופשי) הערכה טקסטואלית של הכמות - אם מופיע בה מספר "
    "עם יחידת מידה ברורה (למשל 'פינוי מכולה 12 מ\"ק') - חלצו את המספר לשדה "
    "הכמות ואת יחידת המידה לשדה יחידת_מידה, בדיוק כפי שהייתם עושים אילו המספר "
    "הזה הופיע בטבלת השקילה עצמה. אם אין בתעודה שום הערכה טקסטואלית עם מספר - "
    "השאירו את שדה הכמות ריק (אל תמלאו 0). "
    "בשדה רמת_ביטחון דווח את הערכתך לגבי איכות הקריאה של התעודה כולה. "
    "בשדה הערות כתוב הערה קצרה בלבד - עד 4-5 מילים, לא משפט מלא - שמציינת רק את "
    "הבעיה עצמה בתמציתיות, למשל: 'כתב יד לא קריא', 'שדה חסר בתעודה', 'מספר תעודה "
    "מטושטש'. אל תכתוב הסברים ארוכים על מה שעשית או איך הגעת לערך - רק את הבעיה. "
    "אל תעיר בשדה הערות על סבירות התאריך (למשל 'תאריך עתידי') - אין לך דרך לדעת "
    "מהו התאריך האמיתי של היום, ובדיקה כזו נעשית באופן מדויק ואוטומטי בקוד לאחר "
    "החילוץ; חלץ את התאריך כפי שהוא כתוב בתעודה בלבד, ללא הערכה משלך על סבירותו. "
    f"חריג: כשמחזירים '{UNCLASSIFIED_WASTE_TYPE}' בסוג_הפסולת, או '{DOCUMENT_TYPE_OTHER}' "
    "בסוג_המסמך, כן יש לכתוב בהערות תיאור מדויק (התיאור המקורי מהתעודה, או סוג "
    "המסמך בהתאמה) גם אם זה יותר מ-4-5 מילים."
)


def _build_tool_schema() -> dict:
    properties = {}
    required = []
    for name, description, json_type in FIELD_DEFS:
        properties[name] = {"type": json_type, "description": description}
        required.append(name)
    return {
        "name": TOOL_NAME,
        "description": "רישום השדות המובנים שחולצו מתעודת פינוי פסולת אחת.",
        "input_schema": {
            "type": "object",
            "properties": properties,
            "required": required,
        },
    }


def _flag_out_of_list_values(record: dict) -> None:
    """Downgrades confidence and appends a note when a closed-list field is out
    of range. UNCLASSIFIED_WASTE_TYPE is deliberately exempt here - it's the
    model correctly following instructions (see FIELD_DEFS's waste_type
    description), not a violation of the closed list; see
    _flag_unclassified_waste_type() for its own, milder handling.
    """
    issues = []
    if (
        record["waste_type"]
        and record["waste_type"] != UNCLASSIFIED_WASTE_TYPE
        and record["waste_type"] not in WASTE_TYPES
    ):
        issues.append(f"סוג_פסולת לא מזוהה: '{record['waste_type']}'")
    if record["region"] and record["region"] not in REGIONS:
        issues.append(f"מרחב לא מזוהה: '{record['region']}'")
    if record["confidence"] not in CONFIDENCE_LEVELS:
        record["confidence"] = record["confidence"] or "נמוכה"

    if issues:
        record["confidence"] = "נמוכה"
        note = "; ".join(issues)
        record["notes"] = f"{record['notes']} | {note}" if record["notes"] else note


def _flag_unclassified_waste_type(record: dict) -> None:
    """Ensures a record the model marked UNCLASSIFIED_WASTE_TYPE always
    surfaces for manual review (to judge whether a new category is needed),
    even if the model's own רמת_ביטחון claimed "גבוהה" - it read the
    certificate clearly, it just found nothing in the closed list to match,
    which is exactly the situation a human should weigh in on.

    Escalates to "בינונית", not "נמוכה" - unlike _flag_out_of_list_values
    (an actual violation of the closed-list instruction) this is the model
    correctly following instructions, so it doesn't warrant the same
    severity as a genuine error; _enforce_core_field_confidence's "already at
    least בינונית" rows are left alone, matching that function's own pattern.
    """
    if record.get("waste_type") != UNCLASSIFIED_WASTE_TYPE:
        return
    if record["confidence"] not in ("נמוכה", "בינונית"):
        record["confidence"] = "בינונית"
    note = "סוג פסולת לא מסווג - לשקול קטגוריה חדשה"
    if note not in (record.get("notes") or ""):
        record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _enforce_core_field_confidence(record: dict) -> None:
    """Forces confidence to at least "בינונית" whenever a core field is blank,
    no matter what the model self-reported for רמת_ביטחון - a certificate
    that's missing a core field should always surface for manual review.
    Only escalates (נמוכה is already "at least as concerned" and is left as-is).
    """
    missing_labels = [label for key, label in _CORE_FIELD_LABELS if not record.get(key)]
    if not missing_labels:
        return
    if record["confidence"] not in ("נמוכה", "בינונית"):
        record["confidence"] = "בינונית"
        note = f"שדות ליבה חסרים: {', '.join(missing_labels)}"
        record["notes"] = f"{record['notes']} | {note}" if record["notes"] else note


def _append_page_note(record: dict, page_index: int, page_total: int) -> None:
    note = f"עמוד {page_index} מתוך {page_total}"
    record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _flag_unreasonable_date(record: dict) -> None:
    """Flags a record whose derived year/month is implausible for a waste
    certificate: later than the current month (a certificate can't predate
    its own creation), or old enough to almost certainly be an OCR/garbled-
    digit misread (see _MIN_PLAUSIBLE_YEAR) rather than a real date.

    There was no date-sanity check anywhere in this pipeline before this -
    derive.py only validates that the string is a well-formed DD/MM/YYYY (or
    DD/MM/YY) date, never that the resulting date makes sense. This adds
    that check; it doesn't fix a pre-existing one.

    Deliberately does not attempt to correct the value (e.g. try swapping
    day/month on the theory the source used MM/DD) - genuinely ambiguous
    cases like 03/04/2026 can't be resolved without more context, and
    guessing would violate this pipeline's "never guess, flag for a human
    instead" policy used everywhere else (see _flag_out_of_list_values,
    _enforce_core_field_confidence).
    """
    year_str, month_name = record.get("year"), record.get("month")
    if not year_str or not month_name:
        return
    try:
        year = int(year_str)
        month = HEBREW_MONTHS.index(month_name) + 1
    except (ValueError, IndexError):
        return

    today = date.today()
    is_future = (year, month) > (today.year, today.month)
    is_implausibly_old = year < _MIN_PLAUSIBLE_YEAR
    if not (is_future or is_implausibly_old):
        return

    reason = "עתידי" if is_future else "ישן מדי"
    note = f"תאריך לא סביר ({reason}): {month_name} {year}"
    record["confidence"] = "נמוכה"
    record["notes"] = f"{record['notes']} | {note}" if record["notes"] else note


_REVERSAL_CHECK_FIELDS = [
    ("site", "אתר/מקור"),
    ("supplier_or_carrier", "אתר קולט"),
    ("driver_name", "שם נהג"),
]


def _flag_possibly_reversed_text(record: dict) -> None:
    """Flags (never auto-corrects - unlike normalize.NameNormalizer, this
    runs per-page with no "known names" list to safely correct against) a
    record whose site/supplier_or_carrier/driver_name looks like its Hebrew
    characters came out in reversed order - see
    normalize.looks_reversed_hebrew for the detection rule and the
    real-world vision-model failure mode (RTL glyphs read in left-to-right
    pixel order) this guards against. Downgrades like
    _enforce_core_field_confidence's pattern: escalates to at least
    "בינונית", never downgrades a row already at "נמוכה"/"בינונית" for
    another reason.
    """
    suspect_labels = [label for key, label in _REVERSAL_CHECK_FIELDS if looks_reversed_hebrew(record.get(key))]
    if not suspect_labels:
        return
    if record["confidence"] not in ("נמוכה", "בינונית"):
        record["confidence"] = "בינונית"
    note = f"יתכן שנקרא הפוך (RTL): {', '.join(suspect_labels)}"
    if note not in (record.get("notes") or ""):
        record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _flag_gross_tare_mismatch(record: dict) -> None:
    """Sanity-checks ברוטו - טרה = נטו whenever both gross_weight and
    tare_weight are present on a record, regardless of whether it's part of
    a weighing/completion document pair (see _cross_check_document_pair,
    which is a separate, pair-level check) - the arithmetic must hold on any
    single certificate that reports all three values.

    Only runs when the unit is blank or ק"ג (gross/tare are almost always
    reported in ק"ג on a weighing slip - see gross_weight/tare_weight's
    FIELD_DEFS description) - _GROSS_TARE_TOLERANCE is a fixed ק"ג margin,
    so comparing it against, say, טון would be meaningless.
    """
    if record.get("unit") not in ("", 'ק"ג'):
        return
    gross = parse_quantity(record.get("gross_weight"))
    tare = parse_quantity(record.get("tare_weight"))
    net = parse_quantity(record.get("quantity"))
    if gross is None or tare is None or net is None:
        return
    gap = (gross - tare) - net
    if abs(gap) <= _GROSS_TARE_TOLERANCE:
        return
    note = (
        f'פער בין ברוטו-טרה לנטו המדווח: {abs(gap):g} ק"ג '
        f"(ברוטו={gross:g}, טרה={tare:g}, נטו={net:g})"
    )
    record["confidence"] = "נמוכה"
    record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _flag_bill_of_lading_estimate(record: dict) -> None:
    """Ensures a record the model tagged CERT_ROLE_BILL_OF_LADING_ZERO (a
    pre-weighing שטר מטען whose own weight/volume field is an explicit '0' -
    see fields.py's cert_role description and the system prompt) always
    surfaces for manual review, even if every field on the document itself
    was read perfectly clearly - its quantity (if any) is, at best, a
    textual estimate extracted from free text, never an actual weighing.

    Escalates to "בינונית" like _flag_unclassified_waste_type's pattern
    (the model correctly followed instructions here; this isn't an error),
    never downgrading a row already at "נמוכה"/"בינונית" for another reason.
    May be superseded later at the batch level, across every file in this
    run - not just this one PDF's pages - if a real weighing record for the
    same vehicle/site/date turns up: see
    pipeline._cross_check_bill_of_lading_estimates, which then overwrites
    this record's quantity/unit and adds its own note on top of this one.
    """
    if record.get("cert_role") != CERT_ROLE_BILL_OF_LADING_ZERO:
        return
    if record["confidence"] not in ("נמוכה", "בינונית"):
        record["confidence"] = "בינונית"
    note = "הערכה משטר מטען - טרם נשקל בפועל"
    if note not in (record.get("notes") or ""):
        record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _append_container_note(record: dict) -> None:
    """Folds container_count/container_type into notes instead of leaving
    them as their own column - per spec, this pair of fields is meant to be
    "collected but appended to notes", not a main-sheet/tracking-sheet
    column (see fields.py's FIELD_DEFS descriptions for both).
    """
    count = (record.get("container_count") or "").strip()
    ctype = (record.get("container_type") or "").strip()
    if not count and not ctype:
        return
    if count and ctype:
        note = f"{count} מכולות ({ctype})"
    elif count:
        note = f"{count} מכולות"
    else:
        note = f"מכולה: {ctype}"
    existing = record.get("notes") or ""
    if note in existing:
        return
    record["notes"] = f"{existing} | {note}" if existing else note


def _normalize_reference(value: str) -> str:
    """Normalizes a certificate/reference id for equality comparison in
    _cross_check_document_pair - a weighing certificate's own
    certificate_number and a paired completion approval's reference_number
    should refer to the same real-world id, but formatting can differ
    (leading zeros, surrounding whitespace, case on an alphanumeric prefix
    like "MNM-142496"). Purely numeric ids are compared as integers (drops
    leading zeros); anything else is compared as a stripped, case-folded
    string.
    """
    text = (value or "").strip()
    if text.isdigit():
        return str(int(text))
    return text.casefold()


def _cross_check_document_pair(weighing: dict, completion: dict) -> None:
    """Cross-checks one detected (CERT_ROLE_WEIGHING, CERT_ROLE_COMPLETION)
    pair - see _detect_and_cross_check_pairs, which finds the pairs this is
    called on. Three independent checks, each skipped (not flagged) when
    either side is missing the relevant field, rather than guessing at a
    mismatch from incomplete data:

      1. Reference number: the weighing doc's own certificate_number should
         equal the completion doc's reference_number (its מספר_אסמכתא is
         meant to point back at the weighing certificate's own number).
      2. Quantity: both documents should report the same net weight, even
         if worded differently (see quantity's FIELD_DEFS description) -
         compared within _PAIR_QUANTITY_TOLERANCE for rounding.
      3. Date: both documents describe the same real-world pickup/disposal
         event, so their dates should match.

    On any mismatch, both records in the pair are downgraded to "נמוכה" with
    one note listing exactly what didn't match (per spec, e.g. "אסמכתא לא
    תואמת: תעודת שקילה=15077, אישור ביצוע=15078"). When every check that
    could run passed, nothing is added - per spec, a clean cross-check is
    the "high confidence, no special note needed" case, not something that
    needs its own confirmation note.
    """
    issues = []

    w_ref = (weighing.get("certificate_number") or "").strip()
    c_ref = (completion.get("reference_number") or "").strip()
    if w_ref and c_ref and _normalize_reference(w_ref) != _normalize_reference(c_ref):
        issues.append(f"אסמכתא לא תואמת: תעודת שקילה={w_ref}, אישור ביצוע={c_ref}")

    w_qty = parse_quantity(weighing.get("quantity"))
    c_qty = parse_quantity(completion.get("quantity"))
    if w_qty is not None and c_qty is not None and abs(w_qty - c_qty) > _PAIR_QUANTITY_TOLERANCE:
        issues.append(f"כמות לא תואמת: תעודת שקילה={w_qty:g}, אישור ביצוע={c_qty:g}")

    w_date = (weighing.get("date") or "").strip()
    c_date = (completion.get("date") or "").strip()
    if w_date and c_date and w_date != c_date:
        issues.append(f"תאריך לא תואם: תעודת שקילה={w_date}, אישור ביצוע={c_date}")

    if not issues:
        return
    note = "; ".join(issues)
    for record in (weighing, completion):
        record["confidence"] = "נמוכה"
        if note not in (record.get("notes") or ""):
            record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _detect_and_cross_check_pairs(records: List[dict]) -> None:
    """Scans one PDF's page records for the two-document weighing+completion
    pattern (see fields.CERT_ROLE_WEIGHING/CERT_ROLE_COMPLETION) and
    cross-checks any adjacent pair found, in either order (a weighing slip
    followed by its completion approval, or the reverse).

    Deliberately supplier-agnostic: cert_role is set by the model from each
    page's own structure/headers (see that field's FIELD_DEFS description),
    never from a vendor name, so this same logic covers any supplier using
    this pattern, including one not seen before.

    Only strictly adjacent pages are considered (matching the spec's "שני
    עמודים סמוכים") - a completion approval that arrives without a
    neighboring weighing certificate (or vice versa) is simply left alone,
    processed like any other standalone certificate: no cross-check, and no
    failure. Not applied across separate uploaded files - only within one
    PDF's own pages, which is the only case "adjacent pages" can mean.
    """
    for i in range(len(records) - 1):
        a, b = records[i], records[i + 1]
        role_a, role_b = a.get("cert_role"), b.get("cert_role")
        if role_a == CERT_ROLE_WEIGHING and role_b == CERT_ROLE_COMPLETION:
            _cross_check_document_pair(a, b)
        elif role_a == CERT_ROLE_COMPLETION and role_b == CERT_ROLE_WEIGHING:
            _cross_check_document_pair(b, a)


def _extract_page(image: Image.Image, client: anthropic.Anthropic, use_examples: bool = False) -> dict:
    """Sends one already-loaded page image to Claude and returns its
    extracted fields. The single-page core that both a plain image file and
    each page of a split PDF funnel through. Doesn't set source_file - the
    caller (extract_certificate_pages) does, since it knows the filename.

    use_examples=True prepends a small few-shot preamble before the real
    page's own message - 1-2 known-hard, known-correct example certificates
    with their ground-truth answers, curated under examples/ (see
    app.examples_library and examples/README.md) - to help the model read
    difficult handwriting more like it read those. A missing or empty
    example library degrades to zero examples, not an error (see
    examples_library.build_fewshot_messages()).

    **Default is False (2026-09 decision)**: this whole few-shot feature is
    built and ready but deliberately inactive for now - not pursued further
    per an explicit user decision - so it must add zero cost/latency to
    ordinary processing unless someone explicitly opts in with
    use_examples=True. False sends the exact same single-message request
    this pipeline always sent before this feature existed.
    """
    jpeg_bytes = downscale_and_encode(image)
    if len(jpeg_bytes) > MAX_IMAGE_BYTES:
        raise ValueError(
            f"התמונה גדולה מדי לשליחה גם אחרי דחיסה ({len(jpeg_bytes) / 1_000_000:.1f}MB)"
        )
    file_block = encode_image_block(jpeg_bytes)

    messages = list(examples_library.build_fewshot_messages()) if use_examples else []
    messages.append(
        {
            "role": "user",
            "content": [
                file_block,
                {"type": "text", "text": "חלץ את הנתונים מהתעודה המצורפת."},
            ],
        }
    )

    response = client.messages.create(
        model=config.MODEL_NAME,
        max_tokens=1024,
        system=_SYSTEM_PROMPT,
        tools=[_build_tool_schema()],
        tool_choice={"type": "tool", "name": TOOL_NAME},
        messages=messages,
    )

    tool_use = next((block for block in response.content if block.type == "tool_use"), None)
    if tool_use is None:
        raise RuntimeError("המודל לא החזיר קריאת כלי (tool_use)")

    record = {name: str(tool_use.input.get(name, "") or "").strip() for name, *_ in FIELD_DEFS}
    if record.get("document_type") == DOCUMENT_TYPE_OTHER:
        # Not a waste certificate at all (see FIELD_DEFS's document_type) -
        # none of the certificate-specific checks below make sense against a
        # billing note or work order, so skip them entirely rather than
        # flag a "problem" that isn't one. app.pipeline.process_files routes
        # this into its own "skipped" list, never into the normal records.
        record["year"], record["month"] = "", ""
    else:
        _flag_out_of_list_values(record)
        _flag_unclassified_waste_type(record)
        record["year"], record["month"] = derive_year_month(record["date"])
        _flag_unreasonable_date(record)
        # Display/fallback id for the "מספר תעודה/אסמכתא" column - most
        # certificates only ever have certificate_number; reference_number
        # is the fallback for one that instead labels its own id "אסמכתא"
        # (distinct from reference_number's cross-document-pairing role -
        # see _detect_and_cross_check_pairs).
        record["certificate_or_reference"] = (
            record.get("certificate_number") or record.get("reference_number") or ""
        )
        _flag_gross_tare_mismatch(record)
        _flag_bill_of_lading_estimate(record)
        _flag_possibly_reversed_text(record)
        _append_container_note(record)
        _enforce_core_field_confidence(record)
    # Not an EXCEL_COLUMNS field - excel_writer.py's row-writing only reads
    # keys it knows about, so this rides along harmlessly for callers (like
    # main.py's CLI) that never look at it. streamlit_app.py uses it to show
    # the original page next to a row for quick manual correction.
    record["_page_image_jpeg"] = jpeg_bytes
    return record


def extract_certificate_pages(
    path: Path, client: Optional[anthropic.Anthropic] = None, use_examples: bool = False
) -> List[dict]:
    """Extracts one record per page of `path` - a plain image file is one
    page; a PDF is split into one page per record. A single bad page (the
    model call fails, or the page is somehow still too large after
    downscaling) never drops the rest of the file's pages - it's recorded via
    fields.empty_record() with the exception text in its notes, matching the
    file-level "one bad file shouldn't kill the batch" policy one level down.

    Raises if `path` itself can't be opened/parsed at all (unsupported
    extension, corrupt PDF/image) - that's a whole-file failure with no page
    count to report against, left for the caller (app.pipeline.process_files)
    to catch exactly like before this per-page split existed.

    use_examples (default False - see _extract_page()'s own docstring) is
    passed straight through to it for every page.
    """
    client = client or anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    ext = path.suffix.lower()

    if ext == ".pdf":
        pdf = pdfium.PdfDocument(str(path))
        try:
            total = len(pdf)
            records = []
            for index in range(total):
                page = pdf[index]
                try:
                    image = page.render(scale=PDF_RENDER_SCALE).to_pil()
                    record = _extract_page(image, client=client, use_examples=use_examples)
                except Exception as exc:  # one bad page shouldn't drop the rest of the file
                    record = empty_record(error=str(exc))
                finally:
                    page.close()
                record["source_file"] = path.name
                if total > 1:
                    _append_page_note(record, index + 1, total)
                    # Lets excel_writer.py show exactly which page of a
                    # multi-page PDF a row came from (e.g. "קובץ.pdf (עמוד
                    # 7)") in the "קובץ מקור" column itself, not just buried
                    # in the notes text above - so a reviewer chasing a
                    # mistake can jump straight to the right page. Only set
                    # for a multi-page file, same gate as the note above -
                    # a single-page PDF's source_file needs no page marker.
                    record["page_number"] = index + 1
                records.append(record)
            _detect_and_cross_check_pairs(records)
            return records
        finally:
            pdf.close()

    if ext in _IMAGE_EXTENSIONS:
        image = Image.open(path)
        try:
            record = _extract_page(image, client=client, use_examples=use_examples)
        except Exception as exc:
            record = empty_record(error=str(exc))
        record["source_file"] = path.name
        return [record]

    raise ValueError(f"סיומת קובץ לא נתמכת: {ext}")
