"""Single source of truth for the certificate schema.

Two lists matter here, and each is edited independently:

  FIELD_DEFS    - what Claude extracts from the certificate image/PDF. Drives
                  the tool_use JSON schema built in extractor.py and the
                  record dict keys used throughout the pipeline.
                  Each entry is (field_name, description_for_model, json_type).

  EXCEL_COLUMNS - the exact column order and Hebrew headers written to the
                  output spreadsheet by excel_writer.py. Includes columns
                  derived after extraction (year/month from date, via
                  derive.py; the source filename, added by the pipeline) that
                  aren't part of FIELD_DEFS. Each entry is
                  (record_key, hebrew_header).
"""

WASTE_TYPES = [
    "גזם",
    "קרטונים",
    "נייר",
    "אריזות",
    "אלקטרוניקה",
    "פסולת גושית",
    "פסולת בנין",
    "עודפי חפירה",
    "מתכות",
]

# The explicit "none of the above fit" value the model is instructed to
# return for waste_type instead of guessing or leaving it blank - see
# FIELD_DEFS below. Deliberately NOT a member of WASTE_TYPES itself (it's not
# a real waste category), and deliberately NOT added as a named row in
# excel_writer.py's summary-sheet pivot: that sheet's existing "לא מסווג / לא
# מזוהה" catch-all row is pure arithmetic (total minus the known types) and
# already captures any waste_type that isn't one of the 9 real ones,
# regardless of whether the cell is blank, this exact sentinel, or something
# else entirely - see excel_writer.py's own note on why that's deliberate.
UNCLASSIFIED_WASTE_TYPE = "לא מסווג/לא מזוהה"

# A PDF batch sometimes mixes in pages that aren't waste certificates at all
# (an internal billing/credit note, a work order, a blank page) - see
# FIELD_DEFS's document_type entry. Classified this way, in the same tool
# call as everything else, not a separate pre-filtering API call: cheaper,
# and the model already has the page in front of it either way. Kept out of
# EXCEL_COLUMNS entirely (app.pipeline.process_files routes anything that
# isn't DOCUMENT_TYPE_CERTIFICATE into its own "skipped" list, never into the
# records that reach excel_writer.py) - see app/pipeline.py.
DOCUMENT_TYPE_CERTIFICATE = "תעודת פינוי"
DOCUMENT_TYPE_OTHER = "לא תעודת פינוי"
DOCUMENT_TYPES = [DOCUMENT_TYPE_CERTIFICATE, DOCUMENT_TYPE_OTHER]

REGIONS = ["צפון", "דרום", "מרכז", "מטה"]

CONFIDENCE_LEVELS = ["גבוהה", "בינונית", "נמוכה"]

HEBREW_MONTHS = [
    "ינואר",
    "פברואר",
    "מרץ",
    "אפריל",
    "מאי",
    "יוני",
    "יולי",
    "אוגוסט",
    "ספטמבר",
    "אוקטובר",
    "נובמבר",
    "דצמבר",
]

FIELD_DEFS = [
    (
        "document_type",
        "קבע זאת ראשון, לפני כל שדה אחר: סוג המסמך - "
        + ", ".join(DOCUMENT_TYPES)
        + f". '{DOCUMENT_TYPE_CERTIFICATE}' - יש בו כמות/משקל פסולת שהוצאה מהאתר "
        + "(תעודת שקילה או תעודת הובלה). "
        + f"'{DOCUMENT_TYPE_OTHER}' - כל מסמך אחר שאינו תעודת פינוי בפועל, גם אם "
        + "יש בו התייחסות לפרויקט/אתר/עבודה - למשל תעודת חיוב/זיכוי פנימית, הזמנת "
        + f"עבודה, או עמוד ריק/לא רלוונטי. אם הסיווג הוא '{DOCUMENT_TYPE_OTHER}' - "
        + "עדיין יש למלא שדה זה ואת שדה ההערות (עם תיאור קצר של סוג המסמך, למשל "
        + "'מסמך חיוב/זיכוי פנימי'), אך אין טעם למלא את שאר השדות ואפשר להשאיר "
        + "אותם ריקים.",
        "string",
    ),
    (
        "date",
        "תאריך הפינוי/העמסה, בפורמט DD/MM/YYYY. יכול להופיע בכל מקום בעמוד - למעלה, "
        "למטה, בתוך טבלה או מחוצה לה - חפש בכל האזורים, לא רק במקום שבו הוא נוטה "
        "להופיע בתעודות אחרות. אם לא ניתן לזהות בבירור - השאר ריק ואל תנחש. "
        "(שדה זה משמש רק לגזירת שנה/חודש בקוד ואינו מוצג כעמודה בפני עצמו).",
        "string",
    ),
    (
        "waste_type",
        "סוג הפסולת - יש לבחור אך ורק מתוך הרשימה הסגורה: "
        + ", ".join(WASTE_TYPES)
        + f", או '{UNCLASSIFIED_WASTE_TYPE}'. התאם קטגוריה מהרשימה רק אם סוג הפריט "
        + "כתוב או מתואר במפורש בתעודה עצמה בצורה התואמת ישירות לאחת מהאפשרויות. "
        + "אל תסווג לפי שם הקובץ, סוג העסק המנפיק, או ניחוש כללי מהקשר. אם הפריט "
        + f"לא תואם באופן ישיר וברור לאף אחת מהאפשרויות - אל תשאיר ריק, החזר '{UNCLASSIFIED_WASTE_TYPE}' "
        + "ותאר בשדה ההערות במדויק את התיאור המקורי כפי שהוא מופיע בתעודה (למשל: "
        + "\"בתעודה נכתב 'סוללות'\") כדי שאפשר יהיה לשקול הוספת קטגוריה חדשה בעתיד.",
        "string",
    ),
    (
        "quantity",
        "הכמות, מספר בלבד (ללא יחידת מידה), למשל 12.5. אין ניסוח קבוע אחד לשדה הזה "
        "- הוא עשוי להופיע תחת כותרות שונות (כמות לחיוב, משקל נטו, משקל טרה/ברוטו, "
        "כמות ק\"ג/טון וכו') תלוי בספק. חפש לפי המשמעות, לא לפי הניסוח המדויק: "
        "אם מופיעים כמה ערכים, קח את הכמות הרלוונטית לחיוב - בדרך כלל 'כמות לחיוב' "
        "אם קיימת, ואחרת 'משקל נטו' (לא ברוטו/טרה - נטו הוא כמות הפסולת בפועל). "
        "זהו השדה הכי קריטי לדיוק בכל התעודה - לפני שאתה קובע את הערך הסופי, קרא "
        "כל ספרה בנפרד ובדוק אותה פעם שנייה (בפרט ספרות שקל להתבלבל ביניהן בכתב "
        "יד: 0/6/8, 1/7, 3/8, 4/9), וודא שמיקום הנקודה העשרונית ומספר הספרות "
        "תואמים בדיוק את מה שכתוב בתעודה. אם אין בתעודה מספר מדויק אלא רק תיאור "
        "מילולי של הכמות (למשל תעודת הובלה שכתוב בה 'עמוסה 3 מכולות' בלי משקל) - "
        "השאר שדה זה ריק אך כתוב את התיאור המילולי בשדה ההערות, כדי שהמידע לא "
        "יאבד לגמרי. אם אין שום אינדיקציה לכמות - השאר ריק ואל תנחש.",
        "string",
    ),
    (
        "unit",
        'יחידת המידה כפי שהיא מופיעה בתעודה (למשל ק"ג, טון, מ"ק, ליטר, יחידות).',
        "string",
    ),
    (
        "site",
        "שם האתר/היחידה שממנה נאספה הפסולת. יכול להופיע בכל מקום בעמוד (למעלה, "
        "למטה, בתוך טבלה או מחוצה לה) - חפש בכל האזורים, לא רק במקום הצפוי.",
        "string",
    ),
    (
        "region",
        "המרחב שאליו שייך האתר - "
        + ", ".join(REGIONS)
        + ". מלא רק אם המרחב כתוב במפורש בתעודה, או אם שם האתר הוא ללא כל ספק "
        + "וללא צורך בהסקה מזוהה אצלך כשייך לאותו מרחב באופן חד-משמעי. אם נדרשת "
        + "הסקה, השערה, או שאתה לא בטוח לחלוטין - השאר ריק.",
        "string",
    ),
    (
        "certificate_number",
        "מספר התעודה כפי שמופיע עליה.",
        "string",
    ),
    (
        "reference_type",
        "שם החברה/המערכת שהנפיקה את התעודה, למשל VERIDIS, אמניר, Ziv Metals, GRI. "
        "יכול להופיע בכל מקום בעמוד (בלוגו, בכותרת, בתחתית העמוד וכו') - חפש בכל "
        "האזורים, לא רק במקום הצפוי. אם הפורמט של התעודה לא מוכר ולא דומה לשום "
        "תעודה שראית בעבר - התייחס אליה באותה רצינות כמו לפורמט מוכר, אל תוותר "
        "ואל תשאיר שדות ריקים רק בגלל שהפורמט חדש.",
        "string",
    ),
    (
        "confidence",
        "הערכה עצמית שלך לגבי איכות הקריאה של התעודה כולה - "
        + ", ".join(CONFIDENCE_LEVELS)
        + ".",
        "string",
    ),
    (
        "notes",
        "הערה קצרה בלבד - עד 4-5 מילים, לא משפט מלא - המציינת רק את הבעיה: "
        "כתב יד לא קריא, שדה חסר בתעודה, מספר תעודה מטושטש וכו'. אם אין בעיה "
        "מיוחדת - השאר ריק.",
        "string",
    ),
]


def empty_record(filename: str = "", error: str = "") -> dict:
    """Builds a placeholder record matching FIELD_DEFS's shape, for a file or
    page that couldn't be extracted at all - flagged for manual review rather
    than silently dropped from the batch.

    Lives here, not in extractor.py or pipeline.py, so both of those modules
    can build one without importing from each other (extractor.py needs it
    for a single failed page; pipeline.py needs it for a whole file that
    couldn't even be opened/split).
    """
    record = {name: "" for name, *_ in FIELD_DEFS}
    record["source_file"] = filename
    record["year"] = ""
    record["month"] = ""
    record["confidence"] = "נמוכה"
    if error:
        record["notes"] = f"שגיאת עיבוד: {error}"
    return record


EXCEL_COLUMNS = [
    ("region", "מרחב"),
    ("site", "אתר/יחידה"),
    ("year", "שנה"),
    ("month", "חודש"),
    ("waste_type", "סוג הפסולת"),
    ("quantity", "כמות"),
    ("unit", "יחידת מידה"),
    ("reference_type", "סוג אסמכתא מצורפת"),
    ("certificate_number", "מספר תעודה"),
    ("confidence", "רמת ביטחון"),
    ("notes", "הערות"),
    ("source_file", "קובץ מקור"),
]
