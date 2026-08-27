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
        "date",
        "תאריך הפינוי/העמסה, בפורמט DD/MM/YYYY. אם לא ניתן לזהות בבירור - השאר ריק ואל תנחש. "
        "(שדה זה משמש רק לגזירת שנה/חודש בקוד ואינו מוצג כעמודה בפני עצמו).",
        "string",
    ),
    (
        "waste_type",
        "סוג הפסולת - יש לבחור אך ורק מתוך הרשימה הסגורה: "
        + ", ".join(WASTE_TYPES)
        + ". התאם קטגוריה רק אם סוג הפריט כתוב או מתואר במפורש בתעודה עצמה בצורה "
        + "התואמת ישירות לאחת מהאפשרויות. אל תסווג לפי שם הקובץ, סוג העסק המנפיק, "
        + "או ניחוש כללי מהקשר - אם הפריט לא תואם באופן ישיר וברור לאף אחת "
        + "מהאפשרויות, השאר ריק.",
        "string",
    ),
    (
        "quantity",
        "הכמות, מספר בלבד (ללא יחידת מידה), למשל 12.5. אם לא ניתן לזהות - השאר ריק.",
        "string",
    ),
    (
        "unit",
        'יחידת המידה כפי שהיא מופיעה בתעודה (למשל ק"ג, טון, מ"ק, ליטר, יחידות).',
        "string",
    ),
    (
        "site",
        "שם האתר/היחידה שממנה נאספה הפסולת.",
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
        "שם החברה/המערכת שהנפיקה את התעודה, למשל VERIDIS, אמניר, Ziv Metals, GRI.",
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
