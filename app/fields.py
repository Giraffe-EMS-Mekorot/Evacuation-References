"""Single source of truth for the certificate schema.

Three lists matter here, and each is edited independently:

  FIELD_DEFS    - what Claude extracts from the certificate image/PDF. Drives
                  the tool_use JSON schema built in extractor.py and the
                  record dict keys used throughout the pipeline.
                  Each entry is (field_name, description_for_model, json_type).

  EXCEL_COLUMNS - the exact column order and Hebrew headers written to the
                  main output-sheet by excel_writer.py. Includes columns
                  derived after extraction (certificate_or_reference, from
                  certificate_number/reference_number - see extractor.py)
                  that aren't part of FIELD_DEFS. Each entry is
                  (record_key, hebrew_header).

  INTERNAL_TRACKING_FIELDS / INTERNAL_TRACKING_SHEET_COLUMNS - fields that
                  ARE extracted (they're in FIELD_DEFS) but are deliberately
                  kept off the main sheet/table by default: vehicle/driver/
                  entry-exit-time/gross-tare-weight. Still collected and
                  saved, just surfaced behind the "הצג פרטים נוספים" toggle
                  in streamlit_app.py and on Excel's second "מעקב פנימי"
                  sheet (see excel_writer.py) rather than cluttering the
                  main view every reviewer sees. INTERNAL_TRACKING_FIELDS is
                  the plain 6-field set (used by the Streamlit toggle, which
                  already shows the identifying columns in the main table);
                  INTERNAL_TRACKING_SHEET_COLUMNS prepends a few identifying
                  columns to that same set, since the Excel sheet is
                  physically separate from the main sheet and needs to be
                  self-sufficient without relying on row order alone.
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

# A distinct, orthogonal classification from document_type above: whether
# THIS page is one half of a common two-document pattern seen across several
# (not one fixed) suppliers, where the same removal is documented twice -
# once as a computerized "תעודת שקילה/משלוח" (weighing/delivery slip, with
# ברוטו/טרה/נטו and its own מספר_תעודה) and once as a usually-handwritten
# "אישור ביצוע עבודה/הטמנה" (work/disposal completion approval, with a
# מספר_אסמכתא pointing back at the weighing slip's מספר_תעודה, a site/
# location, and a signature). See FIELD_DEFS's cert_role entry for the full
# field description sent to the model, and extractor.py's
# _detect_and_cross_check_pairs for how two adjacent pages carrying these
# two roles get cross-checked against each other. Deliberately identified by
# the model from each page's own structure/headers, never from a supplier
# name - the whole point is that this pattern isn't tied to one fixed
# vendor. Blank (not one of these two values) is the common case - most
# certificates aren't part of this pattern at all.
CERT_ROLE_WEIGHING = "תעודת שקילה/משלוח"
CERT_ROLE_COMPLETION = "אישור ביצוע עבודה/הטמנה"
# A third, unrelated cert_role value (2026-09-01): a "שטר מטען" (bill of
# lading, e.g. from a system like טיקטראק) that accompanies a shipment
# *before* it's actually weighed - its own weight/volume field explicitly
# contains "0" (not blank) for that reason, not because the quantity really
# is zero. Tagging this lets extractor.py's _flag_bill_of_lading_estimate
# refuse the literal "0" as quantity and look for a textual estimate
# instead, and lets pipeline._cross_check_bill_of_lading_estimates later
# supersede that estimate with a real weighing record's data if one turns up
# in the same batch (by vehicle+site+nearby date, not a matching id - see
# that function's docstring for why an id match can't be required here).
CERT_ROLE_BILL_OF_LADING_ZERO = "שטר מטען (משקל/נפח אפס - הערכה טרם שקילה)"
CERT_ROLES = [CERT_ROLE_WEIGHING, CERT_ROLE_COMPLETION, CERT_ROLE_BILL_OF_LADING_ZERO]

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
        + "עבודה, עמוד ריק/לא רלוונטי, או חשבונית (מסמך עם כמה שורות פריטים, מחירים, "
        + "עמודת מע\"מ, וסיכום כספי בתחתית - למשל מחברות כמו 'ק.מ.מ. מפעלי מחזור') - "
        + "חשבונית היא מסמך כספי, לא תעודת פינוי, גם אם מוזכר בה משקל או כמות של פריט. "
        + f"אם הסיווג הוא '{DOCUMENT_TYPE_OTHER}' - "
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
        "שדה זה מוצג כעמודה בפני עצמו, וגם משמש בקוד לבדיקת סבירות (תאריך עתידי/"
        "ישן מדי) ולהצלבה מול מסמך מקושר, אם קיים (ראו תפקיד_התעודה_בזוג).",
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
        "אם קיימים בתעודה גם שדות ברוטו/טרה נפרדים - מלא אותם בשדות ברוטו/טרה "
        "בנוסף לשדה זה, לא במקומו. "
        "זהו השדה הכי קריטי לדיוק בכל התעודה - לפני שאתה קובע את הערך הסופי, קרא "
        "כל ספרה בנפרד ובדוק אותה פעם שנייה (בפרט ספרות שקל להתבלבל ביניהן בכתב "
        "יד: 0/6/8, 1/7, 3/8, 4/9), וודא שמיקום הנקודה העשרונית ומספר הספרות "
        "תואמים בדיוק את מה שכתוב בתעודה. אם אין בתעודה מספר מדויק אלא רק תיאור "
        "מילולי של הכמות (למשל תעודת הובלה שכתוב בה 'עמוסה 3 מכולות' בלי משקל) - "
        "השאר שדה זה ריק אך כתוב את התיאור המילולי בשדה ההערות, כדי שהמידע לא "
        "יאבד לגמרי. יוצא מן הכלל: שטר מטען עם משקל/נפח שכתוב בו במפורש '0' (ראו "
        "הוראות שדה תפקיד_התעודה_בזוג) - שם '0' אינו כמות אמיתית אלא סימן שהשקילה "
        "טרם בוצעה, וכן יש לחלץ מספר מתוך תיאור טקסטואלי חלופי אם קיים (למשל "
        "'פינוי מכולה 12 מ\"ק') אל שדה זה ואל שדה יחידת_מידה, בדיוק כאילו הופיע "
        "בטבלת השקילה עצמה. אם אין שום אינדיקציה לכמות - השאר ריק ואל תנחש.",
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
        "מספר התעודה של המסמך הזה עצמו, כפי שמופיע עליו (בדרך כלל בכותרת העליונה, "
        "למשל 'תעודת משלוח מס' ...' או 'מספר: ...'). אל תמלא כאן מספר שמפנה למסמך "
        "אחר - זה שייך לשדה מספר_אסמכתא.",
        "string",
    ),
    (
        "reference_number",
        "מספר אסמכתא - ממלאים רק כשמופיע בתעודה שדה נפרד המפנה במפורש למספר "
        "התעודה של מסמך אחר (לרוב בתעודות מסוג 'אישור ביצוע עבודה/הטמנה', ראו "
        "תפקיד_התעודה_בזוג - שם מצוין מספר האסמכתא של תעודת השקילה/משלוח המתאימה "
        "לה). אל תמלא כאן את מספר התעודה של המסמך הזה עצמו - זה שייך לשדה "
        "מספר_תעודה. אם אין בתעודה שדה 'אסמכתא' נפרד שמפנה למסמך אחר - השאר ריק.",
        "string",
    ),
    (
        "cert_role",
        "סוג המבנה של התעודה, אם היא תואמת באופן מובהק לאחד משני דפוסים "
        "שמופיעים אצל חלק מהספקים (לא כולם), כזוג מסמכים סמוכים שמתעד את אותה "
        f"פעולת פינוי פעמיים: '{CERT_ROLE_WEIGHING}' - מסמך ממוחשב עם שדות "
        f"ברוטו/טרה/נטו ומספר תעודה משלו; '{CERT_ROLE_COMPLETION}' - מסמך "
        "(לרוב כתוב ביד) עם שדה 'אסמכתא' נפרד המפנה במפורש למספר תעודה של "
        "מסמך אחר, אתר/מיקום ביצוע העבודה, וחתימה. זהה לפי המבנה/הכותרות "
        "בפועל של התעודה הזו עצמה בלבד - לא לפי שם הספק (כדי שהזיהוי יעבוד גם "
        "אצל ספק שטרם נתקלת בו) ובלי קשר לשאלה אם יש בקובץ עמוד נוסף שמתאים לה "
        "- זו בדיקה נפרדת שנעשית בקוד, לא משהו שעליך לאמת בעצמך. כלומר: אם "
        "התעודה הזו לבדה מציגה את המבנה המובהק (למשל שדות ברוטו/טרה/נטו "
        f"ומספר תעודה) - סמן '{CERT_ROLE_WEIGHING}', גם אם אינך יודע אם יש "
        "עמוד מקושר. "
        f"ערך שלישי, בלתי קשור לזוג הזה: '{CERT_ROLE_BILL_OF_LADING_ZERO}' - "
        "'שטר מטען' (כפי שמופיע למשל במערכות כמו טיקטראק) שמלווה משלוח לפני "
        "שקילה בפועל, ולכן שדה המשקל/הנפח שלו בטבלה מכיל במפורש '0' (למשל "
        "'0 טון') ולא ריק - זה לא אומר שהכמות באמת אפס. סמנו ערך זה כשמזוהה "
        "תבנית כזו, כדי שהכמות תיבדק כהערכה טקסטואלית חלופית ולא כ-0 (ראו "
        "הוראות שדה כמות). "
        "אם אין לתעודה מבנה מובהק של אף אחד משלושת הדפוסים (המקרה הנפוץ ביותר) - "
        "השאר ריק.",
        "string",
    ),
    (
        "supplier_or_carrier",
        "שם החברה/המערכת שהנפיקה את התעודה, או שם חברת ההובלה (מוביל) המצוינת "
        "בה בנפרד אם קיימת, למשל VERIDIS, אמניר, Ziv Metals, GRI, מ.עבד הובלות. "
        "יכול להופיע בכל מקום בעמוד (בלוגו, בכותרת, בתחתית העמוד וכו') - חפש בכל "
        "האזורים, לא רק במקום הצפוי. אם הפורמט של התעודה לא מוכר ולא דומה לשום "
        "תעודה שראית בעבר - התייחס אליה באותה רצינות כמו לפורמט מוכר, אל תוותר "
        "ואל תשאיר שדות ריקים רק בגלל שהפורמט חדש.",
        "string",
    ),
    (
        "vehicle_number",
        "מספר הרכב שביצע את ההובלה, כפי שמופיע בתעודה. שדה למעקב פנימי בלבד "
        "(לא מוצג בעמודות הראשיות) - השאר ריק אם לא מצוין.",
        "string",
    ),
    (
        "driver_name",
        "שם הנהג שביצע את ההובלה, כפי שמופיע בתעודה. שדה למעקב פנימי בלבד "
        "(לא מוצג בעמודות הראשיות) - השאר ריק אם לא מצוין.",
        "string",
    ),
    (
        "entry_time",
        "שעת הכניסה של הרכב לאתר, כפי שמופיעה בתעודה (למשל 6:29). שדה למעקב "
        "פנימי בלבד (לא מוצג בעמודות הראשיות) - השאר ריק אם לא מצוינת.",
        "string",
    ),
    (
        "exit_time",
        "שעת היציאה של הרכב מהאתר, כפי שמופיעה בתעודה. שדה למעקב פנימי בלבד "
        "(לא מוצג בעמודות הראשיות) - השאר ריק אם לא מצוינת.",
        "string",
    ),
    (
        "gross_weight",
        "משקל ברוטו, מספר בלבד (ללא יחידת מידה) - מופיע בעיקר בתעודות שקילה "
        "ממוחשבות, יחד עם טרה ונטו. משמש בקוד לבדיקת סבירות מול הכמות (נטו) "
        "שדווחה (ברוטו פחות טרה אמור להיות שווה לנטו) - קרא כל ספרה בזהירות, "
        "כמו בשדה הכמות. השאר ריק אם לא מופיע בתעודה.",
        "string",
    ),
    (
        "tare_weight",
        "משקל טרה (משקל הרכב הריק), מספר בלבד (ללא יחידת מידה) - מופיע בעיקר "
        "בתעודות שקילה ממוחשבות, יחד עם ברוטו ונטו. משמש בקוד לבדיקת סבירות מול "
        "הכמות (נטו) שדווחה - קרא כל ספרה בזהירות, כמו בשדה הכמות. השאר ריק אם "
        "לא מופיע בתעודה.",
        "string",
    ),
    (
        "container_count",
        "מספר המכולות המצוין בתעודה, אם רלוונטי (למשל תעודת הובלה שכתוב בה "
        "'עמוסה 3 מכולות') - מספר בלבד. שדה זה מתווסף אוטומטית להערות ואינו "
        "עמודה נפרדת - השאר ריק אם לא רלוונטי/לא מצוין.",
        "string",
    ),
    (
        "container_type",
        "סוג/תיאור המכולה כפי שמופיע בתעודה (למשל 'מכולה פתוחה 8 קוב'), אם "
        "רלוונטי. שדה זה מתווסף אוטומטית להערות ואינו עמודה נפרדת - השאר ריק "
        "אם לא רלוונטי/לא מצוין.",
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
        "מיוחדת - השאר ריק. אל תעיר כאן על סבירות התאריך (למשל 'תאריך עתידי') - "
        "אין לך דרך לדעת מהו התאריך האמיתי של היום, ובדיקה כזו כבר נעשית "
        "באופן מדויק ואוטומטי בקוד לאחר החילוץ.",
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
    record["certificate_or_reference"] = ""
    record["confidence"] = "נמוכה"
    if error:
        record["notes"] = f"שגיאת עיבוד: {error}"
    return record


# Order/headers as of 2026-09-01 - right-to-left reading order in the RTL
# main sheet is LEFT-to-right through this list (first entry = column A =
# visually rightmost, per excel_writer.py's freeze_panes note), confirmed
# with the user as: מרחב | אתר/מקור | אתר קולט | תאריך | מספר תעודה/אסמכתא |
# סוג הפסולת | כמות | יחידת מידה | הערות | קובץ מקור.
#
# "confidence"/רמת ביטחון is deliberately NOT a main-sheet column any more -
# moved to INTERNAL_TRACKING_SHEET_COLUMNS below (see excel_writer.py's
# _write_summary_sheet, which now reads its confidence-breakdown formulas
# from that sheet instead of this one). The main sheet's row-level red/
# yellow confidence FILL (see excel_writer.write_records) is unaffected -
# it reads record["confidence"] in Python directly, not a written column -
# so a low/medium-confidence row is still visually flagged even without the
# text column spelling out why.
EXCEL_COLUMNS = [
    ("region", "מרחב"),
    ("site", "אתר/מקור"),
    ("supplier_or_carrier", "אתר קולט"),
    ("date", "תאריך"),
    ("certificate_or_reference", "מספר תעודה/אסמכתא"),
    ("waste_type", "סוג הפסולת"),
    ("quantity", 'כמות (נטו)'),
    ("unit", "יחידת מידה"),
    ("notes", "הערות"),
    ("source_file", "קובץ מקור"),
]

# Collected from every certificate (they're ordinary FIELD_DEFS entries) but
# kept off the main table/sheet by default - see this module's docstring.
# Used by streamlit_app.py's "הצג פרטים נוספים" toggle, which appends these
# to the already-visible main columns, so no identifying columns are
# repeated here.
INTERNAL_TRACKING_FIELDS = [
    ("vehicle_number", "מס' רכב"),
    ("driver_name", "שם נהג"),
    ("entry_time", "שעת כניסה"),
    ("exit_time", "שעת יציאה"),
    ("gross_weight", "משקל ברוטו"),
    ("tare_weight", "משקל טרה"),
]

# The Excel "מעקב פנימי" sheet's full column set (see excel_writer.py) - a
# few identifying columns prepended to INTERNAL_TRACKING_FIELDS, since that
# sheet is physically separate from the main one and needs to let a reviewer
# find the right row without relying on row order alone.
#
# "confidence" lives here (2026-09-01) rather than on the main sheet - see
# EXCEL_COLUMNS's own note above. Placed right after the identifying columns
# (source_file/certificate_or_reference/site), before the plain tracking
# fields, so it reads as "here's who/what this row is, and how sure we were
# about it" before the raw vehicle/time/weight data.
INTERNAL_TRACKING_SHEET_COLUMNS = [
    ("source_file", "קובץ מקור"),
    ("certificate_or_reference", "מספר תעודה/אסמכתא"),
    ("site", "אתר/מקור"),
    ("confidence", "רמת ביטחון"),
] + INTERNAL_TRACKING_FIELDS
