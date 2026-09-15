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
                  kept off the main sheet/table by default: entry/exit-time
                  and gross/tare-weight (vehicle_number/driver_name moved
                  onto the main sheet 2026-09-03 - see EXCEL_COLUMNS's own
                  comment). Still collected and saved, just surfaced behind
                  the "הצג פרטים נוספים" toggle in streamlit_app.py and on
                  Excel's second "מעקב פנימי" sheet (see excel_writer.py)
                  rather than cluttering the main view every reviewer sees.
                  INTERNAL_TRACKING_FIELDS is the plain 4-field set (used by
                  the Streamlit toggle, which already shows the identifying
                  columns in the main table);
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
# A fourth cert_role value (2026-09-15), again unrelated to the
# WEIGHING/COMPLETION pair above: a half-empty, hand-filled "תעודת שקילה"
# form that accompanies a computerized תעודת משלוח for the same removal.
#
# **The decisive sign is that the weighing table carries no printed values**
# (נטו/טרה/מס' רכב/תאריך/שעה/משקל cells empty or handwritten). That is exactly
# what separates it from CERT_ROLE_WEIGHING, whose ברוטו/טרה/נטו table IS
# computer-printed and full - and it's page-local, which matters: a first
# version of this description asked the model to notice that the header
# company name "appears as אתר קולט on ordinary source documents", which is
# knowledge it cannot have from one page. It returned cert_role blank on a
# real, textbook-matching document (input/תעודת_משלוח_102050.png, live run
# 2026-09-15) - the same failure mode CLAUDE.md records for CERT_ROLE_WEIGHING
# on 2026-08-30, for the same root cause. Every sign is now judgeable from the
# page in front of the model, and the description says outright that
# confirming a matching document exists is not its job.
#
# Deliberately NOT keyed off the printed title text: the real document's
# header misprints "תעודת משלוח מס'" while being a weighing certificate, so
# the description warns about that case explicitly.
#
# The ONLY field extracted from a page tagged this way is its own printed
# header number (certificate_number) - see extractor._strip_weighing_certificate
# for the code that enforces that, and pipeline._cross_check_weighing_certificates
# for the batch-level matching of that number against a תעודת משלוח's
# reference_number. Such a page never becomes a row of its own on the main
# sheet.
CERT_ROLE_WEIGHING_CERTIFICATE = "תעודת שקילה נלווית (מספר מודפס בכותרת)"
CERT_ROLES = [
    CERT_ROLE_WEIGHING,
    CERT_ROLE_COMPLETION,
    CERT_ROLE_BILL_OF_LADING_ZERO,
    CERT_ROLE_WEIGHING_CERTIFICATE,
]

# Record keys written by pipeline._cross_check_weighing_certificates onto a
# תעודת משלוח record once it's been matched 1:1 with a separate
# CERT_ROLE_WEIGHING_CERTIFICATE page in the same batch. They hold the
# *other* document's file/page so excel_writer.py can render a second,
# independently-clickable hyperlink next to the row's own "קובץ מקור" link
# (see WEIGHING_SOURCE_FILE_COLUMN below). Named here rather than inline in
# those two modules so the producer and the consumer can't drift.
WEIGHING_SOURCE_FILE_KEY = "weighing_source_file"
WEIGHING_PAGE_NUMBER_KEY = "weighing_page_number"
# Set on a CERT_ROLE_WEIGHING_CERTIFICATE record itself (not on the תעודת
# משלוח) once it has been consumed by a 1:1 match - excel_writer.py lists
# only the UNmatched ones on the "מעקב פנימי" sheet, so a weighing
# certificate that found its delivery certificate isn't also reported as an
# orphan.
WEIGHING_MATCHED_KEY = "_weighing_matched"

REGIONS = ["צפון", "דרום", "מרכז", "מטה"]

CONFIDENCE_LEVELS = ["גבוהה", "בינונית", "נמוכה"]

# The forced tool-use name extractor.py's real extraction calls use (see its
# _build_tool_schema()) - kept here, not privately in extractor.py, so
# app.examples_library can fabricate a matching assistant tool_use block for
# a few-shot example without importing from extractor.py itself (which
# imports FROM examples_library to prepend those messages - see that
# module's docstring on why that would otherwise be a circular import).
TOOL_NAME = "record_certificate_data"

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
        f"תבנית רביעית, נפרדת ומובהקת: '{CERT_ROLE_WEIGHING_CERTIFICATE}' - "
        "טופס 'תעודת שקילה' ריק-למחצה שממולא ביד, ומגיע כמסמך נלווה לתעודת "
        "משלוח ממוחשבת. הסימן המכריע, וזה שצריך להכריע אצלך: **טבלת השקילה "
        "בגוף המסמך (עמודות כמו נטו / טרה / מס' רכב / תאריך / שעה / משקל) "
        "אינה מכילה ערכים מודפסים** - התאים בה ריקים לגמרי, או שיש בהם כתב "
        "יד. סימנים תומכים, כולם נראים בעמוד הזה עצמו: מספר סידורי מודפס "
        "בכותרת (בניגוד לשאר המסמך שממולא ביד); כותרת/לוגו של חברה (האתר "
        "הקולט או המוביל); שורות ריקות למילוי ביד (המזמין, כתובת, מ-, ל-, "
        "מכונית מס', שעת יציאה); וחתימות ידניות בתחתית. "
        "אזהרת טעות נפוצה: **הכותרת המודפסת עשויה לומר 'תעודת משלוח מס''** - "
        "אל תיתן לזה להטעות אותך. מסמך שטבלת השקילה שלו ריקה/כתובה ביד הוא "
        f"'{CERT_ROLE_WEIGHING_CERTIFICATE}' גם כשכתוב בכותרתו 'תעודת משלוח'. "
        "הקביעה היא לפי המבנה בפועל, לא לפי מה שהכותרת מצהירה. "
        f"ההבדל מ-'{CERT_ROLE_WEIGHING}': שם טבלת ברוטו/טרה/נטו **מודפסת "
        "ומלאה בערכים ממוחשבים**; כאן היא ריקה או ידנית. אם הערכים מודפסים - "
        f"זה '{CERT_ROLE_WEIGHING}' (או תעודה רגילה), לא הערך הזה. "
        f"ההבדל מ-'{CERT_ROLE_COMPLETION}': שם קיים שדה 'אסמכתא' נפרד המפנה "
        "למספר תעודה של מסמך אחר; כאן המספר המודפס בכותרת הוא המספר של "
        "המסמך הזה עצמו, ואין בו שדה אסמכתא. "
        "קבע זאת לפי הצורה של העמוד הזה בלבד, ובלי קשר לשאלה אם קיים מסמך "
        "אחר שמתאים לו - אינך צריך (ואינך יכול) לאמת שקיימת תעודת משלוח "
        "תואמת, וההצלבה ביניהן נעשית בקוד בשלב נפרד. אם העמוד הזה לבדו מציג "
        "את המבנה הזה - סמן את הערך, גם אם אין לך מושג אם יש לו זוג. "
        f"כשאתה מסמן '{CERT_ROLE_WEIGHING_CERTIFICATE}' - חלץ מהמסמך הזה אך "
        "ורק את המספר המודפס בכותרת, אל שדה מספר_תעודה. אל תנסה לחלץ ממנו "
        "שום שדה אחר (תאריך, שם נהג, מספר רכב, משקל, אתר וכו') גם אם חלקם "
        "קריאים בפועל - השאר את כולם ריקים. הנתונים לשורה בטבלה נלקחים תמיד "
        "מתעודת המשלוח הממוחשבת, לא מכאן. "
        "אם אין לתעודה מבנה מובהק של אף אחד מארבעת הדפוסים (המקרה הנפוץ ביותר) - "
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
        "מספר הרכב (מספר רישוי) שביצע את ההובלה, כפי שמופיע בתעודה. נפוץ בשטרי "
        "מטען/אישורי הובלה, ולעיתים חסר לגמרי בתעודת שקילה פשוטה - השאר ריק אם "
        "לא מצוין, אל תנחש ואל תשלים מהקשר.",
        "string",
    ),
    (
        "driver_name",
        "שם הנהג שביצע את ההובלה, כפי שמופיע בתעודה. נפוץ בשטרי מטען/אישורי "
        "הובלה, ולעיתים חסר לגמרי בתעודת שקילה פשוטה - השאר ריק אם לא מצוין, "
        "אל תנחש ואל תשלים מהקשר.",
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


# Order/headers as of 2026-09-03 - right-to-left reading order in the RTL
# main sheet is LEFT-to-right through this list (first entry = column A =
# visually rightmost, per excel_writer.py's freeze_panes note), confirmed
# with the user as: מרחב / יחידה ראשית | יחידה / אתר מקור | סוג פסולת |
# תאריך האיסוף | כמות מדווחת | יחידת מידה | מספר אסמכתא | מספר רכב |
# שם הנהג | אתר קולט | הערות | קובץ מקור.
#
# vehicle_number/driver_name moved onto the main sheet this same change -
# previously internal-tracking-only (see INTERNAL_TRACKING_FIELDS below,
# and excel_writer.py's _LEGACY_HEADER_ALIASES for the older column
# labels this rename supersedes).
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
    ("region", "מרחב / יחידה ראשית"),
    ("site", "יחידה / אתר מקור"),
    ("waste_type", "סוג הפסולת"),
    ("date", "תאריך האיסוף"),
    ("quantity", "כמות מדווחת"),
    ("unit", "יחידת מידה"),
    ("certificate_or_reference", "מספר אסמכתא"),
    ("vehicle_number", "מספר רכב"),
    ("driver_name", "שם הנהג"),
    ("supplier_or_carrier", "אתר קולט"),
    ("notes", "הערות"),
    ("source_file", "קובץ מקור"),
]

# A main-sheet column that is deliberately NOT part of EXCEL_COLUMNS above
# (2026-09-15). It holds the second of the two hyperlinks a cross-matched row
# gets - one to the תעודת משלוח page, one to the separate תעודת שקילה page
# (see CERT_ROLE_WEIGHING_CERTIFICATE and
# pipeline._cross_check_weighing_certificates). Two independently-clickable
# links can't live in one cell: Excel supports exactly one hyperlink per
# cell, so "file_a | file_b" as a single joined string can only ever link to
# one of them.
#
# Kept out of EXCEL_COLUMNS on purpose, rather than added as an ordinary
# entry: streamlit_app.py derives its editable results table's columns from
# EXCEL_COLUMNS, so an entry here would silently add a column to that table
# too - and this change was explicitly scoped to leave the UI alone.
# excel_writer.py appends it as a trailing column on the Excel main sheet
# only (after source_file, so no existing column index shifts - the summary
# sheet's _KEYS.index("source_file") lookup is unaffected), and
# read_existing_records() maps its header back explicitly so re-extending a
# project file doesn't drop the link.
WEIGHING_SOURCE_FILE_COLUMN = (WEIGHING_SOURCE_FILE_KEY, "קובץ תעודת שקילה")

# Collected from every certificate (they're ordinary FIELD_DEFS entries) but
# kept off the main table/sheet by default - see this module's docstring.
# Used by streamlit_app.py's "הצג פרטים נוספים" toggle, which appends these
# to the already-visible main columns, so no identifying columns are
# repeated here.
#
# vehicle_number/driver_name used to live here too, until 2026-09-03 when
# they were promoted to EXCEL_COLUMNS (main sheet) above - see that list's
# own comment.
INTERNAL_TRACKING_FIELDS = [
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
# about it" before the raw time/weight data.
INTERNAL_TRACKING_SHEET_COLUMNS = [
    ("source_file", "קובץ מקור"),
    ("certificate_or_reference", "מספר תעודה/אסמכתא"),
    ("site", "אתר/מקור"),
    ("confidence", "רמת ביטחון"),
] + INTERNAL_TRACKING_FIELDS
