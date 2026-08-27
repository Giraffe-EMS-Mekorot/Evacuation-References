"""Sends a single certificate file to Claude and returns its extracted fields."""
import base64
from pathlib import Path
from typing import Optional

import anthropic

from . import config
from .derive import derive_year_month
from .fields import CONFIDENCE_LEVELS, FIELD_DEFS, REGIONS, WASTE_TYPES

# extension -> (content block type, media type)
_MEDIA_TYPES = {
    ".pdf": ("document", "application/pdf"),
    ".jpg": ("image", "image/jpeg"),
    ".jpeg": ("image", "image/jpeg"),
    ".png": ("image", "image/png"),
}

_TOOL_NAME = "record_certificate_data"

# Core fields (per the original spec) whose absence should always get a
# record flagged for manual review, regardless of what the model itself
# reports for רמת_ביטחון. year/month are derived together from the same
# "date" value, so they're always both blank or both filled - one check
# covers both, labeled as the pair for a clearer note.
_CORE_FIELD_LABELS = [
    ("waste_type", "סוג הפסולת"),
    ("quantity", "כמות"),
    ("unit", "יחידת מידה"),
    ("site", "אתר/יחידה"),
    ("region", "מרחב"),
    ("year", "שנה/חודש"),
]

_SYSTEM_PROMPT = (
    "אתה עוזר שמחלץ מידע מובנה מתעודות פינוי פסולת (PDF או תמונה סרוקה/מצולמת, "
    "חלקן ממוחשבות וחלקן כתובות בכתב יד). קרא את התעודה המצורפת בעיון והפעל את הכלי "
    f"{_TOOL_NAME} עם הנתונים שחילצת. "
    "מלא שדה רק אם הערך שלו כתוב או מופיע בבירור בתעודה עצמה - לא בהסקה מהקשר, "
    "לא משם הקובץ, ולא מדמיון לתעודות אחרות. "
    "אם שדה כלשהו אינו ניתן לזיהוי ודאי, או שמילויו מצריך הסקה/ניחוש - השאר אותו "
    "כמחרוזת ריקה. שדה ריק תמיד עדיף על ניחוש שגוי. "
    "כלל אצבע: אם אתה עומד לכתוב בשדה הערות שהערך שמילאת הוא 'הנחה', 'הערכה' או "
    "'ניחוש' - זהו סימן שהיה עליך להשאיר את השדה עצמו ריק ולתאר את חוסר הוודאות "
    "רק בהערות, לא למלא אותו. "
    "בשדה רמת_ביטחון דווח את הערכתך לגבי איכות הקריאה של התעודה כולה. "
    "בשדה הערות כתוב הערה קצרה בלבד - עד 4-5 מילים, לא משפט מלא - שמציינת רק את "
    "הבעיה עצמה בתמציתיות, למשל: 'כתב יד לא קריא', 'שדה חסר בתעודה', 'מספר תעודה "
    "מטושטש'. אל תכתוב הסברים ארוכים על מה שעשית או איך הגעת לערך - רק את הבעיה."
)


def _build_tool_schema() -> dict:
    properties = {}
    required = []
    for name, description, json_type in FIELD_DEFS:
        properties[name] = {"type": json_type, "description": description}
        required.append(name)
    return {
        "name": _TOOL_NAME,
        "description": "רישום השדות המובנים שחולצו מתעודת פינוי פסולת אחת.",
        "input_schema": {
            "type": "object",
            "properties": properties,
            "required": required,
        },
    }


def _encode_file(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext not in _MEDIA_TYPES:
        raise ValueError(f"סיומת קובץ לא נתמכת: {ext}")
    block_type, media_type = _MEDIA_TYPES[ext]
    data = base64.standard_b64encode(path.read_bytes()).decode("utf-8")
    return {
        "type": block_type,
        "source": {"type": "base64", "media_type": media_type, "data": data},
    }


def _flag_out_of_list_values(record: dict) -> None:
    """Downgrades confidence and appends a note when a closed-list field is out of range."""
    issues = []
    if record["waste_type"] and record["waste_type"] not in WASTE_TYPES:
        issues.append(f"סוג_פסולת לא מזוהה: '{record['waste_type']}'")
    if record["region"] and record["region"] not in REGIONS:
        issues.append(f"מרחב לא מזוהה: '{record['region']}'")
    if record["confidence"] not in CONFIDENCE_LEVELS:
        record["confidence"] = record["confidence"] or "נמוכה"

    if issues:
        record["confidence"] = "נמוכה"
        note = "; ".join(issues)
        record["notes"] = f"{record['notes']} | {note}" if record["notes"] else note


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


def extract_certificate(path: Path, client: Optional[anthropic.Anthropic] = None) -> dict:
    """Sends one certificate file to Claude and returns a dict of extracted fields.

    The returned dict's keys match the field_name entries in fields.FIELD_DEFS,
    plus "source_file", "year" and "month" (the latter two derived in code from
    "date" - see derive.py). Values are always strings ("" when a field could
    not be read); nothing here guesses missing data.
    """
    client = client or anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    file_block = _encode_file(path)

    response = client.messages.create(
        model=config.MODEL_NAME,
        max_tokens=1024,
        system=_SYSTEM_PROMPT,
        tools=[_build_tool_schema()],
        tool_choice={"type": "tool", "name": _TOOL_NAME},
        messages=[
            {
                "role": "user",
                "content": [
                    file_block,
                    {"type": "text", "text": "חלץ את הנתונים מהתעודה המצורפת."},
                ],
            }
        ],
    )

    tool_use = next((block for block in response.content if block.type == "tool_use"), None)
    if tool_use is None:
        raise RuntimeError(f"המודל לא החזיר קריאת כלי (tool_use) עבור {path.name}")

    record = {name: str(tool_use.input.get(name, "") or "").strip() for name, *_ in FIELD_DEFS}
    record["source_file"] = path.name
    _flag_out_of_list_values(record)
    record["year"], record["month"] = derive_year_month(record["date"])
    _enforce_core_field_confidence(record)
    return record
