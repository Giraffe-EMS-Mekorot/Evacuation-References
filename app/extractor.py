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
import base64
import io
from pathlib import Path
from typing import List, Optional

import anthropic
import pypdfium2 as pdfium
from PIL import Image

from . import config
from .derive import derive_year_month
from .fields import CONFIDENCE_LEVELS, FIELD_DEFS, REGIONS, WASTE_TYPES, empty_record

_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}

_TOOL_NAME = "record_certificate_data"

# Longest side, in pixels, after downscaling - comfortably below anything
# that would risk a 413, while staying legible for both print and handwriting.
_MAX_DIMENSION = 2000
_JPEG_QUALITY = 85
# Render PDF pages at ~200 DPI before downscaling - high enough that small
# handwriting survives the resize down to _MAX_DIMENSION, not so high that a
# single page wastes memory rendering detail that gets thrown away anyway.
_PDF_RENDER_SCALE = 200 / 72
# Safety net for step 4 of the fix: matches Anthropic's documented per-image
# limit. In practice a page downscaled to _MAX_DIMENSION/_JPEG_QUALITY should
# never get close to this - this guards the case where it somehow still does
# (e.g. a pathologically noisy scan) rather than sending it and hoping.
_MAX_IMAGE_BYTES = 5 * 1024 * 1024

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


def _downscale_and_encode(image: Image.Image) -> bytes:
    """Resizes `image` so its longest side is at most _MAX_DIMENSION pixels
    (never upscales a smaller image) and re-encodes it as a JPEG at
    _JPEG_QUALITY. This is what actually keeps each request small enough for
    the API - see the module docstring for why every page goes through this,
    not just ones from an oversized multi-page PDF.
    """
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    width, height = image.size
    longest = max(width, height)
    if longest > _MAX_DIMENSION:
        ratio = _MAX_DIMENSION / longest
        new_size = (max(1, round(width * ratio)), max(1, round(height * ratio)))
        image = image.resize(new_size, Image.LANCZOS)
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=_JPEG_QUALITY)
    return buffer.getvalue()


def _encode_image_block(jpeg_bytes: bytes) -> dict:
    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": "image/jpeg",
            "data": base64.standard_b64encode(jpeg_bytes).decode("utf-8"),
        },
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


def _append_page_note(record: dict, page_index: int, page_total: int) -> None:
    note = f"עמוד {page_index} מתוך {page_total}"
    record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note


def _extract_page(image: Image.Image, client: anthropic.Anthropic) -> dict:
    """Sends one already-loaded page image to Claude and returns its
    extracted fields. The single-page core that both a plain image file and
    each page of a split PDF funnel through. Doesn't set source_file - the
    caller (extract_certificate_pages) does, since it knows the filename.
    """
    jpeg_bytes = _downscale_and_encode(image)
    if len(jpeg_bytes) > _MAX_IMAGE_BYTES:
        raise ValueError(
            f"התמונה גדולה מדי לשליחה גם אחרי דחיסה ({len(jpeg_bytes) / 1_000_000:.1f}MB)"
        )
    file_block = _encode_image_block(jpeg_bytes)

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
        raise RuntimeError("המודל לא החזיר קריאת כלי (tool_use)")

    record = {name: str(tool_use.input.get(name, "") or "").strip() for name, *_ in FIELD_DEFS}
    _flag_out_of_list_values(record)
    record["year"], record["month"] = derive_year_month(record["date"])
    _enforce_core_field_confidence(record)
    return record


def extract_certificate_pages(path: Path, client: Optional[anthropic.Anthropic] = None) -> List[dict]:
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
                    image = page.render(scale=_PDF_RENDER_SCALE).to_pil()
                    record = _extract_page(image, client=client)
                except Exception as exc:  # one bad page shouldn't drop the rest of the file
                    record = empty_record(error=str(exc))
                finally:
                    page.close()
                record["source_file"] = path.name
                if total > 1:
                    _append_page_note(record, index + 1, total)
                records.append(record)
            return records
        finally:
            pdf.close()

    if ext in _IMAGE_EXTENSIONS:
        image = Image.open(path)
        try:
            record = _extract_page(image, client=client)
        except Exception as exc:
            record = empty_record(error=str(exc))
        record["source_file"] = path.name
        return [record]

    raise ValueError(f"סיומת קובץ לא נתמכת: {ext}")
