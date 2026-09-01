"""Few-shot example library for extraction accuracy (2026-09) - see
examples/README.md for how a human adds a new example, no code changes
needed. Loads ground-truth rows from examples/index.xlsx and the
certificate files sitting alongside it under examples/<category>/, and
turns up to two of them (the "hardest" ones, per the index's free-text note
- see _is_difficult) into a fixed few-shot message preamble that
extractor.py's _extract_page() prepends before the real page's own message.

**Built, but not currently active** (2026-09 decision): not pursued further
for now, on purpose - see extractor.py's use_examples parameter, which
defaults to False everywhere it's threaded through
(extract_certificate_pages(), pipeline.process_files()). This module itself
still works exactly as described below whenever a caller explicitly passes
use_examples=True; nothing here changed, only the default that decides
whether ordinary processing calls it at all. Left in place, ready, rather
than removed, since none of it costs anything while inactive - see "Fails
safe everywhere" below, which is exactly what makes an inactive-by-default
posture safe.

**Single-pass design, not type-matched to the real page** (confirmed with
the user, 2026-09): fields.py's document_type/cert_role are only known
AFTER the same model call this feature is meant to improve - there's no
cheap way to know a real page's exact structure BEFORE picking which
category's examples to show it, short of a separate classification call
that would roughly double the per-page image-processing cost (the same
real image sent to the API twice). Given that trade-off, every extraction
call always draws its few-shot examples from the single general "תעודת
פינוי" category (GENERAL_CATEGORY) - the category the overwhelming
majority of real certificates fall into anyway, since cert_role stays
blank on most documents (see fields.py's own note on that). The
per-category folder structure under examples/ still exists and is meant to
be populated exactly as specified, so nothing needs reorganizing later -
only the general pool is actively read for now. Precise type-matched
selection would need the two-pass classify-then-extract architecture
described above, if ever wanted badly enough to accept its cost.

Fails safe everywhere: a missing examples/ directory, a missing or empty
index.xlsx, or an index row whose file doesn't actually exist all degrade
to "no examples this run" (an empty list from build_fewshot_messages()),
never an exception - this feature augments extraction, it must never be
why a batch fails.
"""
from pathlib import Path
from typing import Dict, List

import pypdfium2 as pdfium
from openpyxl import load_workbook
from PIL import Image

from . import config
from .fields import (
    CERT_ROLE_COMPLETION,
    CERT_ROLES,
    DOCUMENT_TYPE_CERTIFICATE,
    FIELD_DEFS,
    TOOL_NAME,
)
from .image_utils import PDF_RENDER_SCALE, downscale_and_encode, encode_image_block

_INDEX_FILENAME = "index.xlsx"
_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}

# The category every call draws its few-shot examples from by default - see
# this module's docstring ("single-pass design").
GENERAL_CATEGORY = DOCUMENT_TYPE_CERTIFICATE  # "תעודת פינוי"

# Every category examples/ has (or can have) its own subfolder for - the
# same document_type/cert_role values fields.py already defines, so this
# stays consistent with what the extraction pipeline itself recognizes
# rather than inventing a parallel taxonomy. Exported so a one-time setup
# script (or a future admin UI) can create every folder without duplicating
# this list.
ALL_CATEGORIES = [GENERAL_CATEGORY] + CERT_ROLES

_INDEX_HEADER_TO_KEY = {
    "שם קובץ": "filename",
    "סוג מסמך": "category",
    "מרחב": "region",
    "אתר/מקור": "site",
    "אתר קולט": "supplier_or_carrier",
    "תאריך": "date",
    "מספר תעודה/אסמכתא": "certificate_or_reference",
    "סוג הפסולת": "waste_type",
    "כמות (נטו)": "quantity",
    "יחידת מידה": "unit",
    "הערה": "note",
}

# A note containing this word marks an example as a priority/"hard" one -
# see select_examples(). Plain substring match, not a separate boolean
# column, to keep the index quick for a human to fill by hand.
_DIFFICULT_MARKER = "קשה"
_MAX_EXAMPLES_PER_CALL = 2


def _sanitize_category_folder_name(category: str) -> str:
    """Turns a fields.py category string into a filesystem-safe folder name.
    A few category values (CERT_ROLE_WEIGHING, CERT_ROLE_COMPLETION,
    CERT_ROLE_BILL_OF_LADING_ZERO) contain a literal "/", which would
    otherwise silently create an unintended nested subdirectory instead of
    a single folder named after the category.
    """
    return category.replace("/", "-").replace("\\", "-")


def category_folder(category: str) -> Path:
    return config.EXAMPLES_DIR / _sanitize_category_folder_name(category)


def _cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _read_index() -> List[dict]:
    """Reads examples/index.xlsx (first sheet, header-name column matching -
    same approach as app.excel_input's manual-upload reader, for the same
    reasons: column order doesn't matter, an unrecognized column is
    silently ignored). Returns [] - never raises - if the file or directory
    doesn't exist yet, or isn't a valid workbook; this library is meant to
    work (with zero examples) before any real example has been added.
    """
    index_path = config.EXAMPLES_DIR / _INDEX_FILENAME
    if not index_path.is_file():
        return []
    try:
        wb = load_workbook(index_path, data_only=True)
    except Exception:
        return []
    ws = wb.worksheets[0]
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    col_keys = [_INDEX_HEADER_TO_KEY.get(_cell_text(label)) for label in header_row]
    if not any(col_keys):
        return []

    rows = []
    for row in ws.iter_rows(min_row=2, max_col=len(col_keys), values_only=True):
        if all(value is None for value in row):
            continue
        entry = {}
        for key, value in zip(col_keys, row):
            if key is None:
                continue
            entry[key] = value if key == "quantity" and isinstance(value, (int, float)) else _cell_text(value)
        if entry.get("filename") and entry.get("category"):
            rows.append(entry)
    return rows


def _is_difficult(entry: dict) -> bool:
    return _DIFFICULT_MARKER in (entry.get("note") or "")


def select_examples(category: str, limit: int = _MAX_EXAMPLES_PER_CALL) -> List[dict]:
    """Picks up to `limit` index rows for `category`, examples marked
    "difficult" first (see _is_difficult - "the hardest ones are worth the
    most" per the spec this implements), otherwise in index order. Skips a
    row whose file doesn't actually exist under its category folder
    (printed as a warning, not raised - one missing/renamed example file
    must never fail a real extraction batch over it).
    """
    candidates = [entry for entry in _read_index() if entry.get("category") == category]
    candidates.sort(key=lambda entry: 0 if _is_difficult(entry) else 1)

    selected = []
    for entry in candidates:
        if len(selected) >= limit:
            break
        file_path = category_folder(category) / entry["filename"]
        if not file_path.is_file():
            print(f"אזהרה: קובץ דוגמה חסר, מדולג: {file_path}")
            continue
        selected.append({**entry, "_path": file_path})
    return selected


def _expected_tool_input(entry: dict, category: str) -> dict:
    """Builds the fields.FIELD_DEFS-shaped answer dict this example's
    fabricated assistant turn "returns" - the known-correct ground truth
    from the index for every field the index actually captures, and ""
    (never guessed) for every other FIELD_DEFS key the index doesn't
    collect ground truth for (vehicle/driver/entry-exit-time/gross-tare/
    container - see examples/README.md's column list). document_type is
    always DOCUMENT_TYPE_CERTIFICATE (only real certificates belong in this
    library - see this module's docstring), confidence is always "גבוהה"
    (this is presented to the model as a clean, confident read, on purpose
    - the whole point is showing it CAN be read confidently despite looking
    hard), and notes is always blank for the same reason.

    The index's one combined "מספר תעודה/אסמכתא" value goes into
    certificate_number for most categories, but reference_number for
    CERT_ROLE_COMPLETION specifically - matching that field's real meaning
    (see fields.py's own reference_number description: an "אישור ביצוע
    עבודה/הטמנה" document's own id field points BACK at another
    document's certificate_number, it isn't its own certificate_number).
    """
    data = {name: "" for name, *_ in FIELD_DEFS}
    data["document_type"] = DOCUMENT_TYPE_CERTIFICATE
    data["date"] = entry.get("date", "")
    data["waste_type"] = entry.get("waste_type", "")
    data["quantity"] = _cell_text(entry.get("quantity"))
    data["unit"] = entry.get("unit", "")
    data["site"] = entry.get("site", "")
    data["region"] = entry.get("region", "")
    data["supplier_or_carrier"] = entry.get("supplier_or_carrier", "")

    reference = entry.get("certificate_or_reference", "")
    if category == CERT_ROLE_COMPLETION:
        data["reference_number"] = reference
    else:
        data["certificate_number"] = reference
    if category in CERT_ROLES:
        data["cert_role"] = category

    data["confidence"] = "גבוהה"
    return data


def _load_example_image_block(path: Path) -> dict:
    """Loads and encodes one example file exactly the way a real certificate
    page is encoded (see image_utils.downscale_and_encode) - same
    resolution/compression treatment, so the model isn't shown a
    suspiciously different image quality for the "correct answer" example
    versus the real page it's about to read. A PDF example uses its first
    page only (a "hard example" is conceptually one certificate/page - see
    examples/README.md); an image file is used as-is.
    """
    ext = path.suffix.lower()
    if ext == ".pdf":
        pdf = pdfium.PdfDocument(str(path))
        try:
            image = pdf[0].render(scale=PDF_RENDER_SCALE).to_pil()
        finally:
            pdf.close()
    elif ext in _IMAGE_EXTENSIONS:
        image = Image.open(path)
    else:
        raise ValueError(f"סיומת קובץ דוגמה לא נתמכת: {ext}")
    jpeg_bytes = downscale_and_encode(image)
    return encode_image_block(jpeg_bytes)


# Cached per category for the life of this process - example files don't
# change mid-run, and re-reading/re-encoding the same 1-2 images for every
# page of a multi-page PDF (or every file in a whole batch) would be pure
# waste. Module-level (not per-call) so it also survives across Streamlit
# reruns within the same server process.
_message_cache: Dict[str, List[dict]] = {}


def build_fewshot_messages(category: str = GENERAL_CATEGORY) -> List[dict]:
    """Returns the message list to prepend before the real page's own user
    message in extractor._extract_page() - one (user-with-image,
    assistant-with-tool_use, user-with-tool_result) triple per selected
    example (see select_examples()), matching the Anthropic Messages API's
    required tool_use/tool_result pairing (an assistant tool_use must be
    immediately followed by a user message carrying a matching tool_result,
    or the API rejects the request).

    Returns [] - the same request extractor.py always sent before this
    feature existed - whenever there's nothing usable to show: no
    examples/ directory yet, an empty index, or every candidate row's file
    missing. This makes the feature completely inert (zero behavior change)
    until real examples actually exist, by construction, not by a separate
    on/off flag.
    """
    if category in _message_cache:
        return _message_cache[category]

    messages: List[dict] = []
    for i, entry in enumerate(select_examples(category)):
        try:
            image_block = _load_example_image_block(entry["_path"])
        except Exception as exc:
            print(f"אזהרה: כשל בטעינת קובץ דוגמה, מדולג: {entry['_path']} ({exc})")
            continue
        tool_use_id = f"toolu_example_{i}"
        messages.append(
            {
                "role": "user",
                "content": [image_block, {"type": "text", "text": "חלץ את הנתונים מהתעודה המצורפת."}],
            }
        )
        messages.append(
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "tool_use",
                        "id": tool_use_id,
                        "name": TOOL_NAME,
                        "input": _expected_tool_input(entry, category),
                    }
                ],
            }
        )
        messages.append(
            {
                "role": "user",
                "content": [{"type": "tool_result", "tool_use_id": tool_use_id, "content": "נרשם בהצלחה."}],
            }
        )

    _message_cache[category] = messages
    return messages
