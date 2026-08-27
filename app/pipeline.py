"""Shared batch-processing pipeline: one certificate file in, one record out,
used identically by main.py's CLI and streamlit_app.py's web UI so the two
front ends can never drift in how a bad/corrupt file is handled.
"""
from pathlib import Path
from typing import Callable, List, Optional

import anthropic

from . import config
from .extractor import extract_certificate
from .fields import FIELD_DEFS

# Called as on_progress(index, total, path) right before each file is sent to
# the model - lets a caller (CLI print, Streamlit status widget) show progress
# without this module knowing anything about how it's displayed.
ProgressCallback = Callable[[int, int, Path], None]
# Called as on_error(path, exception) when a file's extraction raised, right
# before it's recorded via empty_record() - lets a caller surface the failure
# (CLI print, Streamlit warning) without changing batch-continuation behavior.
ErrorCallback = Callable[[Path, Exception], None]


def empty_record(filename: str, error: str) -> dict:
    """Builds a placeholder record for a file that raised during extraction,
    flagged for manual review rather than silently dropped from the batch.
    """
    record = {name: "" for name, *_ in FIELD_DEFS}
    record["source_file"] = filename
    record["year"] = ""
    record["month"] = ""
    record["confidence"] = "נמוכה"
    record["notes"] = f"שגיאת עיבוד: {error}"
    return record


def process_files(
    paths: List[Path],
    client: Optional[anthropic.Anthropic] = None,
    on_progress: Optional[ProgressCallback] = None,
    on_error: Optional[ErrorCallback] = None,
) -> List[dict]:
    """Extracts one record per file, in order. A single bad/corrupt file never
    aborts the batch - it's recorded via empty_record() with the exception
    text in its notes, and processing continues with the rest.
    """
    client = client or anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    records = []
    total = len(paths)
    for index, path in enumerate(paths, start=1):
        if on_progress:
            on_progress(index, total, path)
        try:
            record = extract_certificate(path, client=client)
        except Exception as exc:  # one bad certificate shouldn't kill the whole batch
            if on_error:
                on_error(path, exc)
            record = empty_record(path.name, str(exc))
        records.append(record)
    return records
