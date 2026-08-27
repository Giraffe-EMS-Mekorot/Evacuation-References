"""Shared batch-processing pipeline: one certificate file in, one or more
records out (a multi-page PDF yields one record per page - see
extractor.extract_certificate_pages), used identically by main.py's CLI and
streamlit_app.py's web UI so the two front ends can never drift in how a
bad/corrupt file is handled.
"""
from pathlib import Path
from typing import Callable, List, Optional

import anthropic

from . import config
from .extractor import extract_certificate_pages
from .fields import empty_record

# Called as on_progress(index, total, path) right before each file is sent to
# the model - lets a caller (CLI print, Streamlit status widget) show progress
# without this module knowing anything about how it's displayed. Fires once
# per file, not once per page, even for a multi-page PDF.
ProgressCallback = Callable[[int, int, Path], None]
# Called as on_error(path, exception) when a whole file couldn't be
# opened/split at all, right before it's recorded via fields.empty_record() -
# lets a caller surface the failure (CLI print, Streamlit warning) without
# changing batch-continuation behavior. A single bad *page* inside an
# otherwise-readable file doesn't reach this - see extract_certificate_pages.
ErrorCallback = Callable[[Path, Exception], None]


def process_files(
    paths: List[Path],
    client: Optional[anthropic.Anthropic] = None,
    on_progress: Optional[ProgressCallback] = None,
    on_error: Optional[ErrorCallback] = None,
) -> List[dict]:
    """Extracts one or more records per file, in order (one per page for a
    PDF, one for a plain image). A file that can't even be opened/split at
    all never aborts the batch - it's recorded via fields.empty_record() with
    the exception text in its notes, and processing continues with the rest.
    """
    client = client or anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    records = []
    total = len(paths)
    for index, path in enumerate(paths, start=1):
        if on_progress:
            on_progress(index, total, path)
        try:
            records.extend(extract_certificate_pages(path, client=client))
        except Exception as exc:  # the whole file couldn't be opened/split at all
            if on_error:
                on_error(path, exc)
            records.append(empty_record(path.name, str(exc)))
    return records
