"""Shared batch-processing pipeline: one certificate file in, one or more
records out (a multi-page PDF yields one record per page - see
extractor.extract_certificate_pages), used identically by main.py's CLI and
streamlit_app.py's web UI so the two front ends can never drift in how a
bad/corrupt file is handled.
"""
from pathlib import Path
from typing import Callable, List, NamedTuple, Optional

import anthropic

from . import config
from .excel_writer import parse_quantity, read_existing_records
from .extractor import extract_certificate_pages
from .fields import DOCUMENT_TYPE_OTHER, empty_record
from .normalize import NameNormalizer
from .quantity_check import build_quantity_history, flag_quantity_outlier

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


class ProcessResult(NamedTuple):
    """records: real certificates (including ones that failed to extract -
    those still belong here, flagged via confidence/notes, per this
    pipeline's long-standing policy of never silently dropping a genuine
    certificate). skipped: pages the model itself classified as not being a
    waste certificate at all (see fields.DOCUMENT_TYPE_OTHER) - kept
    completely separate so a batch summary of "certificates that need review"
    never gets diluted by irrelevant documents, and vice versa.
    """

    records: List[dict]
    skipped: List[dict]


def process_files(
    paths: List[Path],
    client: Optional[anthropic.Anthropic] = None,
    on_progress: Optional[ProgressCallback] = None,
    on_error: Optional[ErrorCallback] = None,
) -> ProcessResult:
    """Extracts one or more records per file, in order (one per page for a
    PDF, one for a plain image). A file that can't even be opened/split at
    all never aborts the batch - it's recorded via fields.empty_record() with
    the exception text in its notes, and processing continues with the rest.

    A page the model classifies as fields.DOCUMENT_TYPE_OTHER (not a waste
    certificate at all - e.g. an internal billing/credit note mixed into the
    same PDF) is routed into the returned ProcessResult.skipped list instead
    of .records, and skips normalization/quantity-outlier checks entirely -
    neither makes sense against a document that was never trying to report a
    waste quantity in the first place.

    Also runs each certificate record's free-text name fields (site/
    reference_type) through NameNormalizer, and its quantity through a
    historical-outlier sanity check - both bootstrapped from whatever's
    already in output/ריכוז_תעודות.xlsx (see app/normalize.py and
    app/quantity_check.py). This applies uniformly to both the CLI and the
    Streamlit UI, since both call this function.
    """
    client = client or anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    try:
        existing_records = read_existing_records(config.OUTPUT_DIR / config.OUTPUT_FILENAME)
        normalizer = NameNormalizer.build_from_records(existing_records)
        quantity_history = build_quantity_history(existing_records)
    except Exception:
        # Same "missing/unreadable history is an empty starting point, never
        # an error" policy read_existing_records() already applies to a
        # corrupt file - extended here to cover the bootstrap step itself, so
        # a batch can never fail before even a single file is attempted.
        normalizer = NameNormalizer()
        quantity_history = {}

    records = []
    skipped = []
    total = len(paths)
    for index, path in enumerate(paths, start=1):
        if on_progress:
            on_progress(index, total, path)
        try:
            page_records = extract_certificate_pages(path, client=client)
        except Exception as exc:  # the whole file couldn't be opened/split at all
            if on_error:
                on_error(path, exc)
            page_records = [empty_record(path.name, str(exc))]
        for record in page_records:
            if record.get("document_type") == DOCUMENT_TYPE_OTHER:
                skipped.append(record)
                continue
            try:
                normalizer.normalize(record)
                # quantity is still the raw extractor string here - parsed
                # just for this comparison, not written back, so every caller
                # keeps seeing the same raw-string-or-empty value they always
                # have (main.py's CLI never parses it in memory at all;
                # write_records() does that itself at Excel-write time).
                parsed_qty = parse_quantity(record.get("quantity"))
                if parsed_qty is not None:
                    flag_quantity_outlier(record, parsed_qty, quantity_history)
            except Exception as exc:
                # A bug in these two post-processing checks must never crash
                # an otherwise-successfully-extracted record - it still has
                # real data worth keeping, just without this round's
                # normalization/outlier pass. Surfaced as a visible note
                # (not swallowed silently) so a recurrence is diagnosable
                # from the spreadsheet itself, not just from server logs.
                note = f"שגיאה בבדיקת נרמול/כמות: {exc}"
                record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note
            records.append(record)
    return ProcessResult(records=records, skipped=skipped)
