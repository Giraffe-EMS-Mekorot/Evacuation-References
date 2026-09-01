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
from .derive import parse_date
from .excel_writer import parse_quantity, read_existing_records
from .extractor import extract_certificate_pages
from .fields import CERT_ROLE_BILL_OF_LADING_ZERO, DOCUMENT_TYPE_OTHER, empty_record
from .normalize import NameNormalizer
from .quantity_check import build_quantity_history, flag_quantity_outlier

# Date proximity (in days) for _cross_check_bill_of_lading_estimates - same
# rounding-margin reasoning as extractor.py's own tolerances, sized a bit
# more generously here since it's bridging two independently-issued
# documents (a carrier's שטר מטען vs. the receiving site's own weighing
# slip) rather than two adjacent pages of the same PDF.
_ESTIMATE_DATE_PROXIMITY_DAYS = 3


def _is_real_weighing_candidate(record: dict) -> bool:
    """True if `record` looks like an actual completed weighing - real,
    positive gross/tare, or at least a positive net quantity - rather than a
    pre-weighing שטר מטען estimate (see _cross_check_bill_of_lading_estimates).

    Deliberately NOT restricted to cert_role == CERT_ROLE_WEIGHING: that tag
    only fires for the specific two-adjacent-page (weighing/completion)
    pattern from the 2026-08-30 feature (see
    extractor._detect_and_cross_check_pairs) - an ordinary single-document
    weighing slip with no paired completion approval never gets that tag at
    all, but is exactly as valid a "real weighing happened" source here.
    """
    if record.get("cert_role") == CERT_ROLE_BILL_OF_LADING_ZERO:
        return False
    gross = parse_quantity(record.get("gross_weight"))
    tare = parse_quantity(record.get("tare_weight"))
    if gross is not None and tare is not None and gross > 0 and tare > 0:
        return True
    qty = parse_quantity(record.get("quantity"))
    return qty is not None and qty > 0


def _cross_check_bill_of_lading_estimates(records: List[dict]) -> None:
    """Looks, across the WHOLE batch just processed (every file in this
    process_files() call, not just one PDF's adjacent pages - contrast with
    extractor._detect_and_cross_check_pairs, which is strictly same-PDF/
    adjacent-page) for a real weighing record that matches a שטר מטען
    estimate record (see fields.CERT_ROLE_BILL_OF_LADING_ZERO and
    extractor._flag_bill_of_lading_estimate) on vehicle number + site + a
    nearby date.

    Deliberately does NOT require a matching certificate/reference number -
    per the spec this implements, the real-world case is exactly two
    documents from two different systems (e.g. a carrier's own שטר מטען vs.
    the receiving site's own weighing slip) that may each assign their own,
    unrelated id to the same physical shipment; requiring id equality here
    (like extractor._cross_check_document_pair does for the WEIGHING/
    COMPLETION pattern) would simply never match in that case.

    On a match, overwrites the ESTIMATE record's quantity/unit with the real
    weighing record's - per spec, "העדף את נתוני השקילה האמיתית על
    ההערכה" - and appends a note documenting the substitution on top of
    _flag_bill_of_lading_estimate's own note (never replacing it - both
    describe real, distinct facts about this row's history). The real
    weighing record itself is left untouched. No match leaves the estimate
    exactly as it already was - not a failure, just nothing to cross-check
    against yet (e.g. the real weighing certificate hasn't been uploaded in
    this batch at all).
    """
    estimates = [r for r in records if r.get("cert_role") == CERT_ROLE_BILL_OF_LADING_ZERO]
    if not estimates:
        return
    candidates = [r for r in records if _is_real_weighing_candidate(r)]
    if not candidates:
        return

    for estimate in estimates:
        est_vehicle = (estimate.get("vehicle_number") or "").strip().casefold()
        est_site = (estimate.get("site") or "").strip().casefold()
        est_date = parse_date(estimate.get("date"))
        if not est_vehicle or not est_site or not est_date:
            continue  # not enough identifying information to safely match
        for candidate in candidates:
            if candidate is estimate:
                continue
            cand_vehicle = (candidate.get("vehicle_number") or "").strip().casefold()
            cand_site = (candidate.get("site") or "").strip().casefold()
            cand_date = parse_date(candidate.get("date"))
            if cand_vehicle != est_vehicle or cand_site != est_site or not cand_date:
                continue
            if abs((est_date - cand_date).days) > _ESTIMATE_DATE_PROXIMITY_DAYS:
                continue

            estimate["quantity"] = candidate.get("quantity")
            estimate["unit"] = candidate.get("unit")
            estimate["confidence"] = candidate.get("confidence") or estimate["confidence"]
            note = "כמות עודכנה מתעודת שקילה אמיתית שהוצלבה (תאריך+רכב+אתר)"
            if note not in (estimate.get("notes") or ""):
                estimate["notes"] = f"{estimate['notes']} | {note}" if estimate.get("notes") else note
            break

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
    supplier_or_carrier) through NameNormalizer, and its quantity through a
    historical-outlier sanity check - both bootstrapped from whatever's
    already in output/ריכוז_תעודות.xlsx (see app/normalize.py and
    app/quantity_check.py). This applies uniformly to both the CLI and the
    Streamlit UI, since both call this function.

    Finally, once every file's records are collected, runs
    _cross_check_bill_of_lading_estimates() once over the whole batch - a
    שטר מטען page whose weight/volume was an explicit "0" (see
    fields.CERT_ROLE_BILL_OF_LADING_ZERO) gets its estimated quantity
    replaced with a real weighing record's, if one matching by vehicle+site+
    nearby date turns up anywhere else in this same batch.
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
    _cross_check_bill_of_lading_estimates(records)
    return ProcessResult(records=records, skipped=skipped)
