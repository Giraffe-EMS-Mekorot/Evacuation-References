"""Shared batch-processing pipeline: one certificate file in, one or more
records out (a multi-page PDF yields one record per page - see
extractor.extract_certificate_pages), used identically by main.py's CLI and
streamlit_app.py's web UI so the two front ends can never drift in how a
bad/corrupt file is handled.
"""
from pathlib import Path
from typing import Callable, List, NamedTuple, Optional, Sequence

import anthropic

from . import config
from .derive import parse_date
from .excel_writer import parse_quantity, read_existing_records
# _normalize_reference is deliberately reused rather than reimplemented here:
# the weighing-certificate cross-check below must compare ids by exactly the
# same rules as extractor._cross_check_document_pair already does (numeric
# ids with leading zeros stripped, anything else case-folded), and two
# independent copies of that rule would drift.
from .extractor import _normalize_reference, extract_certificate_pages
from .fields import (
    CERT_ROLE_BILL_OF_LADING_ZERO,
    CERT_ROLE_WEIGHING_CERTIFICATE,
    DOCUMENT_TYPE_OTHER,
    WEIGHING_MATCHED_KEY,
    WEIGHING_PAGE_NUMBER_KEY,
    WEIGHING_SOURCE_FILE_KEY,
    empty_record,
)
from .normalize import NameNormalizer
from .quantity_check import build_quantity_history, flag_quantity_outlier

# Date proximity (in days) for _cross_check_bill_of_lading_estimates - same
# rounding-margin reasoning as extractor.py's own tolerances, sized a bit
# more generously here since it's bridging two independently-issued
# documents (a carrier's שטר מטען vs. the receiving site's own weighing
# slip) rather than two adjacent pages of the same PDF.
_ESTIMATE_DATE_PROXIMITY_DAYS = 3

# How close two reported tare weights for the SAME vehicle must be to count as
# "the same tare" in _flag_swapped_tare_net. Relative, with an absolute floor,
# because a vehicle's own empty weight does drift a little between weighings
# (fuel, mud, a tool left in the cab) but not by tens of percent.
_VEHICLE_TARE_REL_TOLERANCE = 0.05
_VEHICLE_TARE_ABS_FLOOR = 50.0


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

def _tares_match(a: float, b: float) -> bool:
    """True if two reported tare weights plausibly describe the same vehicle -
    see _VEHICLE_TARE_REL_TOLERANCE."""
    allowed = max(abs(a), abs(b)) * _VEHICLE_TARE_REL_TOLERANCE
    return abs(a - b) <= max(allowed, _VEHICLE_TARE_ABS_FLOOR)


def _flag_swapped_tare_net(records: List[dict]) -> None:
    """Catches the one gross/tare/net error that arithmetic provably cannot:
    a טרה<->נטו swap.

    `gross - tare = net` and `gross - net = tare` are the same equation, so a
    record that reports those two fields the other way round passes
    `extractor._flag_gross_tare_mismatch()` perfectly (see that function's
    docstring for the real 2026-09-15 example that motivated this). The three
    numbers on one document simply do not contain the information needed to
    tell the two labelings apart - so this check brings in evidence from
    outside them: **the same vehicle's tare as reported on another document
    in the batch.**

    A vehicle's tare is a property of the vehicle, so it should be roughly
    constant across its certificates, while its net load varies per trip.
    For each record of a vehicle seen more than once, if its reported tare
    disagrees with the other documents' tares for that vehicle *and* its
    reported quantity (net) matches them instead, the two fields were almost
    certainly swapped on this record.

    **Flags, never corrects** - consistent with the rest of this pipeline.
    Rewriting quantity here would mean choosing a labeling on heuristic
    grounds, and a wrong auto-correction is worse than a flagged row: the
    note names both candidate values so a reviewer can settle it against the
    document in one look. Escalates to "נמוכה" because a swapped quantity is
    a wrong number in the main data column, not a cosmetic issue.

    Needs at least two documents for the same vehicle to say anything at all;
    a single weighing certificate in a batch is left alone (there is nothing
    to corroborate it against). Note the corroboration is batch-only: tare
    weights live on the "מעקב פנימי" sheet, which `read_existing_records()`
    does not read, so a previously-processed file's tares are not available
    here.
    """
    by_vehicle: dict = {}
    for record in records:
        vehicle = (record.get("vehicle_number") or "").strip().casefold()
        if not vehicle:
            continue
        tare = parse_quantity(record.get("tare_weight"))
        net = parse_quantity(record.get("quantity"))
        if tare is None or net is None or tare <= 0 or net <= 0:
            continue
        by_vehicle.setdefault(vehicle, []).append((record, tare, net))

    for vehicle, entries in by_vehicle.items():
        if len(entries) < 2:
            continue
        for index, (record, tare, net) in enumerate(entries):
            other_tares = [t for k, (_r, t, _n) in enumerate(entries) if k != index]
            if not other_tares:
                continue
            # Consistent with at least one other document for this vehicle -
            # nothing to suspect, whichever way round the others are.
            if any(_tares_match(tare, other) for other in other_tares):
                continue
            # Its tare matches nothing, but its NET matches the tare the rest
            # of this vehicle's documents agree on - the classic swap.
            if not any(_tares_match(net, other) for other in other_tares):
                continue
            note = (
                f"ייתכן שטרה ונטו התחלפו: הכמות המדווחת ({net:g}) תואמת לטרה של "
                f"אותו רכב במסמך אחר, בעוד הטרה המדווחת ({tare:g}) אינה - "
                f'בדוק מול התעודה איזה מהם הנטו'
            )
            record["confidence"] = "נמוכה"
            _append_note(record, note)


_NO_WEIGHING_MATCH_NOTE = "לא אותרה תעודת שקילה תואמת באצווה זו"


def _append_note(record: dict, note: str) -> None:
    """Appends `note` to record["notes"] unless it's already there - the same
    idempotent " | "-joined pattern every flagging function in extractor.py
    uses, lifted to a helper here because the weighing-certificate
    cross-check below writes notes onto both sides of a match from three
    different branches.
    """
    existing = record.get("notes") or ""
    if note in existing:
        return
    record["notes"] = f"{existing} | {note}" if existing else note


def _cross_check_weighing_certificates(
    records: List[dict], weighing_certificates: List[dict]
) -> None:
    """Matches each separate תעודת שקילה page in this batch (a
    CERT_ROLE_WEIGHING_CERTIFICATE record - see fields.py and
    extractor._strip_weighing_certificate, which has already reduced it to
    its printed header number alone) against the תעודת משלוח record whose
    מספר_אסמכתא (reference_number) carries that same number.

    **Matching is by number identity, not by position.** Both sides are
    indexed into {normalized_number: [records]} maps up front, so a batch
    holding many delivery certificates and many weighing certificates
    resolves each pair by key lookup - upload order, page order, and which
    file a page came from are all irrelevant, and a weighing certificate may
    arrive before or after its delivery certificate. Numbers are compared
    through extractor._normalize_reference (numeric ids compare with leading
    zeros stripped; anything else case-folded), the same normalization the
    adjacent-page WEIGHING/COMPLETION cross-check already uses.

    A link is only made when the number is unique on BOTH sides
    (exactly one delivery certificate and exactly one weighing certificate
    carry it). Four outcomes:

      1. Unique 1:1 match - the delivery record's "מספר אסמכתא" column
         (certificate_or_reference) is rewritten to the number as extracted
         from the weighing certificate itself, and the weighing
         certificate's file/page are recorded on the delivery record
         (WEIGHING_SOURCE_FILE_KEY/WEIGHING_PAGE_NUMBER_KEY) so
         excel_writer.py can render a second, separately-clickable
         hyperlink beside the row's own "קובץ מקור" link.
      2. Delivery certificate has an אסמכתא but no weighing certificate in
         this batch carries that number - its "מספר אסמכתא" column is
         cleared and _NO_WEIGHING_MATCH_NOTE is added, per spec.
      3. The number is duplicated on either side (two delivery certificates
         claiming the same אסמכתא, or two weighing certificates printed with
         the same number) - NO match is chosen. Both sides get an explicit
         note naming the counts, and confidence is escalated so the row is
         actually visible in the red/yellow review scan rather than relying
         on someone reading the notes column. Deliberately never "pick the
         first one": an arbitrary link here silently attributes a weighing
         certificate to the wrong shipment.
      4. Weighing certificate matched nothing - left unmatched (no
         WEIGHING_MATCHED_KEY), which is excel_writer.py's cue to list it on
         the "מעקב פנימי" sheet so it doesn't vanish silently. It never gets
         a row of its own.

    Runs once per process_files() batch, across every file in it - contrast
    with extractor._detect_and_cross_check_pairs, which is strictly
    same-PDF/adjacent-page. The two mechanisms are independent and compose:
    this one only ever writes certificate_or_reference and the two link
    keys, never certificate_number/reference_number, so a record can be half
    of a ברוטו/טרה/נטו adjacent pair AND cross-matched to a separate תעודת
    שקילה at the same time without either check disturbing the other's
    inputs.
    """
    if not weighing_certificates:
        return

    certs_by_number: dict = {}
    for cert in weighing_certificates:
        number = _normalize_reference(cert.get("certificate_number") or "")
        if not number:
            continue  # already flagged at extraction time - nothing to match on
        certs_by_number.setdefault(number, []).append(cert)

    deliveries_by_number: dict = {}
    for record in records:
        reference = _normalize_reference(record.get("reference_number") or "")
        if reference:
            deliveries_by_number.setdefault(reference, []).append(record)

    for number, deliveries in deliveries_by_number.items():
        matches = certs_by_number.get(number, [])

        if not matches:
            for delivery in deliveries:
                delivery["certificate_or_reference"] = ""
                _append_note(delivery, _NO_WEIGHING_MATCH_NOTE)
            continue

        if len(matches) > 1 or len(deliveries) > 1:
            raw = (deliveries[0].get("reference_number") or "").strip()
            note = (
                f"כפילות בהצלבת תעודת שקילה: מספר {raw} מופיע ב-{len(deliveries)} "
                f"תעודות משלוח ו-{len(matches)} תעודות שקילה - לא בוצעה התאמה אוטומטית"
            )
            for record in list(deliveries) + list(matches):
                record["certificate_or_reference"] = ""
                if record.get("confidence") not in ("נמוכה", "בינונית"):
                    record["confidence"] = "בינונית"
                _append_note(record, note)
            continue

        delivery, cert = deliveries[0], matches[0]
        # Per spec the column shows the number as it was read off the
        # weighing certificate itself, not the delivery certificate's own
        # transcription of it - they normalize equal but can differ in
        # formatting (leading zeros), and the weighing certificate is where
        # that number is printed.
        delivery["certificate_or_reference"] = (cert.get("certificate_number") or "").strip()
        delivery[WEIGHING_SOURCE_FILE_KEY] = cert.get("source_file", "")
        page_number = cert.get("page_number")
        if page_number:
            delivery[WEIGHING_PAGE_NUMBER_KEY] = page_number
        cert[WEIGHING_MATCHED_KEY] = True


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

    weighing_certificates (added 2026-09-15): separate תעודת שקילה pages
    (fields.CERT_ROLE_WEIGHING_CERTIFICATE) that accompany a computerized
    תעודת משלוח. A third list rather than more `records`, because per spec
    such a page contributes only its printed number to *another* record's
    row and must never become a row of its own - keeping them out of
    `records` is what gives that guarantee on the main Excel sheet and in
    streamlit_app.py's results table for free, with no filtering needed in
    either. Not `skipped` either: these aren't irrelevant documents, they're
    a real part of the data, and the unmatched ones are reported on the
    "מעקב פנימי" sheet (see excel_writer.write_records()).
    """

    records: List[dict]
    skipped: List[dict]
    # Immutable default on purpose: a shared mutable [] default would be
    # appended to by any caller that built a ProcessResult without this
    # field (a test, say) and then mutated it. Callers only ever iterate it.
    weighing_certificates: Sequence[dict] = ()


def process_files(
    paths: List[Path],
    client: Optional[anthropic.Anthropic] = None,
    on_progress: Optional[ProgressCallback] = None,
    on_error: Optional[ErrorCallback] = None,
    use_examples: bool = False,
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

    _flag_swapped_tare_net() also runs batch-wide, flagging a record whose
    טרה/נטו look swapped judged against the same vehicle's tare on another
    document - the one gross/tare/net error the per-record arithmetic check
    in extractor.py provably cannot detect.

    Then _cross_check_weighing_certificates() runs, also once over the whole
    batch: every separate תעודת שקילה page in it (routed into
    ProcessResult.weighing_certificates, never into .records) is matched by
    printed-number identity against the תעודת משלוח whose מספר_אסמכתא
    carries that number. See that function's docstring for the
    many-to-many indexing and the duplicate-number handling.

    use_examples (default False - inactive/not pursued further per an
    explicit 2026-09 user decision, kept ready but off so it adds zero
    cost/latency unless someone opts in) is passed straight through to
    extractor.extract_certificate_pages() for every file - see that
    function's own docstring and app/examples_library.py.
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
    weighing_certificates = []
    total = len(paths)
    for index, path in enumerate(paths, start=1):
        if on_progress:
            on_progress(index, total, path)
        try:
            page_records = extract_certificate_pages(path, client=client, use_examples=use_examples)
        except Exception as exc:  # the whole file couldn't be opened/split at all
            if on_error:
                on_error(path, exc)
            page_records = [empty_record(path.name, str(exc))]
        for record in page_records:
            if record.get("document_type") == DOCUMENT_TYPE_OTHER:
                skipped.append(record)
                continue
            if record.get("cert_role") == CERT_ROLE_WEIGHING_CERTIFICATE:
                # Reduced to its printed number alone at extraction time
                # (extractor._strip_weighing_certificate). Skips normalization
                # and the quantity-outlier check for the same reason
                # DOCUMENT_TYPE_OTHER does - it has no site/supplier/quantity
                # to normalize or compare, by design - and is cross-matched
                # against the batch's delivery certificates below instead.
                weighing_certificates.append(record)
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
    _flag_swapped_tare_net(records)
    _cross_check_weighing_certificates(records, weighing_certificates)
    return ProcessResult(
        records=records, skipped=skipped, weighing_certificates=weighing_certificates
    )
