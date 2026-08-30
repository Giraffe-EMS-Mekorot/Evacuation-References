"""Writes extracted certificate records to the consolidated output Excel file."""
from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import config
from .fields import (
    CONFIDENCE_LEVELS,
    EXCEL_COLUMNS,
    INTERNAL_TRACKING_SHEET_COLUMNS,
    REGIONS,
    WASTE_TYPES,
)

_MAIN_SHEET_NAME = "ריכוז תעודות"
_INTERNAL_SHEET_NAME = "מעקב פנימי"
_SUMMARY_SHEET_NAME = "סיכום"
# A generous fixed bound, not the actual row count, so every SUMIFS/COUNTIF
# formula on the summary sheet keeps covering rows added/edited by hand later
# without the summary sheet needing to be regenerated.
_MAX_DATA_ROW = 10000

_UNIT_LABELS = ['ק"ג', "טון", 'מ"ק', "ליטר", "יחידות"]

_FILL_BY_CONFIDENCE = {
    "נמוכה": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),  # red
    "בינונית": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),  # yellow
}
_BANDED_ROW_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
_TOTAL_ROW_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

_THIN_SIDE = Side(style="thin", color="000000")
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

_HEADER_FONT = Font(bold=True, size=13, color="FFFFFF")
_DATA_FONT = Font(size=11)
_HYPERLINK_FONT = Font(size=11, color="0563C1", underline="single")
_SECTION_TITLE_FONT = Font(bold=True, size=12)
_TOTAL_FONT = Font(bold=True, size=11)

_QTY_NUMBER_FORMAT = "#,##0.##"

# Columns whose values read better centered (numeric) vs. right-aligned (text/RTL default).
_CENTERED_KEYS = {"quantity"}

# Hebrew headers this pipeline used before the core-column reorder/rename
# (2026-08-30) - kept so read_existing_records() still recognizes a column
# from a file written by that earlier version instead of dropping its data;
# see that function's docstring. A value of None means the old column has no
# replacement (it was dropped, not renamed).
_LEGACY_HEADER_ALIASES = {
    "כמות": "quantity",
    "אתר/יחידה": "site",
    "סוג אסמכתא מצורפת": "supplier_or_carrier",
    "שנה": None,
    "חודש": None,
}

_KEYS = [key for key, _label in EXCEL_COLUMNS]
_HEADERS = [label for _key, label in EXCEL_COLUMNS]

# Rows needing the most urgent review should be the first thing a reviewer
# sees, not buried after everything else - note this is the opposite order
# from fields.CONFIDENCE_LEVELS (which lists גבוהה first; that list is about
# validating/describing the closed set of values, unrelated to display order).
CONFIDENCE_SORT_ORDER = {"נמוכה": 0, "בינונית": 1, "גבוהה": 2}


def sort_by_confidence(records: List[dict]) -> List[dict]:
    """Sorts records so נמוכה rows come first, then בינונית, then גבוהה - a
    stable sort, so rows within the same confidence level keep their
    relative order (batch/processing order) rather than being shuffled.
    Public so streamlit_app.py can keep its on-screen table in the same
    order as the file write_records() below is about to produce, instead of
    the two silently drifting apart.
    """
    return sorted(records, key=lambda r: CONFIDENCE_SORT_ORDER.get(r.get("confidence"), len(CONFIDENCE_SORT_ORDER)))


def parse_quantity(value: str) -> Optional[float]:
    """Parses a quantity string like '15,520.00' into a float, or None if it
    isn't a valid number (blank, unreadable, etc.) - needed so SUM/SUMIFS on
    the summary sheet actually work: openpyxl writes plain strings as text
    cells, which Excel's aggregate functions silently skip.

    Public (not `_`-prefixed) because streamlit_app.py needs the exact same
    parsing for its in-memory records before displaying/re-serializing them -
    see the note there about why raw extractor output can't be shown as-is.
    """
    if not value:
        return None
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def _esc(text: str) -> str:
    """Escapes a literal for embedding inside a double-quoted Excel formula string."""
    return text.replace('"', '""')


def write_records(records: List[dict], output_path: Path) -> None:
    """Writes one row per record plus a live-formula "סיכום" sheet, overwriting
    any existing file.

    Column order/headers on the main sheet come from fields.EXCEL_COLUMNS. Rows
    whose רמת_ביטחון is "נמוכה"/"בינונית" are highlighted red/yellow so they're
    easy to find for manual review; that always wins over the plain banded-row
    background. The source-file column is a hyperlink to the original file
    under input/. Rows are written in confidence order (נמוכה first) - see
    sort_by_confidence() - so the ones needing review are at the top, not
    scattered through the sheet in whatever order they were processed.
    """
    records = sort_by_confidence(records)

    wb = Workbook()
    ws = wb.active
    ws.title = _MAIN_SHEET_NAME
    ws.sheet_view.rightToLeft = True

    ws.append(_HEADERS)
    ws.row_dimensions[1].height = 26
    for cell in ws[1]:
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = _HEADER_FILL
        cell.border = _THIN_BORDER

    for row_idx, record in enumerate(records, start=2):
        row_values = []
        for key in _KEYS:
            raw = record.get(key, "")
            if key == "quantity":
                parsed = parse_quantity(raw)
                row_values.append(parsed if parsed is not None else raw)
            else:
                row_values.append(raw)
        ws.append(row_values)

        confidence_fill = _FILL_BY_CONFIDENCE.get(record.get("confidence", ""))
        band_fill = _BANDED_ROW_FILL if row_idx % 2 == 1 else None
        row_fill = confidence_fill or band_fill  # confidence coloring always wins

        for col_idx, key in enumerate(_KEYS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = _THIN_BORDER
            if row_fill:
                cell.fill = row_fill
            if key == "source_file":
                cell.font = _HYPERLINK_FONT
                cell.alignment = Alignment(horizontal="right")
                filename = record.get("source_file", "")
                if filename:
                    cell.hyperlink = (config.INPUT_DIR / filename).resolve().as_uri()
            else:
                cell.font = _DATA_FONT
                cell.alignment = Alignment(horizontal="center" if key in _CENTERED_KEYS else "right")
                if key == "quantity" and isinstance(cell.value, (int, float)):
                    cell.number_format = _QTY_NUMBER_FORMAT

    for col_idx, column_cells in enumerate(ws.columns, start=1):
        longest = max((len(str(c.value)) if c.value is not None else 0) for c in column_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 2, 10), 40)

    # Freezes the header row and the rightmost column (מרחב, logical column A - stays on
    # the visual right in this RTL sheet) so both stay visible while scrolling either way.
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    _write_internal_tracking_sheet(wb, records)
    _write_summary_sheet(wb)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _write_internal_tracking_sheet(wb: Workbook, records: List[dict]) -> None:
    """Writes the "מעקב פנימי" sheet: fields collected from the certificate
    but not shown on the main sheet by default (vehicle/driver/entry-exit-
    time/gross-tare-weight - see fields.INTERNAL_TRACKING_FIELDS), plus a
    few identifying columns (source file, certificate/reference number,
    site - see fields.INTERNAL_TRACKING_SHEET_COLUMNS) so a row here can be
    matched back to its main-sheet certificate without relying on row order.

    Written in the same order as the main sheet's rows (both come from the
    same already-confidence-sorted `records` list passed into
    write_records()), so row N here lines up with row N there for a given
    run - but unlike the סיכום sheet below, this one is plain extracted
    data, not a live formula, so a later manual edit on the main sheet won't
    be reflected here without rerunning the pipeline.
    """
    ws = wb.create_sheet(_INTERNAL_SHEET_NAME)
    ws.sheet_view.rightToLeft = True

    keys = [key for key, _label in INTERNAL_TRACKING_SHEET_COLUMNS]
    headers = [label for _key, label in INTERNAL_TRACKING_SHEET_COLUMNS]

    ws.append(headers)
    ws.row_dimensions[1].height = 22
    for cell in ws[1]:
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = _HEADER_FILL
        cell.border = _THIN_BORDER

    for row_idx, record in enumerate(records, start=2):
        ws.append([record.get(key, "") for key in keys])
        for col_idx in range(1, len(keys) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = _DATA_FONT
            cell.alignment = Alignment(horizontal="right")
            cell.border = _THIN_BORDER
            if row_idx % 2 == 1:
                cell.fill = _BANDED_ROW_FILL

    for col_idx, column_cells in enumerate(ws.columns, start=1):
        longest = max((len(str(c.value)) if c.value is not None else 0) for c in column_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 2, 10), 30)

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions


def read_existing_records(path: Path) -> List[dict]:
    """Reads a previously-written output file's main sheet back into a list of
    record dicts - the inverse of write_records()'s row-writing. Used by the
    Streamlit UI so a new processing run can extend an existing file instead
    of silently overwriting whatever earlier sessions already put in it (the
    CLI doesn't need this: main.py's "process everything in input/" model has
    no equivalent history to preserve).

    Returns [] if the file doesn't exist, isn't a valid workbook, or has no
    main sheet - a missing/unreadable file is treated as an empty starting
    point, never an error, since callers use this only to seed a fresh batch.

    Matches columns by their header-row TEXT against EXCEL_COLUMNS's current
    labels (falling back to _LEGACY_HEADER_ALIASES for a header this version
    renamed/dropped), not by fixed column position - so a file written by an
    older version of this pipeline, with a different column order/count
    (e.g. before the 2026-08-30 core-column reorder), is still read
    correctly for whatever columns still match by name, instead of silently
    misaligning every value one version's reorder to the left/right.
    """
    if not path.is_file():
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []
    if _MAIN_SHEET_NAME not in wb.sheetnames:
        return []

    ws = wb[_MAIN_SHEET_NAME]
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    label_to_key = dict(zip(_HEADERS, _KEYS))
    col_keys = [label_to_key.get(label, _LEGACY_HEADER_ALIASES.get(label)) for label in header_row]

    records = []
    for row in ws.iter_rows(min_row=2, max_col=len(col_keys), values_only=True):
        if all(value is None for value in row):
            continue
        record = {}
        for key, value in zip(col_keys, row):
            if key is None:
                continue  # a column this version no longer recognizes
            record[key] = value if key == "quantity" and value is not None else ("" if value is None else str(value))
        records.append(record)
    return records


def _range_ref(col_letter: str) -> str:
    return f"'{_MAIN_SHEET_NAME}'!${col_letter}$2:${col_letter}${_MAX_DATA_ROW}"


def _write_summary_sheet(wb: Workbook) -> None:
    """Builds the "סיכום" sheet entirely out of live formulas (SUM/SUMIF/SUMIFS/
    COUNTIF) against the main sheet, so manual edits or additions on the main
    sheet are reflected automatically - nothing here is a precomputed value.
    """
    ws = wb.create_sheet(_SUMMARY_SHEET_NAME)
    ws.sheet_view.rightToLeft = True

    region_col = get_column_letter(_KEYS.index("region") + 1)
    waste_col = get_column_letter(_KEYS.index("waste_type") + 1)
    qty_col = get_column_letter(_KEYS.index("quantity") + 1)
    unit_col = get_column_letter(_KEYS.index("unit") + 1)
    conf_col = get_column_letter(_KEYS.index("confidence") + 1)
    src_col = get_column_letter(_KEYS.index("source_file") + 1)

    region_range = _range_ref(region_col)
    waste_range = _range_ref(waste_col)
    qty_range = _range_ref(qty_col)
    unit_range = _range_ref(unit_col)
    conf_range = _range_ref(conf_col)
    src_range = _range_ref(src_col)

    def q(text: str) -> str:
        return f'"{_esc(text)}"'

    def write_cell(r, c, value, *, font=None, fill=None, align="right", number_format=None, border=True):
        cell = ws.cell(row=r, column=c, value=value)
        cell.font = font or _DATA_FONT
        if fill:
            cell.fill = fill
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if number_format:
            cell.number_format = number_format
        if border:
            cell.border = _THIN_BORDER
        return cell

    row = 1

    def section_header(text):
        nonlocal row
        write_cell(row, 1, text, font=_SECTION_TITLE_FONT, border=False)
        row += 1

    def table_header(labels):
        nonlocal row
        for c, label in enumerate(labels, start=1):
            write_cell(row, c, label, font=_HEADER_FONT, fill=_HEADER_FILL, align="center")
        row += 1

    # --- Section 1: total certificates --------------------------------------------------
    section_header('סה"כ תעודות שעובדו')
    write_cell(row, 1, 'סה"כ תעודות', font=_TOTAL_FONT)
    write_cell(row, 2, f"=COUNTA({src_range})", font=_TOTAL_FONT, align="center", fill=_TOTAL_ROW_FILL)
    row += 2

    # --- Section 2: confidence breakdown ---------------------------------------------------
    section_header("פילוח לפי רמת ביטחון")
    table_header(["רמת ביטחון", "מספר תעודות"])
    for level in CONFIDENCE_LEVELS:
        write_cell(row, 1, level)
        write_cell(row, 2, f"=COUNTIF({conf_range},{q(level)})", align="center")
        row += 1
    row += 1

    # --- Section 3: quantity by waste type x unit ------------------------------------------
    section_header("כמות שפונתה לפי סוג פסולת ויחידת מידה")
    unit_cols = list(range(2, 2 + len(_UNIT_LABELS)))  # B..F
    no_unit_col = unit_cols[-1] + 1  # G
    total_col = no_unit_col + 1  # H
    first_unit_letter = get_column_letter(unit_cols[0])
    last_unit_letter = get_column_letter(unit_cols[-1])
    no_unit_letter = get_column_letter(no_unit_col)
    total_letter = get_column_letter(total_col)

    table_header(["סוג פסולת"] + _UNIT_LABELS + ["ללא יחידה מזוהה", 'סה"כ'])
    first_type_row = row
    for waste_type in WASTE_TYPES:
        write_cell(row, 1, waste_type)
        for c, unit in zip(unit_cols, _UNIT_LABELS):
            formula = f"=SUMIFS({qty_range},{waste_range},{q(waste_type)},{unit_range},{q(unit)})"
            write_cell(row, c, formula, align="center", number_format=_QTY_NUMBER_FORMAT)
        # remainder for this waste type: total for the type minus what landed in a named unit
        write_cell(
            row, no_unit_col,
            f"=SUMIFS({qty_range},{waste_range},{q(waste_type)})"
            f"-SUM({first_unit_letter}{row}:{last_unit_letter}{row})",
            align="center", number_format=_QTY_NUMBER_FORMAT,
        )
        write_cell(
            row, total_col,
            f"=SUM({first_unit_letter}{row}:{no_unit_letter}{row})",
            align="center", number_format=_QTY_NUMBER_FORMAT, font=_TOTAL_FONT,
        )
        row += 1
    last_type_row = row - 1

    # catch-all for waste_type blank/unrecognized: whatever's left in each unit
    # column after subtracting the named waste-type rows above.
    unclassified_row = row
    write_cell(row, 1, "לא מסווג / לא מזוהה")
    for c, unit in zip(unit_cols, _UNIT_LABELS):
        col_letter = get_column_letter(c)
        write_cell(
            row, c,
            f"=SUMIF({unit_range},{q(unit)},{qty_range})"
            f"-SUM({col_letter}{first_type_row}:{col_letter}{last_type_row})",
            align="center", number_format=_QTY_NUMBER_FORMAT,
        )
    write_cell(
        row, no_unit_col,
        f"=SUM({qty_range})-SUM({total_letter}{first_type_row}:{total_letter}{last_type_row})"
        f"-SUM({first_unit_letter}{row}:{last_unit_letter}{row})",
        align="center", number_format=_QTY_NUMBER_FORMAT,
    )
    write_cell(
        row, total_col,
        f"=SUM({first_unit_letter}{row}:{no_unit_letter}{row})",
        align="center", number_format=_QTY_NUMBER_FORMAT, font=_TOTAL_FONT,
    )
    row += 1

    # grand-total row across all waste types (named + unclassified) - should
    # reconcile exactly with SUM(qty_range) in the bottom-right (סה"כ) cell.
    write_cell(row, 1, 'סה"כ', font=_TOTAL_FONT, fill=_TOTAL_ROW_FILL)
    for c in range(2, total_col + 1):
        col_letter = get_column_letter(c)
        write_cell(
            row, c,
            f"=SUM({col_letter}{first_type_row}:{col_letter}{unclassified_row})",
            align="center", number_format=_QTY_NUMBER_FORMAT, font=_TOTAL_FONT, fill=_TOTAL_ROW_FILL,
        )
    row += 2

    # --- Section 4: breakdown by region -----------------------------------------------------
    # Deliberately avoids any "region is blank"/"region <> known value" formula:
    # COUNTIF/SUMIF with "" or "<>" criteria against the padded 10000-row range
    # either count the thousands of never-written padding rows as "blank" (real
    # Excel behavior, not a bug, but not what we want here) or - for "<>" - were
    # empirically found to misbehave in testing. The catch-all row below is pure
    # arithmetic instead (total minus the 4 known-region sums), which only relies
    # on COUNTA/SUM and exact-value COUNTIF/SUMIF - all independently verified
    # correct. It merges "no region on the certificate" and "region value that
    # doesn't match one of the 4" into one row instead of splitting them.
    section_header("פירוט לפי מרחב")
    table_header(["מרחב", "מספר תעודות", 'סה"כ כמות (כל היחידות יחד)'])
    first_region_row = row
    for region in REGIONS:
        write_cell(row, 1, region)
        write_cell(row, 2, f"=COUNTIF({region_range},{q(region)})", align="center")
        write_cell(row, 3, f"=SUMIF({region_range},{q(region)},{qty_range})",
                   align="center", number_format=_QTY_NUMBER_FORMAT)
        row += 1
    last_known_region_row = row - 1

    write_cell(row, 1, "ללא מרחב מזוהה")
    write_cell(
        row, 2,
        f"=COUNTA({src_range})-SUM(B{first_region_row}:B{last_known_region_row})",
        align="center",
    )
    write_cell(
        row, 3,
        f"=SUM({qty_range})-SUM(C{first_region_row}:C{last_known_region_row})",
        align="center", number_format=_QTY_NUMBER_FORMAT,
    )
    last_region_row = row
    row += 1

    write_cell(row, 1, 'סה"כ', font=_TOTAL_FONT, fill=_TOTAL_ROW_FILL)
    write_cell(row, 2, f"=SUM(B{first_region_row}:B{last_region_row})",
               align="center", font=_TOTAL_FONT, fill=_TOTAL_ROW_FILL)
    write_cell(row, 3, f"=SUM(C{first_region_row}:C{last_region_row})",
               align="center", number_format=_QTY_NUMBER_FORMAT, font=_TOTAL_FONT, fill=_TOTAL_ROW_FILL)

    ws.column_dimensions["A"].width = 24
    for col_idx in range(2, total_col + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18
