"""Web UI for the tteudot-agent certificate pipeline - runnable locally or
deployed to Streamlit Community Cloud (see README.md).

An alternative front end alongside main.py's CLI, not a replacement for it:
upload files here (drag-and-drop) instead of dropping them into input/, and
get an Excel file built through the identical app.pipeline.process_files()
the CLI uses, so the two front ends can never disagree on how a certificate
is read or how a bad file is handled.

Each *project* here is its own separate Excel file under output/, explicitly
named by the user - see the "Session state model" comment below for exactly
how a project's file grows without ever silently overwriting another one's
data or losing anything already saved. main.py's CLI is unaffected by this -
it still always writes the one fixed output/ריכוז_תעודות.xlsx.

Gated behind a single shared password (see _check_password() below) - never
hardcoded, always read from st.secrets["APP_PASSWORD"].

Run with:
    streamlit run streamlit_app.py
"""
import re
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from app import config
from app.excel_writer import parse_quantity, read_existing_records, sort_by_confidence, write_records
from app.fields import CONFIDENCE_LEVELS, EXCEL_COLUMNS, HEBREW_MONTHS, REGIONS, UNCLASSIFIED_WASTE_TYPE, WASTE_TYPES
from app.pipeline import process_files

_FLAGGED_LEVELS = {"נמוכה", "בינונית"}
# Matches the row-fill colors excel_writer.py uses for these same confidence
# levels in the downloaded file, so the on-screen table and the spreadsheet
# always agree visually.
_ROW_COLORS = {"נמוכה": "#FFC7CE", "בינונית": "#FFEB9C"}
_LOGO_PATH = Path(__file__).resolve().parent / "logo.jpg"
# The upload area's accent border/background, per the design request - not
# in .streamlit/config.toml because there's no theme token for "this one
# specific widget's background"; see the CSS block below for why this
# targets the card's own key rather than a fixed/generic selector.
_UPLOAD_AREA_BORDER = "#A47E5B"
_UPLOAD_AREA_BACKGROUND = "#F3EADC"
# A lighter green than primaryColor, for the header's gradient depth - not a
# theme token (config.toml has no "secondary green" slot), just a plain
# constant used directly in the CSS block below.
_ACCENT_GREEN_LIGHT = "#3D6B4F"
# Matches config.toml's orangeColor - used both there (so Streamlit's own
# orange-tinted elements pick it up) and here directly, for the one bit of
# custom CSS that needs the literal value: the high-confidence accent border
# in the results table (see _highlight_source_cell).
_ACCENT_GOLD = "#C9982F"
# Symbols shown in the read-only "status" column next to רמת ביטחון, in
# addition to (not instead of) that column's own text value and the existing
# background-fill coloring - a second, non-color signal for the same
# information, per the design request.
_STATUS_SYMBOLS = {"גבוהה": "✓", "בינונית": "⚠", "נמוכה": "⚠"}


def _default_project_name() -> str:
    return f"פרויקט {date.today():%d-%m-%Y}"


def _sanitize_filename_part(text: str) -> str:
    """Strips characters that are illegal (or awkward) in a filename and
    collapses whitespace to underscores, so a free-text project name can't
    produce a broken path. Never returns an empty string.
    """
    cleaned = re.sub(r'[\\/:*?"<>|]', "", text).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "פרויקט"


def _build_output_path(project_name: str) -> Path:
    """Every project gets its own file: ריכוז_תעודות_<project>_<date>.xlsx.
    The date is always today's (when the project is activated), independent
    of whatever the project name itself contains - see _activate_project().
    """
    safe_name = _sanitize_filename_part(project_name)
    return config.OUTPUT_DIR / f"ריכוז_תעודות_{safe_name}_{date.today():%d-%m-%Y}.xlsx"


st.set_page_config(
    page_title="מערכת ריכוז תעודות פינוי",
    page_icon=":material/recycling:",
    layout="wide",
)

# Streamlit has no built-in RTL layout mode; the first rule below is the
# minimal CSS needed to mirror the app right-to-left for Hebrew - a
# functional requirement, not decorative styling. Everything else in this
# block IS deliberately decorative (the "card" look, the header gradient),
# targeted at specific containers' own st.container(key=...) values rather
# than generic selectors that would repaint unrelated widgets:
#
#   header_card   - the logo+title+caption banner (both here and in the
#                   password screen - the two never render in the same
#                   script run, so reusing one key is safe). Dark gradient
#                   background needs the title/caption text forced light -
#                   config.toml's :green[]/:gray[] markdown colors are for
#                   the plain cream background elsewhere, not readable here.
#   upload_card / results_card / skipped_card - plain "card" containers:
#                   distinct background, soft border, subtle shadow. Cheap
#                   visual separation Streamlit's own border=True can't give
#                   (it only draws a border in the page's own background).
#
# upload_card additionally gets the gold dashed border from the earlier
# design pass, now on the whole card instead of just the file_uploader
# widget, since the card itself now visually *is* the upload area.
st.html(f"""<style>
.stApp, .stApp * {{ direction: rtl; }}

.st-key-header_card {{
    background: linear-gradient(135deg, {"#1B4332"} 0%, {_ACCENT_GREEN_LIGHT} 100%);
    border-radius: 16px;
    padding: 1.25rem 1.75rem;
    margin-bottom: 0.5rem;
}}
.st-key-header_card, .st-key-header_card * {{
    color: #FAF7F0 !important;
}}

.st-key-upload_card, .st-key-results_card, .st-key-skipped_card {{
    background-color: #FFFFFF;
    border-radius: 14px;
    padding: 1.25rem 1.5rem;
    box-shadow: 0 2px 10px rgba(27, 67, 50, 0.08);
    margin-bottom: 1rem;
}}
.st-key-upload_card {{
    border: 2px dashed {_UPLOAD_AREA_BORDER};
    background-color: {_UPLOAD_AREA_BACKGROUND};
}}
</style>""")


def _check_password() -> bool:
    """Single shared-password gate (no username) for the Streamlit Community
    Cloud deployment. The password itself never lives in code - only in
    st.secrets["APP_PASSWORD"] (see .streamlit/secrets.toml.example) - and
    nothing past this function's call site renders until it's entered
    correctly, since the caller st.stop()s on a False return.
    """
    if st.session_state.get("authenticated", False):
        return True

    try:
        correct_password = st.secrets["APP_PASSWORD"]
    except Exception:
        st.error(
            "APP_PASSWORD לא הוגדר כ-secret. הגדירו אותו לפני הפעלת "
            "האפליקציה - ראו .streamlit/secrets.toml.example.",
            icon=":material/error:",
        )
        return False

    _, center, _ = st.columns([1, 2, 1])
    with center:
        with st.container(key="header_card"):
            if _LOGO_PATH.is_file():
                st.image(str(_LOGO_PATH), width=100)
            st.title(":material/lock: מערכת ריכוז תעודות פינוי", text_alignment="right")
        with st.form("login_form", border=True):
            entered_password = st.text_input("סיסמה", type="password")
            submitted = st.form_submit_button(
                "כניסה", icon=":material/login:", type="primary"
            )
        if submitted:
            if entered_password == correct_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("סיסמה שגויה.", icon=":material/error:")

    return False


if not _check_password():
    st.stop()

# --- Session state model ---------------------------------------------------
# Each project is its own separate Excel file - explicit, user-named, never
# mixed with another project's data (see requirements this implements: a
# project name field, and a clear "add more to this project" vs. "start a
# new project" choice after processing). Two lists per active project, same
# reasoning as the previous single-file design (a session must never show a
# previous project's data by default, but nothing already written to disk
# may ever be lost):
#
#   baseline_records - whatever's already on the active project's specific
#     file when it's first activated this session. Normally empty (a
#     project's filename is freshly built from its name + today's date), but
#     if the exact same project name is reused the same day, this safely
#     extends that file instead of silently overwriting it.
#   records - certificates *this session* has processed into the active
#     project since it was activated (by the first "עבד תעודות" click after
#     naming it) or since "התחל פרויקט חדש" was last clicked.
#
# Invariant kept true on every single write: file content == baseline_records
# + records. output_path is None until a project is activated - see the
# process-button handling below, and "התחל פרויקט חדש" for how it's cleared
# (never deleting the file that project already wrote - just detaching this
# session's memory from it, per the requirement that old projects stay on
# disk as history).
if "records" not in st.session_state:
    st.session_state.records = []
if "skipped" not in st.session_state:
    # Pages the model classified as fields.DOCUMENT_TYPE_OTHER - not a waste
    # certificate at all (e.g. an internal billing/credit note mixed into the
    # same PDF). Shown in its own section below, never mixed into `records`
    # or written to the project's Excel file - see app/pipeline.py.
    st.session_state.skipped = []
if "baseline_records" not in st.session_state:
    st.session_state.baseline_records = []
if "output_path" not in st.session_state:
    st.session_state.output_path = None
if "active_project_name" not in st.session_state:
    st.session_state.active_project_name = None
if "uploader_key_suffix" not in st.session_state:
    st.session_state.uploader_key_suffix = 0
if "project_name_key_suffix" not in st.session_state:
    st.session_state.project_name_key_suffix = 0
if "just_completed" not in st.session_state:
    st.session_state.just_completed = False

with st.container(key="header_card"):
    # Logo defined first so it lands on the visual right under the page's
    # RTL direction (see the RTL CSS above) - the natural "branding at the
    # reading start" position for a Hebrew header, with the title flowing to
    # its left.
    logo_col, title_col = st.columns([1, 5], vertical_alignment="center")
    with logo_col:
        if _LOGO_PATH.is_file():
            st.image(str(_LOGO_PATH), width=100)
    with title_col:
        st.title(":material/recycling: מערכת ריכוז תעודות פינוי", text_alignment="right")
    st.caption("העלאת תעודות פינוי פסולת וחילוץ אוטומטי של הנתונים לקובץ Excel מרוכז אחד.")

# Reserved now, filled in after the upload/processing section below, so the
# summary always reflects the active project's file while still rendering at the top.
summary_slot = st.container()
st.divider()

with st.container(key="upload_card"):
    st.subheader(":material/upload_file: העלאת תעודות", text_alignment="right")

    if st.session_state.output_path is None:
        # No project activated yet this session - the name is still editable.
        # It "locks in" (see below) the moment the first processing run starts.
        project_name_input = st.text_input(
            "שם פרויקט",
            value=_default_project_name(),
            key=f"project_name_{st.session_state.project_name_key_suffix}",
            help='כל פרויקט נשמר לקובץ Excel נפרד משלו. השם ננעל עם לחיצת "עבד תעודות" הראשונה.',
        )
    else:
        # Locked - shown read-only so it's clear editing it now wouldn't do
        # anything (renaming an already-written file mid-session isn't supported).
        st.caption(
            f":material/folder_open: **פרויקט פעיל:** {st.session_state.active_project_name} "
            f"&nbsp;·&nbsp; קובץ: `{st.session_state.output_path.name}`"
        )
        project_name_input = st.session_state.active_project_name

    allowed_types = sorted(ext.lstrip(".") for ext in config.SUPPORTED_EXTENSIONS)
    uploaded_files = st.file_uploader(
        "גררו לכאן קובצי תעודות, או לחצו לבחירה",
        type=allowed_types,
        accept_multiple_files=True,
        help=f"סוגי קבצים נתמכים: {', '.join(allowed_types)}",
        # Suffix bumped by "התחל פרויקט חדש" below - Streamlit treats a widget
        # with a new key as a brand-new, empty instance, which is the standard
        # way to force-clear a file_uploader's selection programmatically.
        key=f"uploader_{st.session_state.uploader_key_suffix}",
    )

    process_clicked = st.button(
        "עבד תעודות",
        icon=":material/play_arrow:",
        type="primary",
        disabled=not uploaded_files,
    )

if process_clicked:
    if not config.ANTHROPIC_API_KEY:
        st.error(
            "ANTHROPIC_API_KEY חסר - ודאו שקובץ .env בתיקיית הפרויקט מכיל אותו, "
            "ואז הפעילו את השרת מחדש.",
            icon=":material/error:",
        )
    else:
        if st.session_state.output_path is None:
            # First processing run of a fresh session/project: lock in the
            # name typed above and build this project's dedicated file path.
            # read_existing_records() here (not []) is what makes reusing an
            # existing project name safe instead of an overwrite hazard.
            chosen_name = (project_name_input or "").strip() or _default_project_name()
            st.session_state.active_project_name = chosen_name
            st.session_state.output_path = _build_output_path(chosen_name)
            st.session_state.baseline_records = read_existing_records(st.session_state.output_path)

        errors = []

        def on_error(path, exc):
            errors.append((path.name, str(exc)))

        with st.status(f"מעבד {len(uploaded_files)} תעודות...", expanded=True) as status:
            with tempfile.TemporaryDirectory() as tmp_dir:
                tmp_paths = []
                for uploaded in uploaded_files:
                    tmp_path = Path(tmp_dir) / uploaded.name
                    tmp_path.write_bytes(uploaded.getvalue())
                    tmp_paths.append(tmp_path)

                def on_progress(index, total, path):
                    status.update(label=f"מעבד ({index}/{total}): {path.name}")

                new_records, new_skipped = process_files(tmp_paths, on_progress=on_progress, on_error=on_error)

            # process_files() always returns "quantity" as a raw string (like
            # every other extracted field); the numeric column downstream
            # (the editable table, the Excel cell) needs a clean float-or-""
            # invariant instead. Parse it the same way write_records() will
            # anyway, so the type is consistent from here on.
            for record in new_records:
                raw_qty = record.get("quantity", "")
                parsed_qty = parse_quantity(raw_qty)
                if parsed_qty is not None:
                    record["quantity"] = parsed_qty
                elif raw_qty:
                    # The model is instructed to return quantity as digits
                    # only (or blank) - a non-empty value that still doesn't
                    # parse means it ignored that instruction. Don't just
                    # blank it silently: note what it actually said, then
                    # blank it, so nothing is lost.
                    note = f"כמות לא תקינה: {raw_qty}"
                    record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note
                    record["quantity"] = ""
                else:
                    record["quantity"] = ""

            st.session_state.records.extend(new_records)
            st.session_state.skipped.extend(new_skipped)
            # Keep the in-memory order matching what write_records() below is
            # about to produce on disk (נמוכה first) - otherwise downloading
            # right after processing would reorder rows relative to what was
            # just shown on screen, which reads as if the download lost/
            # scrambled something.
            st.session_state.records[:] = sort_by_confidence(st.session_state.records)
            # baseline_records + records, always - see the session-state-model
            # comment near the top of this file. Never write records alone;
            # that would overwrite this project's file with only this
            # session's own batch and silently drop whatever was already in it.
            write_records(
                st.session_state.baseline_records + st.session_state.records,
                st.session_state.output_path,
            )
            skip_note = f", {len(new_skipped)} דולגו כלא-תעודות" if new_skipped else ""
            status.update(
                label=f"הושלם - עובדו {len(new_records)} תעודות חדשות{skip_note}", state="complete"
            )
            st.session_state.just_completed = True

        for filename, error in errors:
            st.warning(f"שגיאה בעיבוד {filename}: {error}", icon=":material/warning:")

# --- Explicit "what next" choice, right after a batch finishes -------------
if st.session_state.just_completed and (st.session_state.records or st.session_state.skipped):
    st.success(
        f'העיבוד הושלם ונשמר בקובץ "{st.session_state.output_path.name}". מה ברצונך לעשות?',
        icon=":material/check_circle:",
    )
    with st.container(horizontal=True):
        add_more_clicked = st.button(
            "הוסף עוד תעודות לפרויקט הזה", icon=":material/add:", type="primary"
        )
        new_project_clicked = st.button(
            "התחל פרויקט חדש", icon=":material/create_new_folder:", type="secondary"
        )
    if add_more_clicked:
        # Nothing about the active project changes - just dismiss this
        # banner so the uploader above is ready for the next batch.
        st.session_state.just_completed = False
        st.rerun()
    if new_project_clicked:
        # The just-finished project is already safely written to its own
        # file (every processing run writes immediately - see above), so
        # starting fresh only needs to detach this session's memory from it,
        # never a fold-forward like the old single-shared-file design had.
        st.session_state.records = []
        st.session_state.skipped = []
        st.session_state.baseline_records = []
        st.session_state.output_path = None
        st.session_state.active_project_name = None
        st.session_state.just_completed = False
        st.session_state.uploader_key_suffix += 1
        st.session_state.project_name_key_suffix += 1
        st.rerun()

# --- Live results table + Excel download ---------------------------------
st.divider()

if not st.session_state.records:
    st.subheader(":material/checklist: תוצאות", text_alignment="right")
    st.info(
        'טרם עובדו תעודות בפרויקט זה. תנו שם לפרויקט למעלה (או השאירו את ברירת '
        'המחדל), גררו קבצים לתיבת ההעלאה ולחצו על "עבד תעודות".',
        icon=":material/info:",
    )
    with summary_slot:
        with st.container(horizontal=True, horizontal_alignment="right"):
            st.metric('סה"כ תעודות בפרויקט', 0)
            st.metric("מסומנות לבדיקה ידנית", 0)
else:
    header_by_key = dict(EXCEL_COLUMNS)
    keys = [key for key, _label in EXCEL_COLUMNS]
    # "status_symbol" is a display-only column (not one of EXCEL_COLUMNS, not
    # touched by the merge-back loop below) - a second, symbol-based signal
    # for רמת ביטחון alongside its own text value and the existing
    # background-fill coloring, per the design request. Inserted right
    # before confidence in column order, wherever that lands in EXCEL_COLUMNS.
    _STATUS_KEY = "status_symbol"
    display_keys = []
    for key in keys:
        if key == "confidence":
            display_keys.append(_STATUS_KEY)
        display_keys.append(key)
    header_by_key[_STATUS_KEY] = "מצב"

    with st.container(key="results_card"):
        st.subheader(":material/checklist: תוצאות", text_alignment="right")

        filter_choice = st.segmented_control(
            "הצג",
            options=["הכל", "נמוכה ובינונית", "רק נמוכה"],
            default="הכל",
            required=True,
            label_visibility="collapsed",
        )
        _levels_by_filter = {
            "הכל": set(CONFIDENCE_LEVELS),
            "נמוכה ובינונית": {"נמוכה", "בינונית"},
            "רק נמוכה": {"נמוכה"},
        }
        visible_levels = _levels_by_filter.get(filter_choice, set(CONFIDENCE_LEVELS))
        filtered_indices = [
            i for i, r in enumerate(st.session_state.records) if r.get("confidence") in visible_levels
        ]
        visible_records = [st.session_state.records[i] for i in filtered_indices]

        def _selectbox_options(closed_list, current_values):
            # Always offer the closed list, plus any value already present in
            # the visible rows that isn't on it. _flag_out_of_list_values()
            # (in extractor.py) deliberately *keeps* an out-of-list
            # waste_type/region on the record instead of blanking it,
            # precisely so a reviewer can see what the model actually
            # returned - constraining the dropdown to only the closed list
            # would silently discard that value the moment this table renders.
            extra = sorted({v for v in current_values if v and v not in closed_list})
            return [""] + list(closed_list) + extra

        # __idx__ carries each row's position in st.session_state.records
        # through the editor and back - hidden from view via column_config
        # below, but still round-trips in the returned data (per
        # st.column_config's own "hides from the UI, data still there"
        # behavior). Matching edits back by this explicit id, rather than by
        # row position, means the merge below stays correct even if a future
        # Streamlit version adds interactive sorting to st.data_editor.
        rows = [
            {
                **{k: r.get(k, "") for k in keys},
                _STATUS_KEY: _STATUS_SYMBOLS.get(r.get("confidence"), ""),
                "__idx__": i,
            }
            for i, r in zip(filtered_indices, visible_records)
        ]
        df = pd.DataFrame(rows, columns=display_keys + ["__idx__"])
        df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
        df = df.rename(columns=header_by_key)

        def _highlight_source_cell(row):
            # Pandas Styler styles only survive on non-editable columns in
            # st.data_editor (documented behavior) - "קובץ מקור" and "מצב"
            # are both disabled below anyway (editing "קובץ מקור" would break
            # its role as the row's link back to the original file; "מצב" is
            # a computed display-only symbol), which happens to be exactly
            # the two columns this can still color.
            confidence = row["רמת ביטחון"]
            color = _ROW_COLORS.get(confidence)
            styles = []
            for col in row.index:
                if color and col == "קובץ מקור":
                    styles.append(f"background-color: {color}")
                elif col == "מצב" and confidence == "גבוהה":
                    # The "high confidence" accent border from the design
                    # request - the editable columns can't carry a Styler
                    # style at all (see the docstring note above), so this
                    # read-only symbol column is where it actually can render.
                    styles.append(f"border: 2px solid {_ACCENT_GOLD}; background-color: #FCF3DE")
                else:
                    styles.append("")
            return styles

        styled = df.style.apply(_highlight_source_cell, axis=1)

        edited_df = st.data_editor(
            styled,
            key="results_editor",
            hide_index=True,
            column_config={
                "__idx__": None,
                "מצב": st.column_config.TextColumn(disabled=True, width="small"),
                "קובץ מקור": st.column_config.TextColumn(disabled=True),
                "כמות": st.column_config.NumberColumn(format="%,.2f"),
                "רמת ביטחון": st.column_config.SelectboxColumn(options=CONFIDENCE_LEVELS),
                "סוג הפסולת": st.column_config.SelectboxColumn(
                    options=_selectbox_options(
                        WASTE_TYPES + [UNCLASSIFIED_WASTE_TYPE], (r.get("waste_type") for r in visible_records)
                    )
                ),
                "מרחב": st.column_config.SelectboxColumn(
                    options=_selectbox_options(REGIONS, (r.get("region") for r in visible_records))
                ),
                "חודש": st.column_config.SelectboxColumn(
                    options=_selectbox_options(HEBREW_MONTHS, (r.get("month") for r in visible_records))
                ),
            },
        )

    # Merge edits back into session_state.records by __idx__, then keep the
    # file in sync with whatever's now on screen - every rerun, not just
    # ones an edit triggered, since that's simpler and correct either way
    # (a no-op write when nothing actually changed) rather than trying to
    # detect which reruns need it.
    for _, edited_row in edited_df.iterrows():
        target = st.session_state.records[int(edited_row["__idx__"])]
        for key in keys:
            value = edited_row.get(header_by_key[key])
            if key == "quantity":
                target[key] = "" if pd.isna(value) else float(value)
            else:
                target[key] = "" if value is None else value

    st.session_state.records[:] = sort_by_confidence(st.session_state.records)
    # baseline_records + records, same as after processing above - an edit
    # here must not overwrite the project's file with only what's on screen.
    all_records = st.session_state.baseline_records + st.session_state.records
    write_records(all_records, st.session_state.output_path)

    with summary_slot:
        with st.container(horizontal=True, horizontal_alignment="right"):
            st.metric('סה"כ תעודות בפרויקט', len(all_records))
            st.metric(
                "מסומנות לבדיקה ידנית",
                sum(1 for r in all_records if r.get("confidence") in _FLAGGED_LEVELS),
            )

    st.download_button(
        "הורדת קובץ Excel",
        data=st.session_state.output_path.read_bytes(),
        file_name=st.session_state.output_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
        type="primary",
    )

    # --- Source page viewer, next to the row it came from --------------------
    st.divider()
    st.subheader(":material/description: תצוגת עמוד מקור", text_alignment="right")

    row_labels = [
        f"{i + 1}. {r.get('source_file', '')} - {r.get('site') or r.get('waste_type') or '(ללא זיהוי)'}"
        for i, r in enumerate(st.session_state.records)
    ]
    selected_label = st.selectbox("בחרו שורה לצפייה בעמוד המקורי", options=row_labels)
    if selected_label:
        selected_record = st.session_state.records[row_labels.index(selected_label)]
        image_bytes = selected_record.get("_page_image_jpeg")
        image_col, fields_col = st.columns(2)
        with image_col:
            if image_bytes:
                try:
                    st.image(image_bytes)
                except Exception:
                    st.warning("לא ניתן להציג את התמונה השמורה לשורה זו.", icon=":material/warning:")
            else:
                st.info(
                    "אין תמונה שמורה לשורה זו - היא עובדה בהרצה קודמת של השרת "
                    "(תמונות נשמרות רק בזיכרון, לאורך ההרצה הנוכחית בלבד).",
                    icon=":material/info:",
                )
        with fields_col:
            for key, label in EXCEL_COLUMNS:
                st.write(f"**{label}:** {selected_record.get(key) or '—'}")

# --- Skipped non-certificate documents --------------------------------------
# Deliberately outside the records if/else above (renders even when a whole
# batch turned out to be entirely non-certificates and `records` is empty),
# and deliberately never mixed into the results table or the Excel file - see
# app/pipeline.py's ProcessResult. A row here means the model looked at the
# page and actively decided it isn't a waste certificate at all (a billing
# note, a work order...); that's a fundamentally different situation from a
# row in the results table with "נמוכה" confidence, which IS a certificate
# the model tried and struggled with. Mixing the two would make the "needs
# review" count above misleading either way.
if st.session_state.skipped:
    st.divider()
    with st.container(key="skipped_card"):
        st.subheader(
            f":material/block: מסמכים שדולגו ({len(st.session_state.skipped)})", text_alignment="right"
        )
        st.caption(
            ":gray[עמודים שזוהו כמסמכים שאינם תעודות שקילה/פינוי כלל (למשל תעודות חיוב/"
            "זיכוי פנימיות) - לא נכשלו בחילוץ, ולכן לא נספרים בתקציר למעלה ולא נכתבים "
            "לקובץ ה-Excel. מוצגים כאן רק למידע, כדי שיהיה ברור שהם לא אבדו בטעות.]"
        )
        skipped_table = pd.DataFrame(
            [
                {"קובץ מקור": r.get("source_file", ""), "סיבה": r.get("notes", "") or "-"}
                for r in st.session_state.skipped
            ]
        )
        st.table(skipped_table, hide_index=True)
