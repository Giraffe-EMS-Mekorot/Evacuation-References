"""Web UI for the tteudot-agent certificate pipeline - runnable locally or
deployed to Streamlit Community Cloud (see README.md).

An alternative front end alongside main.py's CLI, not a replacement for it:
upload files here (drag-and-drop) instead of dropping them into input/, and
get the same output/ריכוז_תעודות.xlsx - built through the identical
app.pipeline.process_files() the CLI uses, so the two front ends can never
disagree on how a certificate is read or how a bad file is handled.

Gated behind a single shared password (see _check_password() below) - never
hardcoded, always read from st.secrets["APP_PASSWORD"].

Run with:
    streamlit run streamlit_app.py
"""
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from app import config
from app.excel_writer import parse_quantity, read_existing_records, write_records
from app.fields import EXCEL_COLUMNS
from app.pipeline import process_files

_OUTPUT_PATH = config.OUTPUT_DIR / config.OUTPUT_FILENAME
_FLAGGED_LEVELS = {"נמוכה", "בינונית"}
# Matches the row-fill colors excel_writer.py uses for these same confidence
# levels in the downloaded file, so the on-screen table and the spreadsheet
# always agree visually.
_ROW_COLORS = {"נמוכה": "#FFC7CE", "בינונית": "#FFEB9C"}

st.set_page_config(
    page_title="מערכת ריכוז תעודות פינוי",
    page_icon=":material/recycling:",
    layout="wide",
)

# Streamlit has no built-in RTL layout mode; this is the minimal CSS needed to
# mirror the app right-to-left for Hebrew. This is a functional requirement,
# not decorative styling - all actual visual theming (colors, fonts, radius)
# lives in .streamlit/config.toml instead, per Streamlit's native theming.
st.html("<style>.stApp, .stApp * { direction: rtl; }</style>")


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

if "records" not in st.session_state:
    # Seed from whatever's already in the output file (from an earlier
    # browser session, or from main.py's CLI) so this session's first "עבד
    # תעודות" click extends that history instead of silently overwriting it.
    st.session_state.records = read_existing_records(_OUTPUT_PATH)

st.title(":material/recycling: מערכת ריכוז תעודות פינוי", text_alignment="right")
st.caption("העלאת תעודות פינוי פסולת וחילוץ אוטומטי של הנתונים לקובץ Excel מרוכז אחד.")

# Reserved now, filled in after the upload/processing section below, so the
# summary always reflects this run's results while still rendering at the top.
summary_slot = st.container()
st.divider()

allowed_types = sorted(ext.lstrip(".") for ext in config.SUPPORTED_EXTENSIONS)
uploaded_files = st.file_uploader(
    "גררו לכאן קובצי תעודות, או לחצו לבחירה",
    type=allowed_types,
    accept_multiple_files=True,
    help=f"סוגי קבצים נתמכים: {', '.join(allowed_types)}",
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

                new_records = process_files(tmp_paths, on_progress=on_progress, on_error=on_error)

            # extract_certificate() always returns "quantity" as a raw string
            # (like every other field); read_existing_records() returns it as
            # a float for cells write_records() already parsed. Left alone,
            # extending session_state.records mixes both types in the same
            # column - pandas/pyarrow can't serialize that for st.table and
            # falls back to a broken, unformatted display. Parse it the same
            # way write_records() will anyway, so the type is consistent from
            # here on regardless of which path a record came from.
            for record in new_records:
                parsed_qty = parse_quantity(record.get("quantity", ""))
                record["quantity"] = parsed_qty if parsed_qty is not None else record.get("quantity", "")

            st.session_state.records.extend(new_records)
            write_records(st.session_state.records, _OUTPUT_PATH)
            status.update(label=f"הושלם - עובדו {len(new_records)} תעודות חדשות", state="complete")

        for filename, error in errors:
            st.warning(f"שגיאה בעיבוד {filename}: {error}", icon=":material/warning:")

# --- Quick summary (rendered into the slot reserved above the upload area) ---
total_certs = len(st.session_state.records)
flagged_certs = sum(1 for r in st.session_state.records if r.get("confidence") in _FLAGGED_LEVELS)
with summary_slot:
    with st.container(horizontal=True, horizontal_alignment="right"):
        st.metric('סה"כ תעודות בקובץ', total_certs)
        st.metric("מסומנות לבדיקה ידנית", flagged_certs)

# --- Live results table + Excel download ---------------------------------
st.divider()
st.subheader("תוצאות", text_alignment="right")

if not st.session_state.records:
    st.info(
        'טרם עובדו תעודות. גררו קבצים לתיבת ההעלאה למעלה ולחצו על "עבד תעודות".',
        icon=":material/info:",
    )
else:
    header_by_key = dict(EXCEL_COLUMNS)
    keys = [key for key, _label in EXCEL_COLUMNS]
    df = pd.DataFrame(st.session_state.records, columns=keys).rename(columns=header_by_key)

    def _highlight_row(row):
        color = _ROW_COLORS.get(row["רמת ביטחון"])
        return [f"background-color: {color}" if color else "" for _ in row]

    def _format_quantity(value):
        # Mirrors excel_writer.py's "#,##0.##" number format: thousands
        # separator, up to 2 decimals, no trailing zeros. Blank/unparseable
        # values become "" rather than passing the raw value through.
        if isinstance(value, (int, float)):
            return f"{value:,.2f}".rstrip("0").rstrip(".")
        return "" if value in (None, "") else str(value)

    # Applied eagerly (not via Styler.format) so the column ends up uniformly
    # str, not a mix of float and str/None - a mixed-type object column is
    # exactly what breaks st.table's Arrow serialization the moment a batch
    # includes both an already-parsed quantity (from a prior run) and a blank
    # or unparseable one (a record missing this core field is common and
    # expected, not an edge case) sitting in the same column.
    df["כמות"] = df["כמות"].map(_format_quantity)

    styled = df.style.apply(_highlight_row, axis=1)
    st.table(styled, hide_index=True)

    st.download_button(
        "הורדת קובץ Excel",
        data=_OUTPUT_PATH.read_bytes(),
        file_name=config.OUTPUT_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
        type="primary",
    )
