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
from app.duplicate_check import flag_potential_duplicates
from app.excel_input import read_manual_excel
from app.excel_writer import (
    parse_quantity,
    read_existing_records,
    read_existing_weighing_certificates,
    sort_by_confidence,
    write_records,
)
from app.fields import (
    BATCH_SELECTION_KEYS,
    CONFIDENCE_LEVELS,
    EXCEL_COLUMNS,
    INTERNAL_TRACKING_FIELDS,
    REGIONS,
    UNCLASSIFIED_WASTE_TYPE,
    WASTE_TYPES,
)
from app.pipeline import BatchSelection, process_files
from app import sites_config

# Manually-filled tracking sheets, read directly (no vision call) - see
# app/excel_input.py. A DIFFERENT feature from config.SUPPORTED_EXTENSIONS,
# which is main.py's CLI file-type filter for input/ and deliberately left
# untouched: the CLI has no code path for this at all, only this file's
# upload area does (see the process_clicked handling below, which branches
# by extension instead of sending a .xlsx through extract_certificate_pages).
_MANUAL_EXCEL_EXTENSIONS = {".xlsx"}

_FLAGGED_LEVELS = {"נמוכה", "בינונית"}
# Matches the row-fill colors excel_writer.py uses for these same confidence
# levels in the downloaded file, so the on-screen table and the spreadsheet
# always agree visually.
_ROW_COLORS = {"נמוכה": "#FFC7CE", "בינונית": "#FFEB9C"}
_LOGO_PATH = Path(__file__).resolve().parent / "logo.jpg"

# --- Design tokens (2026-09 bolder-palette pass) ----------------------------
# Deep green, matching config.toml's primaryColor - kept as its own constant
# (not read back from theme) since Python has no access to the active theme
# at render time, same reasoning as the pre-existing _ACCENT_GREEN_LIGHT
# below. Used directly wherever a literal color is needed in the CSS/HTML
# blocks in this file (gradients, card backgrounds, icon badges).
_ACCENT_GREEN_DARK = "#12291C"
# A lighter step of the same green, for the header banner's gradient depth
# and for text-on-gold contexts - not a theme token (config.toml has no
# "secondary green" slot), recomputed here to stay in proportion to the
# deeper _ACCENT_GREEN_DARK above (the old #3D6B4F was tuned against the
# previous, lighter #1B4332 primary and would look too pale against this one).
_ACCENT_GREEN_LIGHT = "#2F5940"
# Cream - matches config.toml's backgroundColor; kept as its own constant
# since it's used as an explicit foreground/fill color in several places
# below (circular-logo backing, card text, watermark SVG), not just as the
# page background config.toml already handles on its own.
_CREAM = "#FAF7F0"
_CREAM_URL = _CREAM.replace("#", "%23")  # '#' must be percent-escaped inside a data: URI
# Bolder gold accent (2026-09 request) - matches config.toml's orangeColor,
# used both there (so Streamlit's own orange-tinted elements pick it up) and
# here directly for every other bit of custom CSS/HTML that needs the
# literal value: the upload-area border, the stat-card/upload-badge accents,
# the primary-button border, and the high-confidence accent border in the
# results table (see _highlight_source_cell).
_ACCENT_GOLD = "#D4A24C"
# The upload area's accent border/background, per the design request - not
# in .streamlit/config.toml because there's no theme token for "this one
# specific widget's background"; see the CSS block below for why this
# targets the card's own key rather than a fixed/generic selector.
_UPLOAD_AREA_BORDER = _ACCENT_GOLD
_UPLOAD_AREA_BACKGROUND = "#FBF1DE"
# "confidence" is no longer part of fields.EXCEL_COLUMNS (2026-09-01 - see
# that list's own comment: it moved to the output file's "מעקב פנימי" sheet
# instead), but this review UI still needs to show/edit it directly - core
# reviewer workflow, unrelated to the written file's own column layout. Used
# below to re-insert it into the results table and the source-page-viewer's
# field list, both of which used to get it for free from EXCEL_COLUMNS.
_CONFIDENCE_COLUMN = ("confidence", "רמת ביטחון")
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
# block IS deliberately decorative (the "card" look, the header gradient,
# the 2026-09 bolder-palette pass below), targeted at specific containers'
# own st.container(key=...) values rather than generic selectors that would
# repaint unrelated widgets:
#
#   header_card   - the logo+title+caption banner (both here and in the
#                   password screen - the two never render in the same
#                   script run, so reusing one key is safe). Dark gradient
#                   background needs the title/caption text forced light -
#                   config.toml's :green[]/:gray[] markdown colors are for
#                   the plain cream background elsewhere, not readable here.
#   login_logo / main_logo - the same header_card's logo image, just two
#                   different sizes (see _check_password() and the main
#                   header section) - a shared img-styling rule (circular,
#                   gold ring) keyed to both.
#   login_header_title - wraps ONLY the login screen's st.title() call, so
#                   the 32px override below never touches the main screen's
#                   title (still the theme's default h1 size, per the
#                   design request: main screen gets "the same title, just
#                   the new background").
#   upload_card / results_card / skipped_card - plain "card" containers:
#                   distinct background, soft border, subtle shadow. Cheap
#                   visual separation Streamlit's own border=True can't give
#                   (it only draws a border in the page's own background).
#
# upload_card additionally gets the gold dashed border from the earlier
# design pass, now on the whole card instead of just the file_uploader
# widget, since the card itself now visually *is* the upload area.
#
# .stat-card* / .upload-badge - custom-HTML replacements for st.metric() and
# the file_uploader's own (icon-less, in this Streamlit version) dropzone -
# see _render_stat_cards() and the upload_card block below for why plain
# Streamlit widgets can't get a per-instance colored background + icon.
# Their icon glyphs use "Material Symbols Rounded" - the exact font
# Streamlit's own :material/...: icons already render with (confirmed in
# its bundled, locally-served font file - no extra network request), so a
# literal icon name like "cloud_upload" renders as the matching glyph and
# stays visually consistent with every other icon in the app.
#
# [data-testid="stBaseButton-primary"] - Streamlit's own stable test-id for
# every type="primary" button (confirmed against the installed version's
# source), used here instead of a class name so this survives a Streamlit
# upgrade's internal class-name churn better than most alternatives would.
st.html(f"""<style>
.stApp, .stApp * {{ direction: rtl; }}

.st-key-header_card {{
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, {_ACCENT_GREEN_DARK} 0%, {_ACCENT_GREEN_LIGHT} 100%);
    border-radius: 16px;
    padding: 1.25rem 1.75rem;
    margin-bottom: 0.5rem;
}}
.st-key-header_card, .st-key-header_card * {{
    color: {_CREAM} !important;
}}
/* Subtle leaf watermark - purely decorative texture, kept at very low
   opacity and positioned away from the logo/title (physical left, which
   under this page's RTL direction is the side opposite the branding) so it
   never competes with real content for attention. z-index:-1 relies on
   position:relative above to keep it confined behind this card's own
   normal-flow content, not the page underneath it. */
.st-key-header_card::before {{
    content: "";
    position: absolute;
    left: -55px;
    top: -65px;
    width: 320px;
    height: 320px;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Cpath d='M50 2C18 20 4 52 50 98C96 52 82 20 50 2Z' fill='{_CREAM_URL}'/%3E%3Cpath d='M50 12V88' stroke='{_CREAM_URL}' stroke-width='1.5'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-size: contain;
    opacity: 0.07;
    z-index: -1;
    pointer-events: none;
}}

/* Circular, gold-ringed logo badge - object-fit:contain (not cover) is
   deliberate: logo.jpg is a wide wordmark, not a compact mark, so a tight
   circular CROP would clip the lettering at both edges. Framing it inside a
   slightly-larger circle instead keeps the whole logo intact. */
.st-key-login_logo img, .st-key-main_logo img {{
    border-radius: 50%;
    border: 2.5px solid {_ACCENT_GOLD};
    background: {_CREAM};
    object-fit: contain;
    box-sizing: border-box;
}}
.st-key-login_logo img {{ padding: 16px; }}
.st-key-main_logo img {{ padding: 10px; }}

/* Flex+gap (not relying on each child's own default margin, which
   Streamlit sets per element and varies by element type) keeps the eyebrow
   label, title, and underline tight against each other regardless of
   Streamlit's own internal spacing for a heading vs. a raw st.html() block. */
.st-key-login_header_title [data-testid="stVerticalBlock"] {{
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 2px;
}}
.st-key-login_header_title [data-testid="stElementContainer"] {{
    margin: 0 !important;
}}
.st-key-login_header_title h1 {{
    font-size: 32px !important;
    margin: 0 !important;
}}
/* More specific + also !important, so this beats the blanket
   ".st-key-header_card *{{color:cream !important}}" rule above for just
   these two decorative bits (the gold eyebrow label and underline). */
.st-key-header_card .login-eyebrow {{
    color: {_ACCENT_GOLD} !important;
    font-size: 13px;
    font-weight: 600;
    letter-spacing: 0.09em;
    margin-bottom: 0.2rem;
}}
.st-key-header_card .login-underline {{
    background: {_ACCENT_GOLD} !important;
    width: 60px;
    height: 3px;
    margin-top: 0.1rem;
}}

.st-key-upload_card, .st-key-results_card, .st-key-skipped_card {{
    background-color: #FFFFFF;
    border-radius: 14px;
    padding: 1.25rem 1.5rem;
    box-shadow: 0 2px 10px rgba(18, 41, 28, 0.10);
    margin-bottom: 1rem;
}}
.st-key-upload_card {{
    border: 2px dashed {_UPLOAD_AREA_BORDER};
    background-color: {_UPLOAD_AREA_BACKGROUND};
}}
.upload-badge {{
    display: flex;
    align-items: center;
    justify-content: center;
    width: 64px;
    height: 64px;
    border-radius: 50%;
    background: {_ACCENT_GREEN_DARK};
    margin: 0 auto 0.75rem auto;
}}
.upload-badge span {{
    font-family: "Material Symbols Rounded";
    font-size: 30px;
    color: {_ACCENT_GOLD};
    line-height: 1;
}}

.stat-card-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0.9rem;
    width: 100%;
}}
.stat-card {{
    display: flex;
    align-items: center;
    gap: 0.75rem;
    border-radius: 12px;
    padding: 1rem 1.1rem;
}}
.stat-card__icon {{
    display: flex;
    align-items: center;
    justify-content: center;
    width: 42px;
    height: 42px;
    border-radius: 50%;
    flex-shrink: 0;
}}
.stat-card__icon span {{
    font-family: "Material Symbols Rounded";
    font-size: 24px;
    line-height: 1;
}}
.stat-card__value {{
    font-size: 26px;
    font-weight: 700;
    line-height: 1.15;
}}
.stat-card__label {{
    font-size: 13px;
    font-weight: 500;
    opacity: 0.9;
}}
.stat-card--green {{
    background: {_ACCENT_GREEN_DARK};
    color: {_CREAM};
}}
.stat-card--green .stat-card__icon {{ background: rgba(250, 247, 240, 0.14); }}
.stat-card--green .stat-card__icon span {{ color: {_ACCENT_GOLD}; }}
.stat-card--gold {{
    background: {_ACCENT_GOLD};
    color: {_ACCENT_GREEN_DARK};
}}
.stat-card--gold .stat-card__icon {{ background: rgba(18, 41, 28, 0.12); }}
.stat-card--gold .stat-card__icon span {{ color: {_ACCENT_GREEN_DARK}; }}

[data-testid="stBaseButton-primary"] {{
    border: 2px solid {_ACCENT_GOLD} !important;
    border-radius: 10px !important;
    color: {_CREAM} !important;
}}
</style>""")


def _stat_card_html(value, label: str, icon: str, variant: str) -> str:
    return f"""
    <div class="stat-card stat-card--{variant}">
        <div class="stat-card__icon"><span>{icon}</span></div>
        <div>
            <div class="stat-card__value">{value}</div>
            <div class="stat-card__label">{label}</div>
        </div>
    </div>
    """


def _render_stat_cards(total: int, flagged: int) -> None:
    """Renders the two summary numbers as a pair of custom, high-contrast
    cards (deep green / bold gold) instead of st.metric() - a pure rendering
    swap for the 2026-09 design pass, not a data change: both call sites
    below compute `total`/`flagged` exactly as before (len(...) and the same
    confidence-filtered sum()), this function only decides how to display
    them. st.metric() has no per-instance background-color or icon support,
    which the design explicitly calls for here, hence custom HTML via
    st.html() instead.
    """
    st.html(f"""
    <div class="stat-card-grid">
        {_stat_card_html(total, 'סה"כ תעודות בפרויקט', "description", "green")}
        {_stat_card_html(flagged, "מסומנות לבדיקה ידנית", "warning", "gold")}
    </div>
    """)


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
                with st.container(key="login_logo"):
                    st.image(str(_LOGO_PATH), width=120)
            with st.container(key="login_header_title"):
                st.html('<div class="login-eyebrow">כניסה מאובטחת</div>')
                st.title(":material/lock: מערכת ריכוז תעודות פינוי", text_alignment="right")
                st.html('<div class="login-underline"></div>')
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
if "weighing_certificates" not in st.session_state:
    # Separate תעודת שקילה pages (fields.CERT_ROLE_WEIGHING_CERTIFICATE) from
    # this session's batches. Never rows of their own - and therefore
    # deliberately NOT in `records`, which is what keeps them out of the
    # results table below with no filtering there at all. Held in state only
    # so every later write_records() call (e.g. after an inline edit) can
    # keep reporting the unmatched ones on the "מעקב פנימי" sheet instead of
    # losing them after the first write. See app/pipeline.py.
    st.session_state.weighing_certificates = []
if "batch_region" not in st.session_state:
    # The per-batch מרחב/אתר/סוג-פסולת selection (2026-09-15). These three
    # columns are no longer read off the documents at all - see
    # app/pipeline.py's apply_batch_selection - so this selection is their
    # only source for every row a new batch produces. Kept in session state
    # (not just read from the widgets) so the cascading reset below can clear
    # the site when the מרחב changes.
    st.session_state.batch_region = None
if "batch_site" not in st.session_state:
    st.session_state.batch_site = None
if "batch_waste_type" not in st.session_state:
    st.session_state.batch_waste_type = None
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
            with st.container(key="main_logo"):
                st.image(str(_LOGO_PATH), width=80)
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

    st.divider()

    # --- Per-batch מרחב / אתר / סוג פסולת selection (2026-09-15) ----------
    # These three values are applied to every row this batch produces and are
    # NOT read from the documents (see app/pipeline.py apply_batch_selection).
    # All three are therefore required before anything can be uploaded -
    # otherwise the rows would be written with three permanently blank
    # columns and no way to recover them from the source files.
    st.markdown("##### :material/checklist_rtl: שיוך האצווה")
    st.caption(
        "השיוך חל על **כל הקבצים שמועלים יחד** באצווה הזו, ולא על קובץ בודד. "
        "שלושת הערכים נלקחים מהבחירה כאן ולא נקראים מהמסמכים."
    )

    sel_col1, sel_col2, sel_col3 = st.columns(3)

    with sel_col1:
        region = st.selectbox(
            "מרחב",
            options=sites_config.REGION_NAMES,
            index=None,
            placeholder="בחרו מרחב...",
            key="batch_region_widget",
        )

    # A מרחב change must not leave the previously-picked site behind (it would
    # belong to a different מרחב) - sites_config.is_valid_selection() would
    # reject it, so clear it here rather than silently blocking the button.
    if region != st.session_state.batch_region:
        st.session_state.batch_region = region
        st.session_state.batch_site = None

    site_options = sites_config.sites_for(region) if region else []
    has_sites = bool(region) and sites_config.region_has_sites(region)

    with sel_col2:
        if has_sites:
            site = st.selectbox(
                "אתר / יחידה",
                options=site_options,
                index=None,
                placeholder="בחרו אתר...",
                key=f"batch_site_widget_{region}",
            )
        elif region:
            # חטיבת הפיתוח has no second level at all - a blank site is the
            # correct and complete answer here, so show why instead of an
            # empty dropdown the user would try to fill.
            site = ""
            st.selectbox(
                "אתר / יחידה",
                options=["(לא רלוונטי למרחב זה)"],
                index=0,
                disabled=True,
                key=f"batch_site_disabled_{region}",
                help=f"למרחב {region} אין אתרים משניים - השדה יישאר ריק, וזו בחירה שלמה.",
            )
        else:
            site = None
            st.selectbox(
                "אתר / יחידה",
                options=["בחרו מרחב תחילה"],
                index=0,
                disabled=True,
                key="batch_site_placeholder",
            )
    st.session_state.batch_site = site

    with sel_col3:
        # The already-existing closed waste-type list, not a second copy of it.
        waste_type = st.selectbox(
            "סוג פסולת",
            options=WASTE_TYPES,
            index=None,
            placeholder="בחרו סוג פסולת...",
            key="batch_waste_type_widget",
        )
    st.session_state.batch_waste_type = waste_type

    selection_valid = (
        bool(region)
        and bool(waste_type)
        and sites_config.is_valid_selection(region, site or "")
    )

    missing = []
    if not region:
        missing.append("מרחב")
    elif has_sites and not site:
        missing.append("אתר / יחידה")
    if not waste_type:
        missing.append("סוג פסולת")

    if not selection_valid:
        st.warning(
            "יש להשלים את שיוך האצווה לפני העלאת קבצים - חסר: "
            + ", ".join(missing)
            + ".",
            icon=":material/warning:",
        )
    else:
        st.success(
            "שיוך האצווה: **"
            + region
            + ("** · **" + site + "**" if site else "** (ללא אתר משני)")
            + " · **"
            + waste_type
            + "**",
            icon=":material/check_circle:",
        )

    st.divider()

    # Purely decorative - this Streamlit version's file_uploader dropzone has
    # no illustrative icon of its own (confirmed against the installed
    # version's source), just text + a small "Browse files" button; this
    # circle+icon badge is added independently, above the real widget, not a
    # restyle of anything Streamlit itself renders.
    st.html('<div class="upload-badge"><span>cloud_upload</span></div>')
    allowed_types = sorted(ext.lstrip(".") for ext in config.SUPPORTED_EXTENSIONS | _MANUAL_EXCEL_EXTENSIONS)
    uploaded_files = st.file_uploader(
        "גררו לכאן קובצי תעודות, או לחצו לבחירה (כולל קובצי Excel ממולאים ידנית)",
        disabled=not selection_valid,
        type=allowed_types,
        accept_multiple_files=True,
        help=f"סוגי קבצים נתמכים: {', '.join(allowed_types)}. קובץ xlsx נקרא ישירות "
        "כגיליון ממולא ידנית, ולא נשלח למודל.",
        # Suffix bumped by "התחל פרויקט חדש" below - Streamlit treats a widget
        # with a new key as a brand-new, empty instance, which is the standard
        # way to force-clear a file_uploader's selection programmatically.
        key=f"uploader_{st.session_state.uploader_key_suffix}",
    )

    process_clicked = st.button(
        "עבד תעודות",
        icon=":material/play_arrow:",
        type="primary",
        disabled=not uploaded_files or not selection_valid,
        help=None if selection_valid else "יש להשלים את שיוך האצווה למעלה (מרחב, אתר, סוג פסולת).",
    )

if process_clicked:
    # Captured once for the whole batch, from the selection card above - the
    # only source for the מרחב/אתר/סוג-פסולת columns now (see
    # app/pipeline.py's apply_batch_selection). Manual-Excel rows are
    # deliberately NOT given this: those come from a human-filled sheet that
    # already carries its own values, and overwriting them would discard real
    # typed data (app/excel_input.py's whole premise).
    batch_selection = BatchSelection(
        region=st.session_state.batch_region or "",
        site=st.session_state.batch_site or "",
        waste_type=st.session_state.batch_waste_type or "",
    )

    # .xlsx uploads never touch the model at all (see app/excel_input.py) -
    # split them out first so an API-key check below only blocks the batch
    # when it actually needs Claude for something.
    manual_uploads = [f for f in uploaded_files if Path(f.name).suffix.lower() in _MANUAL_EXCEL_EXTENSIONS]
    vision_uploads = [f for f in uploaded_files if Path(f.name).suffix.lower() not in _MANUAL_EXCEL_EXTENSIONS]

    if vision_uploads and not config.ANTHROPIC_API_KEY:
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
            # Orphan תעודות שקילה already recorded on this project's
            # 'מעקב פנימי' sheet have no main-sheet row, so the line
            # above can't recover them - and write_records() rebuilds
            # that sheet from scratch on every write. Without this,
            # reopening a project and processing one more batch erased
            # every previously-reported orphan from the file.
            st.session_state.weighing_certificates = read_existing_weighing_certificates(
                st.session_state.output_path
            )

        errors = []

        def on_error(path, exc):
            errors.append((path.name, str(exc)))

        new_records, new_skipped, manual_records, new_weighing = [], [], [], []

        with st.status(f"מעבד {len(uploaded_files)} קבצים...", expanded=True) as status:
            if vision_uploads:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    tmp_paths = []
                    for uploaded in vision_uploads:
                        tmp_path = Path(tmp_dir) / uploaded.name
                        tmp_path.write_bytes(uploaded.getvalue())
                        tmp_paths.append(tmp_path)

                    def on_progress(index, total, path):
                        status.update(label=f"מעבד ({index}/{total}): {path.name}")

                    result = process_files(
                        tmp_paths,
                        on_progress=on_progress,
                        on_error=on_error,
                        selection=batch_selection,
                    )
                    new_records, new_skipped = result.records, result.skipped
                    new_weighing = list(result.weighing_certificates)

                # process_files() always returns "quantity" as a raw string
                # (like every other extracted field); the numeric column
                # downstream (the editable table, the Excel cell) needs a
                # clean float-or-"" invariant instead. Parse it the same way
                # write_records() will anyway, so the type is consistent
                # from here on. Manual-Excel records (below) deliberately
                # skip this - see app/excel_input.py's docstring on why an
                # unparseable value there must stay verbatim, not be blanked.
                for record in new_records:
                    raw_qty = record.get("quantity", "")
                    parsed_qty = parse_quantity(raw_qty)
                    if parsed_qty is not None:
                        record["quantity"] = parsed_qty
                    elif raw_qty:
                        # The model is instructed to return quantity as
                        # digits only (or blank) - a non-empty value that
                        # still doesn't parse means it ignored that
                        # instruction. Don't just blank it silently: note
                        # what it actually said, then blank it, so nothing
                        # is lost.
                        note = f"כמות לא תקינה: {raw_qty}"
                        record["notes"] = f"{record['notes']} | {note}" if record.get("notes") else note
                        record["quantity"] = ""
                    else:
                        record["quantity"] = ""

            if manual_uploads:
                status.update(label=f"קורא {len(manual_uploads)} קובצי Excel ידניים...")
                with tempfile.TemporaryDirectory() as tmp_dir:
                    for uploaded in manual_uploads:
                        tmp_path = Path(tmp_dir) / uploaded.name
                        tmp_path.write_bytes(uploaded.getvalue())
                        try:
                            manual_records.extend(read_manual_excel(tmp_path))
                        except Exception as exc:
                            on_error(tmp_path, exc)

            all_new_records = new_records + manual_records
            st.session_state.records.extend(all_new_records)
            st.session_state.skipped.extend(new_skipped)
            st.session_state.weighing_certificates.extend(new_weighing)
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
            combined_records = st.session_state.baseline_records + st.session_state.records
            # Manual-Excel rows vs. AI-extracted rows for the same physical
            # shipment (e.g. logged by hand today, the scanned certificate
            # uploaded next week) - flag for review, never auto-merge. See
            # app/duplicate_check.py; safe to call on every write, including
            # ones with no new manual/extracted rows at all (a no-op then).
            flag_potential_duplicates(combined_records)
            write_records(
                combined_records,
                st.session_state.output_path,
                weighing_certificates=st.session_state.weighing_certificates,
            )
            skip_note = f", {len(new_skipped)} דולגו כלא-תעודות" if new_skipped else ""
            manual_note = f", {len(manual_records)} מקובץ Excel ידני" if manual_records else ""
            status.update(
                label=f"הושלם - עובדו {len(all_new_records)} תעודות חדשות{skip_note}{manual_note}",
                state="complete",
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
        st.session_state.weighing_certificates = []
        st.session_state.batch_region = None
        st.session_state.batch_site = None
        st.session_state.batch_waste_type = None
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
        _render_stat_cards(0, 0)
else:
    with st.container(key="results_card"):
        st.subheader(":material/checklist: תוצאות", text_alignment="right")

        # Off by default - entry/exit-time and gross/tare-weight are
        # collected from every certificate (see extractor.py) but kept out
        # of the main columns unless a reviewer specifically asks for them
        # here. The same fields are always in the Excel file's own separate
        # "מעקב פנימי" sheet regardless of this toggle - see excel_writer.py's
        # _write_internal_tracking_sheet. vehicle_number/driver_name used to
        # be part of this toggle too, until 2026-09-03 when they were
        # promoted to always-visible main columns (see fields.EXCEL_COLUMNS).
        show_internal_tracking = st.checkbox(
            "הצג פרטים נוספים (מעקב פנימי)",
            help="מוסיף לטבלה עמודות שנאספות מהתעודה אך לא מוצגות כברירת מחדל: "
            "שעות כניסה/יציאה, משקל ברוטו/טרה. עמודות אלה תמיד נשמרות גם "
            'בגיליון "מעקב פנימי" הנפרד שבקובץ ה-Excel, גם כשהתיבה הזו לא '
            "מסומנת.",
        )

        header_by_key = dict(EXCEL_COLUMNS)
        keys = [key for key, _label in EXCEL_COLUMNS]
        # Re-insert confidence for on-screen display/editing only - see
        # _CONFIDENCE_COLUMN's own comment - in the same position it always
        # occupied (right before notes) so nothing downstream (display_keys'
        # status-symbol insertion, the data_editor columns, the merge-back
        # loop) needs to know this changed.
        header_by_key[_CONFIDENCE_COLUMN[0]] = _CONFIDENCE_COLUMN[1]
        notes_pos = keys.index("notes")
        keys = keys[:notes_pos] + [_CONFIDENCE_COLUMN[0]] + keys[notes_pos:]
        if show_internal_tracking:
            header_by_key.update(INTERNAL_TRACKING_FIELDS)
            keys = keys + [key for key, _label in INTERNAL_TRACKING_FIELDS]
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

        def _display_quantity(value):
            # "כמות מדווחת" is rendered as free text, not st.column_config's
            # NumberColumn - deliberately: a manual-Excel row's quantity can
            # be a non-numeric human annotation like "כ-1.8" (see
            # app/excel_input.py's docstring), and NumberColumn would either
            # reject or silently blank that. A plain float still gets the
            # same fixed 2-decimal look NumberColumn's format="%,.2f" gave
            # it before; anything else (a raw string, or "") passes through
            # untouched.
            return f"{value:,.2f}" if isinstance(value, (int, float)) else (value or "")

        # __idx__ carries each row's position in st.session_state.records
        # through the editor and back - hidden from view via column_config
        # below, but still round-trips in the returned data (per
        # st.column_config's own "hides from the UI, data still there"
        # behavior). Matching edits back by this explicit id, rather than by
        # row position, means the merge below stays correct even if a future
        # Streamlit version adds interactive sorting to st.data_editor.
        rows = [
            {
                **{k: (_display_quantity(r.get(k)) if k == "quantity" else r.get(k, "")) for k in keys},
                _STATUS_KEY: _STATUS_SYMBOLS.get(r.get("confidence"), ""),
                "__idx__": i,
            }
            for i, r in zip(filtered_indices, visible_records)
        ]
        df = pd.DataFrame(rows, columns=display_keys + ["__idx__"])
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
                "כמות מדווחת": st.column_config.TextColumn(),
                "רמת ביטחון": st.column_config.SelectboxColumn(options=CONFIDENCE_LEVELS),
                "סוג הפסולת": st.column_config.SelectboxColumn(
                    options=_selectbox_options(
                        WASTE_TYPES + [UNCLASSIFIED_WASTE_TYPE], (r.get("waste_type") for r in visible_records)
                    )
                ),
                "מרחב / יחידה ראשית": st.column_config.SelectboxColumn(
                    options=_selectbox_options(REGIONS, (r.get("region") for r in visible_records))
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
                # "כמות מדווחת" is a free-text column now (see the
                # TextColumn/_display_quantity note above it), specifically
                # so a manual-Excel row's non-numeric quantity (e.g.
                # "כ-1.8") survives being displayed and re-edited without
                # being coerced to a number or blanked. parse_quantity()
                # reuses the exact same numeric-or-raw-string fallback
                # excel_writer.write_records() applies at write time, so a
                # normal numeric edit still round-trips as a float.
                text_value = "" if value is None else str(value).strip()
                parsed = parse_quantity(text_value) if text_value else None
                target[key] = parsed if parsed is not None else text_value
            else:
                target[key] = "" if value is None else value

    st.session_state.records[:] = sort_by_confidence(st.session_state.records)
    # baseline_records + records, same as after processing above - an edit
    # here must not overwrite the project's file with only what's on screen.
    all_records = st.session_state.baseline_records + st.session_state.records
    # Same duplicate check as after processing above - an edit here (e.g.
    # correcting a site/date/quantity by hand) can newly create or resolve
    # an apparent match, so it must be re-evaluated on every write, not just
    # right after a batch is processed.
    flag_potential_duplicates(all_records)
    write_records(
        all_records,
        st.session_state.output_path,
        weighing_certificates=st.session_state.weighing_certificates,
    )

    with summary_slot:
        _render_stat_cards(
            len(all_records),
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
            # confidence re-inserted here too, right before notes, for the
            # same reason as the results table above - see
            # _CONFIDENCE_COLUMN's comment.
            notes_pos = next(i for i, (key, _label) in enumerate(EXCEL_COLUMNS) if key == "notes")
            fields_to_show = (
                EXCEL_COLUMNS[:notes_pos]
                + [_CONFIDENCE_COLUMN]
                + EXCEL_COLUMNS[notes_pos:]
                + (INTERNAL_TRACKING_FIELDS if show_internal_tracking else [])
            )
            for key, label in fields_to_show:
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
