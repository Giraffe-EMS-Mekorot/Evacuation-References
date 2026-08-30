# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An agent that reads Israeli waste-disposal certificates (תעודות פינוי פסולת) —
PDFs and photos, some computer-generated, some handwritten — and extracts
structured data into one consolidated Excel file, using Claude's vision
capability to read the certificates directly (no OCR library, no PDF-to-image
conversion: PDFs are sent to the Messages API as native `document` content
blocks, images as `image` blocks).

## Commands

```bash
pip install -r requirements.txt        # setup (or use the existing .venv)

python main.py                         # process every file in input/
python main.py --file <name>           # process a single file in input/ (useful while iterating)
```

There is no lint/test tooling configured yet. `.venv/Scripts/python -c "..."` is
how ad-hoc smoke checks have been run so far (e.g. importing `app.*` and
calling `write_records()` with fabricated dicts) since there's no sample
certificate committed to the repo to run the real pipeline against —
`input/` itself has real (uncommitted, gitignored) certificates from actual
use, though, so a pure-function/Excel-round-trip smoke test can be paired
with one real, cheap (single-page) live extraction call against one of
those when validating extraction-quality-sensitive changes (not just
logic) — see the gross/tare/cert_role verification done this way on
2026-08-30, described below.

`streamlit.testing.v1.AppTest` (ships with the installed Streamlit, no
extra install) can drive `streamlit_app.py` headlessly — `AppTest.from_file(...)`,
`.run()`, `at.text_input[0].set_value(...)`/`at.button[0].click()` to get
past the password gate (reads the real `.streamlit/secrets.toml` locally),
then seed `at.session_state["records"]` directly (bypassing
`file_uploader`, which AppTest can't script) to exercise the
results-table/toggle/round-trip rendering path without a live API call.
`at.exception` is empty on success; widgets are queried via `at.checkbox`,
`at.subheader`, `at.metric`, etc., each exposing a plain `.value`.

To verify a `סיכום`-sheet formula actually computes right (not just that it's
syntactically well-formed), `pip install formulas` and evaluate the generated
file:
```python
import formulas
sol = formulas.ExcelModel().loads("output/ריכוז_תעודות.xlsx").finish().calculate()
```
`sol` keys look like `'[file.xlsx]סיכום'!H22`; compare against hand-computed
expected values. See the `סיכום`-sheet note below for a real bug this caught.

Requires `ANTHROPIC_API_KEY` in `.env` at the repo root (already present).
Optional `CLAUDE_MODEL` in `.env` overrides the default vision model
(`claude-sonnet-5`, set in `app/config.py`).

## Architecture

**`app/fields.py` holds several related but distinct lists — don't conflate them.**
- `FIELD_DEFS`: `(field_name, description_for_model, json_type)` tuples for
  what Claude actually extracts from the certificate. Drives the tool-use
  JSON schema in `extractor.py` and the record dict keys used through the
  pipeline. Add/remove a *model-extracted* field here.
- `EXCEL_COLUMNS`: `(record_key, hebrew_header)` tuples defining the exact
  main-sheet column order and headers, independent of extraction order.
  Includes keys that aren't in `FIELD_DEFS` at all —
  `certificate_or_reference` (derived in code, see below) and `source_file`
  (filled in by the pipeline). Add/reorder a *main-sheet* column here — this
  list is what `excel_writer.py` iterates for the main sheet, so column
  order changes never require touching the extraction code. As of
  2026-08-30 this is a deliberately short, priority-ordered "core" list
  (date, certificate_or_reference, supplier_or_carrier, site, waste_type,
  quantity, unit, region, confidence, notes, source_file) — see
  `INTERNAL_TRACKING_FIELDS` below for fields that are extracted but kept
  off this list on purpose. `year`/`month` used to be separate
  `EXCEL_COLUMNS` entries derived from `date`; they no longer are (the raw
  `date` field is shown directly instead), but `derive_year_month()` still
  computes them internally for `_flag_unreasonable_date()` and
  `_enforce_core_field_confidence()` in `extractor.py` — don't be surprised
  to find `record["year"]`/`record["month"]` set on a record that never
  shows them in the spreadsheet.
- `INTERNAL_TRACKING_FIELDS`: `(record_key, hebrew_header)` pairs for fields
  that ARE in `FIELD_DEFS` (vehicle_number, driver_name, entry_time,
  exit_time, gross_weight, tare_weight) but are deliberately excluded from
  `EXCEL_COLUMNS` — always collected, never shown on the main sheet by
  default. Surfaced two ways: streamlit_app.py's "הצג פרטים נוספים" checkbox
  appends these to the on-screen/editable table, and excel_writer.py's
  `_write_internal_tracking_sheet()` always writes them to a second
  "מעקב פנימי" sheet regardless of that checkbox. `INTERNAL_TRACKING_SHEET_COLUMNS`
  is the same 6 fields with `source_file`/`certificate_or_reference`/`site`
  prepended, since that Excel sheet is physically separate from the main one
  and needs identifying columns to be self-sufficient.

The closed-list enums (`WASTE_TYPES`, `REGIONS`, `CONFIDENCE_LEVELS`,
`HEBREW_MONTHS`) live in the same file and are woven into the field
descriptions the model receives.

**`app/derive.py` computes `year`/`month` from the extracted `date` in code,
not via the model.** `derive_year_month()` parses `DD/MM/YYYY` (and tolerates
`DD/MM/YY` — the model doesn't always follow the requested format, e.g. it
returned `06/07/25` once during testing; a stricter 4-digit-only regex
silently dropped a real date). Returns `("", "")` on anything unparseable —
same "blank over guess" policy as the rest of the pipeline, but enforced
deterministically here since date arithmetic doesn't need an LLM.

**Extraction (`app/extractor.py`) uses forced tool-use, not free-form JSON
parsing.** `extract_certificate()` calls `messages.create()` with
`tool_choice={"type": "tool", "name": "record_certificate_data"}`, so Claude
must return arguments matching the generated schema — no markdown-fence/JSON
parsing needed. Fields the model can't read confidently come back as empty
strings by instruction (never guessed — see the caveat below though).
`date` is extracted only to feed `derive_year_month()`; it is not itself an
output column.

After the tool call returns, three code-side passes run in sequence, each a
safety net against the model's self-reported `confidence` being wrong in a
specific way:
1. `_flag_out_of_list_values()` — if `waste_type`/`region` isn't actually one
   of the closed values, or confidence is missing, force-downgrades to
   "נמוכה" and appends a note. Guards against the model drifting off the enum.
2. `record["year"], record["month"] = derive_year_month(record["date"])`.
3. `_enforce_core_field_confidence()` — if any core field
   (`_CORE_FIELD_LABELS`: waste_type/quantity/unit/site/region/year+month) is
   blank, forces confidence to at least "בינונית" and appends a note listing
   what's missing. Guards against the model reporting "גבוהה" purely on
   *legibility* of what it did read, while a whole core field is empty —
   e.g. it can read a certificate perfectly clearly and still have no מרחב on
   it. Added 2026-08-23 after that exact case slipped through unflagged.
   Only escalates, never downgrades below "נמוכה"; if a row is already at
   "נמוכה"/"בינונית" it's left untouched with no added note (avoids
   re-flagging the same row for two overlapping reasons).

**Excel output (`app/excel_writer.py`)** writes two sheets: `ריכוז תעודות`
(the data) and `סיכום` (a live-formula summary, see below).

The main sheet is written in `EXCEL_COLUMNS` order: RTL layout, bold 13pt
white-on-dark-blue header (taller row height, centered), thin black borders
on every cell, autofilter over the full used range, and column widths
auto-sized from content length (openpyxl has no true text-measuring autofit,
so this is a character-count heuristic, clamped to 10–40). Data rows are
11pt, right-aligned except `quantity`/`year` which are centered (see
`_CENTERED_KEYS` if adding another numeric column). `freeze_panes = "B2"`
freezes both the header row and logical column A — which is `region`/מרחב,
i.e. the *rightmost* column in this RTL sheet, not the leftmost; don't "fix"
this to `"A2"` thinking it's a typo.

`quantity` is written as a real numeric cell (`_parse_quantity()` strips
thousands-separator commas and casts to `float`; falls back to the raw
string only if unparseable), not text — this is required, not cosmetic: the
`סיכום` sheet's SUM/SUMIFS formulas silently return 0 against a text-typed
column, since openpyxl-written strings aren't the same as Excel's numeric
type. If a future column ever needs summing on the summary sheet, it needs
the same numeric-cast treatment when written.

Row background layers two independent things, in priority order:
1. **Confidence fill** (red = נמוכה, yellow = בינונית, none = גבוהה) — this
   is the main UX mechanism of the whole tool, not a cosmetic extra: it's
   what lets a reviewer find certificates that need a manual look.
2. **Banded-row fill** (light gray on alternating rows) — purely a
   readability aid, applied only when there's no confidence fill for that
   row. Never let banding override or blend with confidence color.

The `source_file` column (last in `EXCEL_COLUMNS`, so visually leftmost) is
rendered as a real hyperlink — blue/underlined — to
`(config.INPUT_DIR / filename).resolve().as_uri()`, so clicking it in Excel
opens the original certificate. This assumes the file still lives in
`input/` under that exact name; it isn't a durable reference if files get
moved/renamed after a run.

**The `סיכום` sheet (`_write_summary_sheet()`) is built entirely from live
Excel formulas** (`COUNTA`/`COUNTIF`/`SUMIF`/`SUMIFS`/`SUM`), never
precomputed Python values, per an explicit user requirement: editing a row on
the main sheet by hand must update the summary without rerunning the
pipeline. Every range is a fixed, generous `$col$2:$col$10000` bound
(`_MAX_DATA_ROW`), not the true row count — deliberately, so rows added later
by hand are still covered. It has four sections: total certificate count,
confidence-level breakdown, a waste-type × unit pivot table (with a
"catch-all" row/column absorbing anything that doesn't match a known
waste-type or unit, and row/column totals that cross-check against
`SUM(qty_range)`), and a per-region breakdown with a similar catch-all row.

**Before trusting a new formula here, verify it — don't just eyeball the
formula text.** openpyxl never evaluates formulas, so `write_records()`
alone can't tell you a formula is *wrong*, only that it's syntactically
present. This bit us for real during development (2026-08-23): a first
version tried to split the region catch-all into two rows — "blank region"
via `COUNTIF(range,"")` and "non-blank but unrecognized" via
`COUNTIF(range,"<>")`. Both were wrong, for two different reasons:
- `COUNTIF(range,"")` over a padded `$A$2:$A$10000` range counts every
  never-written padding row as "blank" too, not just real records with an
  empty region — this is genuine Excel behavior (confirmed against real
  semantics), not a tooling bug, but it made the number meaningless here.
- `COUNTIF(range,"<>")` was empirically found to misbehave in the `formulas`
  Python package used to verify this (returned the *same* count as `""`,
  i.e. exactly backwards) - a library limitation, not necessarily how real
  Excel behaves, but since there's no way to cross-check against real Excel
  in this environment, formulas relying on `"<>"`/`""` blank-testing over a
  padded range should be treated as unverifiable and avoided.

The fix that's actually in the code now: the catch-all rows (region section,
and the waste-type "לא מסווג" row) use *only* `COUNTA`, `SUM`, and
exact-value `COUNTIF`/`SUMIF`/`SUMIFS` (matching a literal string like
"צפון" or "ק\"ג") — arithmetic subtraction from a trusted total, never a
blank/non-blank test. E.g. "ללא מרחב מזוהה" count is
`COUNTA(source_file_range) - SUM(the 4 known-region counts)`, not a direct
blank-region count. This pattern was verified end-to-end (not just by
inspection) using the `formulas` PyPI package
(`ExcelModel().loads(path).finish().calculate()`) against hand-computed
expected numbers on both a synthetic 7-record dataset and the real
certificates - see the "verify formulas" approach below before adding new
aggregate formulas to this sheet.

**`main.py`** is the batch loop: iterates `input/`, calls `extract_certificate()`
per file inside a try/except so one bad/corrupt certificate doesn't abort the
run — it's recorded instead as a low-confidence row (via `_empty_record()`,
which also fills `year`/`month` as `""`) with the exception text in `notes`,
then processing continues. `--file` restricts to one filename for faster
iteration when working on the pipeline itself.

## Non-obvious constraints from the spec

- `quantity`/`date`/etc. must never be guessed when illegible — this is
  enforced via the field descriptions sent to the model, not by code. If
  Claude ever fabricates values here, the fix is prompt wording in
  `fields.py`/`extractor.py`, not new validation code.
- `region` (מרחב) has no hardcoded site→region lookup table; the model is
  asked to infer it from context on the certificate itself and leave it blank
  otherwise.
- `unit` is intentionally free text ("as it appears on the certificate"), not
  a closed enum like `waste_type`.
- `vehicle_number` and `driver_or_carrier` were extracted fields originally
  but were removed on user feedback (2026-08-23) as "not needed". Reinstated
  2026-08-30 as `vehicle_number`/`driver_name` under `INTERNAL_TRACKING_FIELDS`
  (see above) — collected again, but kept off the main sheet this time
  instead of being a full main-sheet column like before.

## Document-pair detection and cross-check (2026-08-30)

Some suppliers (not all, and not a fixed/known list — see below) issue two
adjacent-page documents per removal instead of one: a computerized
"תעודת שקילה/משלוח" (weighing/delivery slip, with ברוטו/טרה/נטו and its own
`certificate_number`) and a usually-handwritten "אישור ביצוע עבודה/הטמנה"
(work/disposal completion approval, with a `reference_number` pointing back
at the weighing slip's certificate number, a site, and a signature).

**Detection is per-page and model-driven, not a Python heuristic.** Each
page is still sent to the model independently (see the multi-page-PDF note
below) — there's no code path that lets one page's extraction see another's.
`fields.py`'s `cert_role` field asks the model to classify *this page's own*
structure/headers as `CERT_ROLE_WEIGHING`, `CERT_ROLE_COMPLETION`, or blank
(the common case — most certificates aren't part of this pattern at all),
**regardless of whether it can confirm a matching neighboring page exists**
— it's judging this page's own shape only; the actual pairing is a separate
code-side step. This distinction had to be spelled out explicitly in both
the field description and the system prompt after an early test (real
certificate, `input/1226_פסולת בניין 07.2026_1226.pdf`) came back with
`cert_role` blank on an obviously-weighing-shaped page — the model was
reading the field as "confirm this is actually paired," which it structurally
can't do from one page, and defaulting to blank. Once reworded to "classify
this page's own shape; the pairing check happens elsewhere," the same page
correctly returned `CERT_ROLE_WEIGHING`. Tagging cert_role liberally based on
structure alone is intentionally low-risk: a page tagged `CERT_ROLE_WEIGHING`
with no complementary neighbor just never triggers a cross-check (see next
paragraph) — under-tagging (silently skipping a real pair) is the failure
mode actually worth avoiding, not over-tagging.

**The cross-check itself is code, not the model** —
`extractor._detect_and_cross_check_pairs()`, called once per PDF right after
all its pages are extracted (inside `extract_certificate_pages()`, not
`_extract_page()`, since it needs the *list* of page records, not just one).
It scans strictly-adjacent record pairs (`records[i]`/`records[i+1]` — matches
the spec's "שני עמודים סמוכים"; not applied across separate uploaded files,
and not applied to non-adjacent pages even within one PDF) for one
`CERT_ROLE_WEIGHING` + one `CERT_ROLE_COMPLETION`, in either order, and runs
`_cross_check_document_pair()` on any match: reference-number equality
(`certificate_number` vs `reference_number`, normalized via
`_normalize_reference()` — numeric ids compare with leading zeros stripped,
non-numeric ids compare case-folded), quantity equality (within
`_PAIR_QUANTITY_TOLERANCE`, currently 1.0), and date equality. Any mismatch
downgrades **both** records in the pair to `נמוכה` with one note listing
exactly what didn't match (e.g. `"אסמכתא לא תואמת: תעודת שקילה=15077, אישור
ביצוע=15078"`); a clean pass adds nothing (per spec, "high confidence, no
note needed" — not an explicit confirmation note). A `CERT_ROLE_COMPLETION`
page with no matching neighbor (or vice versa) is left completely alone —
not a failure, just no cross-check, per spec.

Unrelated to the pairing above: `_flag_gross_tare_mismatch()` checks
ברוטו − טרה = נטו (± `_GROSS_TARE_TOLERANCE`, 1.0) on **any single record**
that has both `gross_weight` and `tare_weight`, whether or not it's part of
a detected pair — this is the per-document arithmetic sanity check from the
spec, distinct from the pair-level quantity cross-check above. Only runs
when `unit` is blank or `ק"ג` (the tolerance is a fixed ק"ג margin, meaningless
against, say, טון). Verified against real certificate data
(`input/1226_פסולת בניין 07.2026_1226.pdf`'s actual ברוטו/טרה/נטו, e.g.
33670/12440/21230) to confirm it doesn't false-positive on genuinely
consistent values — see the smoke-test approach below.

`container_count`/`container_type` are extracted fields but never their own
column anywhere (main sheet, tracking sheet, or tracking toggle) — they're
folded into `notes` by `_append_container_note()` right after extraction,
per spec ("שדות שמתווספים אוטומטית להערות").

`certificate_or_reference` (the main sheet's "מספר תעודה/אסמכתא" column) is
computed in `_extract_page()`, not asked of the model directly: `certificate_number`
if present, else `reference_number` as a fallback — handles a certificate
that labels its own id "אסמכתא" instead of "מספר תעודה," while keeping the
two fields distinct for the cross-check above (which needs to know which
role each number is playing, not just "some id was present").

**`excel_writer.read_existing_records()` matches columns by header-row
TEXT now, not by fixed position** — a direct consequence of this
core-column reorder. Before 2026-08-30 it zipped `_KEYS` straight against
row values by position; a file written by an older pipeline version with a
different column count/order would have silently misaligned every value.
It now builds a `{header_label: record_key}` map from the current
`EXCEL_COLUMNS`, falling back to `_LEGACY_HEADER_ALIASES` for the specific
headers this reorder renamed (`"כמות"` → quantity, `"אתר/יחידה"` → site,
`"סוג אסמכתא מצורפת"` → supplier_or_carrier) or dropped (`"שנה"`/`"חודש"` →
`None`, meaning "no replacement, skip this column"). A column in the old
file that matches neither is silently skipped (not an error) — this is what
lets a project file from before this change keep accumulating correctly
after it, instead of a reused project name corrupting its own history.

## Known limitation: "don't guess" is prompt-only, and imperfect

Tested against 4 real certificates (2026-08-23): the model sometimes fills
`waste_type`/`region` with an inferred value and *admits in `הערות` that it's
a guess*, rather than leaving the field empty as instructed — e.g. classifying
a battery line item as "פסולת בנין" because nothing else in the closed list
fit, while noting in the same breath that the match isn't real.
`_flag_out_of_list_values()` only catches values that are literally outside
the closed list, so a wrong-but-valid guess like that slips past it.

Pushing the system prompt/field descriptions harder (tried once) has limited,
inconsistent payoff: it suppressed one bad guess but also made the model
blank out a previously-correct, unambiguous inference (Eilat → דרום) that
didn't need suppressing. Prompt wording alone hasn't fully solved this class
of problem — don't assume another wording tweak will either.

What *does* reliably work as the safety net: guessed/uncertain fields
consistently come with a downgraded `רמת_ביטחון` (נמוכה/בינונית) and an
explanatory note, so the Excel red/yellow highlighting still surfaces these
rows for manual review even when the cell itself isn't blank. If tighter
enforcement is ever needed, the next lever to try is probably a second
model call that specifically audits closed-list fields against "is this
literally written on the document, or inferred?" rather than further prompt
tuning on the extraction call itself.

`הערות` is deliberately kept to ~4-5 words (per the field description and
system prompt in `extractor.py`/`fields.py`, 2026-08-23) — e.g. "כתב יד לא
קריא" rather than a full sentence explaining what was inferred and why. This
was a user readability request, not a data-quality fix, so it doesn't change
anything about the guessing behavior above - a short note can still describe
a guessed value, it just won't spell out the reasoning anymore.
