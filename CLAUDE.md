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
`"סוג אסמכתא מצורפת"` → supplier_or_carrier, and `"ספק/מוביל"` →
supplier_or_carrier too — see the 2026-08-30 follow-up below, same day, that
renamed the *display* header again to "אתר קולט" without touching the
`supplier_or_carrier` key) or reconstructed (`"שנה"`/`"חודש"` → the sentinel
`_LEGACY_YEAR_KEY`/`_LEGACY_MONTH_KEY`, post-processed into a synthesized
`MM/YYYY` `date` — no day, since none exists on those old rows; see the
follow-up below for why this isn't a straight drop like it originally was).
A column in the old file that matches neither is silently skipped (not an
error) — this is what lets a project file from before this change keep
accumulating correctly after it, instead of a reused project name
corrupting its own history.

## Follow-up fixes, same day (2026-08-30): RTL-reversed names, a
## misdiagnosed date-flagging "bug", legacy date migration, one more rename

Four asks that came in immediately after the core-column reorder above, from
the user reviewing its first real output. Two turned out to be real; two did
not (traced down, not just dismissed) — worth reading past the fix list to
the investigation notes, so a similar report next time gets checked instead
of re-"fixed":

**1. RTL-reversed supplier name.** The user reported a name reading backwards
in `output/ריכוז_תעודות.xlsx`. Inspected that file's actual cell codepoints
directly (`repr()` + `hex(ord(c))`, not just eyeballing rendered text) —
**the stored data was correct**, spelled forward, no bidi control
characters. The garbled appearance was almost certainly a copy/paste or
terminal-rendering artifact on the way from Excel into the chat, not a
pipeline bug — nothing in that specific file needed fixing or removing, and
none was (a backup was still taken before this session's other Excel-facing
changes, per the ask, out of caution). Real risk taken seriously anyway,
since a vision model misreading RTL Hebrew glyphs in left-to-right pixel
order is a genuine, known failure mode even though this particular report
wasn't an instance of it:
- `normalize.looks_reversed_hebrew()` (new): a dictionary-free heuristic —
  Hebrew has 5 letters with a mandatory word-final ("sofit") form
  (ך ם ן ף ץ); a final form appearing mid-word, or a regular form that has
  one (כ מ נ פ צ) ending a word, is a strong signal that word's characters
  came out reversed. Only fires on a whitespace-split token that's pure
  Hebrew letters (a mixed token like "מ.עבד" or "VERIDIS" is skipped
  entirely, in either direction) — and only catches a reversal that
  actually involves one of those 5 letters, which not every word does (e.g.
  "הובלות" reversed trips nothing on its own; see the module docstring's
  worked example before assuming this heuristic is a general-purpose
  garbled-text detector — it isn't).
- `NameNormalizer.normalize()`'s existing "no match → new name" branch now
  tries the value's full-string reversal (`value[::-1]`) against the known-
  names list *before* concluding "genuinely new" — a match there rewrites
  to the known canonical spelling with its own distinct note (`תוקן מטקסט
  שנקרא הפוך`), same mechanism as the existing OCR/handwriting-variant
  path, just naming the different root cause. This one has real reference
  data to correct against, unlike the next item.
- `extractor._flag_possibly_reversed_text()` (new, called on
  site/supplier_or_carrier/driver_name in `_extract_page`): flags only —
  escalates confidence like `_enforce_core_field_confidence`'s pattern,
  never auto-corrects. Extraction time has no "known names" list to check a
  reversal against (that only exists in `pipeline.py`'s per-batch
  `NameNormalizer`), so guessing which direction is "right" here would
  violate this pipeline's whole "never guess, flag instead" policy.

**2. "`_flag_unreasonable_date` wrongly flags valid dates like 27/07/26"**
— investigated, not just patched. Traced the actual real system clock
(`date.today()`, confirmed `2026-08-30` — matches the environment, not a
sandbox-clock mismatch), then re-derived `derive_year_month("27/07/26")` by
hand against it: `(2026, 7) > (2026, 8)` is `False`, so `is_future` is
correctly `False` — **no bug in `derive.py` or `_flag_unreasonable_date`**.
Confirmed against real run data too: none of a 10-record real batch ever hit
`confidence == "נמוכה"` from this path, and none carried this function's own
note text (`"תאריך לא סביר (עתידי): ..."`). What the user actually saw was
the *model* volunteering "תאריך עתידי חריג" in `notes` on its own initiative
— similar-sounding text, different source, and wrong: the model has no
reliable way to know the real "today" (it was very likely comparing against
its own training-cutoff-era sense of "recent," or the certificate's own
reprint/export timestamp, neither of which is "today"). **The actual fix is
a prompt change, not a `derive.py`/`extractor.py` logic change**: both the
system prompt and the `notes` field's own `FIELD_DEFS` description now
explicitly tell the model not to comment on date plausibility at all — this
system already checks it deterministically and correctly after extraction,
so the model doing it too just risked exactly this kind of confusing,
occasionally-wrong duplicate signal.
Also checked, since the user asked directly: **no "date-format-memory-per-
supplier" mechanism exists anywhere in this codebase** (`normalize.py` only
normalizes `site`/`supplier_or_carrier` name spellings — nothing about
dates). If this comes up again, the user may be thinking of
`derive_year_month()` in `derive.py`, which does read a full date including
the day — but only to derive year+month from it, discarding the day itself;
it was never a "memory" of anything.

**3. Legacy year/month → partial date, on top of the reorder's own
backward-compat.** The reorder above already made `read_existing_records()`
header-name-based; this follow-up specifically taught it to *reconstruct*
(not just skip) a `date` for a pre-reorder row that only ever had separate
"שנה"/"חודש" columns — synthesized as `MM/YYYY` (e.g. `"06/2026"`), never
inventing a day that was never on the certificate. See
`_LEGACY_YEAR_KEY`/`_LEGACY_MONTH_KEY` in `excel_writer.py`.

**4. Header rename: "ספק/מוביל" → "אתר קולט".** Display-only, same day as
the reorder that introduced "ספק/מוביל" in the first place — the
`supplier_or_carrier` record key and its `FIELD_DEFS` description are
unchanged; only the `EXCEL_COLUMNS` label changed. `"ספק/מוביל"` was added
to `_LEGACY_HEADER_ALIASES` alongside the older `"סוג אסמכתא מצורפת"`, so a
file written in the few hours between these two renames still reads back
correctly too.

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

## Four real-world-driven improvements (2026-09-01)

**1. Invoice (חשבונית) filtering.** Extended the existing, already-model-driven
`document_type` classification (see DOCUMENT_TYPE_OTHER above) to explicitly
recognize an invoice pattern - a title containing "חשבונית", multiple line-item
rows, price/מע"מ columns, a financial summary (the real example: "ק.מ.מ. מפעלי
מחזור") - and classify it as `DOCUMENT_TYPE_OTHER`, same as an internal
billing/credit note. Pure prompt-engineering (fields.py's `document_type`
field description + extractor.py's `_extract_page` skip-path was already
generic), no new code path - verified against a real certificate (not an
invoice) that it still correctly returns `DOCUMENT_TYPE_CERTIFICATE` and
doesn't over-trigger.

**2. שטר מטען (bill of lading) with an explicit "0" weight.** A document like
a טיקטראק שטר מטען that accompanies a shipment *before* weighing has its own
weight/volume field literally showing "0" - not blank, and not a real zero
quantity. A new `cert_role` value, `CERT_ROLE_BILL_OF_LADING_ZERO`
(fields.py), is a third, unrelated option alongside the existing
WEIGHING/COMPLETION pair pattern (see the 2026-08-30 section above) - model-
classified from the page's own structure, same as those two. When set:
- The model is instructed (system prompt + `quantity`'s field description)
  to ignore the literal "0" and instead look for a textual estimate
  elsewhere on the page (e.g. a "פרטים" field like "פינוי מכולה 12 מ\"ק") and
  extract a real number+unit from it into `quantity`/`unit`, exactly as if
  that number had been in the weighing table itself.
- `extractor._flag_bill_of_lading_estimate()` always escalates such a record
  to at least "בינונית" with a "הערכה משטר מטען - טרם נשקל בפועל" note - the
  model reading the page perfectly doesn't make an estimate a real weighing.
- `pipeline._cross_check_bill_of_lading_estimates()` runs once per
  `process_files()` batch (across every file in the batch, NOT just one
  PDF's adjacent pages - contrast with `extractor._detect_and_cross_check_pairs`,
  which is same-PDF/adjacent-page only) looking for a real weighing record
  (positive gross/tare, or just a positive quantity - see
  `_is_real_weighing_candidate`) matching the estimate on vehicle_number +
  site + date within 3 days. Deliberately does NOT require a matching
  certificate/reference number - the spec's own real scenario is two
  different systems (the carrier's שטר מטען vs. the receiving site's own
  weighing slip) assigning their own unrelated ids to the same shipment.
  On a match, the estimate's quantity/unit are overwritten with the real
  weighing's, noted, and the real record is left untouched.
- `derive.parse_date()` (new, additive - `derive_year_month()` itself is
  untouched to avoid touching this project's most incident-prone piece of
  date logic, see the 2026-08-30 follow-up above) parses a date string into
  a real `date` object for this day-level proximity comparison, reused by
  `duplicate_check.py` below too.

**3. Manual Excel upload (`app/excel_input.py`, new module).**
streamlit_app.py's existing upload area now also accepts `.xlsx` (detected
by extension, `_MANUAL_EXCEL_EXTENSIONS`) alongside PDF/image files -
`read_manual_excel()` reads it directly via openpyxl, no Claude call at all.
Column matching is by header TEXT (`_HEADER_TO_KEY`), not position, so an
unrecognized "helper" column (dropdown source lists like "אגפים"/"סוגי
פסולת" sitting off to the side of the real table) is silently never read -
no separate skip-list needed. Deliberately bypasses `app.pipeline.process_files`
entirely (NameNormalizer, quantity-history outlier check, gross/tare check,
closed-list enforcement, date-plausibility check all exist to catch a
*vision model's* mistakes; a manually-typed row was never extracted by one)
- streamlit_app.py appends these records straight into
`st.session_state.records`, alongside whatever `process_files()` returned
for any PDF/images in the same batch. Confidence defaults to "גבוהה" (a
human already reviewed this data) except a missing year, which is left
blank (never guessed from the sheet name) with a "שנה חסרה במקור" note and
"נמוכה" confidence. Quantity is kept EXACTLY as the cell holds it - a human
annotation like "כ-1.8" is never rejected/cleaned the way an unparseable
AI-extracted value is (see streamlit_app.py's own quantity-reparse block,
which now explicitly only runs on `process_files()`'s output, not manual
records) - `excel_writer.write_records()`'s existing parse-or-fall-back-to-
raw-string behavior already renders this correctly with no extra code.

This forced a related fix in streamlit_app.py's results table: "כמות (נטו)"
was `st.column_config.NumberColumn`, which can't hold a string like "כ-1.8"
- it's now `TextColumn` (`_display_quantity()` still right-formats a real
float to 2 decimals for the common case), and the edit-merge-back logic
uses `parse_quantity()` with a raw-string fallback instead of an
unconditional `float(value)`, so a normal numeric edit still round-trips as
a float while non-numeric manual text survives being displayed and re-edited.

**4. Duplicate detection between a manual Excel row and a later PDF scan
(`app/duplicate_check.py`, new module).** The same shipment can get reported
twice - once by hand ahead of time, once for real when the scanned
certificate arrives. `flag_potential_duplicates()` compares every
(manual-Excel, AI-extracted) pair - origin is inferred from `source_file`'s
own extension (`_is_manual_excel_record`), not a separate tracked field, so
it works identically on records built this session and on rows read back
from a previously-saved project file - for a close/equal date (3 days),
identical site + waste_type, and a close/equal quantity (5%). A match never
merges or deletes anything (too dangerous to automate) - it appends a
"⚠ כפילות אפשרית" note to BOTH rows, referencing the OTHER row by source
file name (+ date), deliberately NOT a spreadsheet row number: row position
shifts on every confidence re-sort (`sort_by_confidence`, which
`write_records()` always applies) or later batch, so a numeric reference
would go stale almost immediately - a source-file reference stays valid and
is idempotent to boot (calling this again on unchanged data doesn't
re-append the note). Wired into streamlit_app.py right before both of its
`write_records()` calls (after processing a new batch, and after an
edit-merge-back), since either can newly create or resolve an apparent
match. CLI (main.py) is untouched - it has no Excel-upload code path at all,
so there's nothing for this to cross-reference there.

**Testing note:** all four were verified with fabricated-record pure-function
tests (the bill-of-lading flag/cross-check, the manual-Excel column-mapping/
edge-cases, the duplicate-flagging match/no-match/idempotency cases) plus a
full `write_records()`/`read_existing_records()` Excel round-trip with mixed
numeric/text quantities, plus a `streamlit.testing.v1.AppTest` run seeding
`session_state.records` with a mix of normal/manual/flagged-duplicate rows to
confirm the results table renders without exception. No real חשבונית or שטר
מטען certificate exists in this project's `input/` to run a live extraction
call against (unlike the gross/tare/cert_role work on 2026-08-30, which had
one) - #1 and #2's PROMPT changes are therefore unverified against a real
matching document; what WAS run live is a single real page from an existing
certificate (`מחזור אלקטרוניקה אחיסמך 2025.pdf`, page 1) to confirm the
system prompt/field-description additions don't regress ordinary extraction
- it still returned `DOCUMENT_TYPE_CERTIFICATE` with `cert_role` correctly
blank.

## Weighing-certificate pairing (2026-09-15)

A second, independent document-pairing mechanism, added alongside - not
replacing - the 2026-08-30 adjacent-page ברוטו/טרה/נטו pair check. Read both
sections before touching either; they are genuinely different features that
happen to share the word "pair", and a document can legitimately participate
in both at once.

**The pattern:** a computerized תעודת משלוח arrives together with a separate,
usually-handwritten תעודת שקילה for the same removal. The delivery note
carries all the real data; the weighing certificate carries essentially
nothing except a printed running number in its header, which the delivery
note echoes in its own "אסמכתא" field.

**Classification is structural, not "is there handwriting".** A fourth
`cert_role` value, `CERT_ROLE_WEIGHING_CERTIFICATE` (fields.py), is
model-assigned from three signs: a *printed* (not handwritten) running number
in the header, a header company name that appears as "אתר קולט" on ordinary
source documents, and a weighing table (משקל/שעה/תאריך/מס' רכב/טרה/נטו) that's
mostly blank or hand-filled. Deliberately **not** keyed off the printed title:
a document whose header misprints "תעודת משלוח מס'" is still classified here if
its structure matches. Same per-page, model-driven, supplier-agnostic approach
as the other three cert_role values.

**Extraction is deliberately crippled for this role.**
`extractor._strip_weighing_certificate()` force-blanks every FIELD_DEFS field
except an allow-list (`_WEIGHING_CERTIFICATE_KEPT_FIELDS`:
document_type/cert_role/certificate_number/confidence/notes) - applied
unconditionally, *not* trusting the model to have obeyed the field
description. Per spec this document is not a data source for anything but its
own printed number: a plausible-but-unauthorized date/driver/vehicle read off
it is worse than a blank, because it would compete with the delivery note's
real value. None of the ordinary certificate checks run on such a record
(same skip-path shape as DOCUMENT_TYPE_OTHER) - they would all fire on a
document this feature blanks on purpose.

**The matching is many-to-many by number identity, and lives in pipeline.py,
not extractor.py.** `pipeline._cross_check_weighing_certificates()` runs once
per `process_files()` batch, across every file in it - upload order, page
order, and file boundaries are all irrelevant (contrast
`extractor._detect_and_cross_check_pairs`, which is strictly same-PDF/
adjacent-page). It builds **two** `{normalized_number: [records]}` indexes up
front - one keyed on each weighing certificate's `certificate_number`, one on
each delivery note's `reference_number` - so a batch with 50 delivery notes
and 50 weighing certificates resolves every pair by key lookup rather than by
scanning for a first match. Numbers compare through
`extractor._normalize_reference` (deliberately imported rather than
reimplemented, so the two cross-checks cannot drift on leading zeros/casing).

A link is made **only when the number is unique on both sides**
(`len(deliveries) == 1 and len(matches) == 1`). Four outcomes:
- **Unique 1:1** - the "מספר אסמכתא" column (`certificate_or_reference`) is
  rewritten to the number as read off the *weighing certificate itself*, and
  that document's file/page are recorded on the delivery record
  (`WEIGHING_SOURCE_FILE_KEY`/`WEIGHING_PAGE_NUMBER_KEY`) for the second
  hyperlink.
- **Delivery note has an אסמכתא, no weighing certificate carries it** - the
  column is *cleared* and `_NO_WEIGHING_MATCH_NOTE` added, per spec. Note the
  consequence: such a row no longer displays its own `certificate_number`
  there, which is what the spec asked for.
- **Number duplicated on either side** - no match is chosen (never "pick the
  first": an arbitrary link silently attributes a weighing certificate to the
  wrong shipment). Both sides get a note naming the counts, and confidence is
  escalated to at least "בינונית" - the spec only asked for a note, but
  without escalation the row stays uncolored and the note is invisible in the
  red/yellow review scan this tool is built around.
- **Weighing certificate matched nothing** - stays unmatched and is listed on
  the "מעקב פנימי" sheet (below). It never gets a row.

A delivery note with **no** `reference_number` at all is left completely
alone - existing `certificate_or_reference` behavior is unchanged for it.

**Why the two pairing mechanisms do not collide:** this one only ever writes
`certificate_or_reference` and the two link keys - never `certificate_number`
or `reference_number`, which are what `_cross_check_document_pair()` compares.
So a page can be half of an adjacent ברוטו/טרה/נטו pair *and* be cross-matched
to a separate תעודת שקילה, with neither check disturbing the other's inputs.
Verified explicitly in the smoke test below.

**`ProcessResult` gained a third field, `weighing_certificates`** - these
records are kept out of `.records` entirely rather than tagged-and-filtered.
That is what gives the "never a row of its own" guarantee for free on both the
Excel main sheet *and* streamlit_app.py's results table, with no filtering
code in either. Not `.skipped` either - those are irrelevant documents; these
are real data. The two `records, skipped = process_files(...)` unpack sites
(main.py, streamlit_app.py) had to become attribute access as a result; that
is the entire extent of the streamlit_app.py change besides state plumbing -
no UI was touched.

**Two hyperlinks needs two cells.** Excel supports exactly one hyperlink per
cell, so the spec's suggested single joined cell
(`"a.pdf (עמוד 3) | b.pdf (עמוד 1)"`) could only ever link to one of the two
files. `excel_writer.py` instead appends a trailing main-sheet column,
"קובץ תעודת שקילה", carrying the weighing certificate's own link.
`fields.WEIGHING_SOURCE_FILE_COLUMN` defines it **outside `EXCEL_COLUMNS` on
purpose**: streamlit_app.py derives its editable table's columns from
`EXCEL_COLUMNS`, so an ordinary entry there would have silently added a column
to the UI too. Appended *after* `source_file`, so no existing column index
shifts (`_write_summary_sheet`'s `_KEYS.index("source_file")` is unaffected),
and `read_existing_records()` maps its header back explicitly - without that,
re-extending a project file would silently drop the second link from every row
already written to it.

**Orphan weighing certificates go in their own block on "מעקב פנימי"**
(`_write_orphan_weighing_section`), below a blank spacer row, with its own
header - not as extra rows in the tracking table, because its columns are
genuinely different (a printed number and nothing else) and because mixing it
in would break that sheet's "row N here lines up with row N on the main sheet"
property. The autofilter ref is captured *before* this block is appended, so
it still covers the tracking table only. Only unmatched ones are listed: a
matched certificate carries `WEIGHING_MATCHED_KEY` and is already represented
by the second hyperlink on its delivery note's row.

**Durability fix, same day: orphans used to be erased on project reuse.**
`write_records()` rebuilds "מעקב פנימי" from scratch on every write, and an
unmatched weighing certificate has no main-sheet row for
`read_existing_records()` to bring back - so reopening a project and
processing one more batch silently deleted every previously-reported orphan
from the file, defeating the whole point of that block.
`excel_writer.read_existing_weighing_certificates()` now parses the orphan
block back (keyed off `_ORPHAN_WEIGHING_TITLE`, reading until the first blank
row), and streamlit_app.py seeds
`st.session_state.weighing_certificates` from it on the project-reuse path,
right next to the existing `read_existing_records()` call.
`_write_orphan_weighing_section()` is now deduplicated by
`(source_file, page_number, certificate_number)`, since the same certificate
can legitimately arrive both from the reloaded sheet and from a re-upload.
Recovered orphans are deliberately **not** re-matched against a later batch's
delivery notes - matching stays batch-scoped per spec; this only stops the
sheet from losing them. main.py needs no equivalent: the CLI has no
history-preserving path at all (it rewrites the file from that run's records).

**The first live run failed, and why it's worth reading.** Run against two
real documents (`input/תעודת_משלוח_102050.png`,
`input/תעודת_שקילה_70483.png` - note the filenames are swapped relative to
their contents; the file named "משלוח" holds the weighing certificate and
vice versa), `cert_role` came back **blank** on a textbook-matching weighing
certificate, so it got its own main-sheet row and nothing cross-matched.

Root cause: the original field description listed as one of its three signs
"the header company name appears as אתר קולט on ordinary source documents" -
information the model **cannot have from a single page**. This is the exact
failure mode this file already records for `CERT_ROLE_WEIGHING` on
2026-08-30, reintroduced for the same reason: a classification criterion that
silently requires cross-document knowledge makes the model default to blank
rather than guess. **When adding a cert_role value, every sign must be
judgeable from the page in front of the model, and the description must say
outright that confirming a matching document exists is not its job.**

The rewrite leads with the one decisive, page-local discriminator - **the
weighing table carries no printed values** (נטו/טרה/מס' רכב/תאריך/שעה/משקל
cells empty or handwritten), which is precisely what separates this from
`CERT_ROLE_WEIGHING`, whose ברוטו/טרה/נטו table is computer-printed and full -
then supporting page-local signs (printed serial in the header, blank
fill-in-by-hand lines, handwritten signatures), then explicit contrasts
against `CERT_ROLE_WEIGHING` and `CERT_ROLE_COMPLETION`. It also warns
explicitly that **the printed title on this real document misprints
"תעודת משלוח מס'"** while being a weighing certificate. After the rewrite the
same two documents classified correctly on the first try.

**Live run result (2026-09-15, `claude-sonnet-5`, 2 vision calls):** weighing
certificate correctly tagged with only its printed number `70483` extracted
and every other field verifiably stripped (date/driver/vehicle were all
legible on the page and none survived); one main-sheet row, not two; its
"מספר אסמכתא" column showing `70483` - the number off the weighing
certificate, not the delivery note's own `102050`; and two separate,
differently-targeted hyperlinks in "קובץ מקור" / "קובץ תעודת שקילה". The
weighing certificate was deliberately processed **first** in the batch, ahead
of the delivery note it matches, confirming order independence on real data.
Incidentally a good coexistence demonstration: the delivery note itself came
back tagged `CERT_ROLE_WEIGHING` (it genuinely is a computerized weighing
slip), which changed nothing - no adjacent COMPLETION page existed, so that
check simply didn't fire, while the new cross-match did.

**Pre-existing bug this run surfaced, NOT caused by this feature and not
fixed here:** on the delivery note, `quantity` came back `21500`, which is
the **טרה**, not the נטו (`41,260`). `gross_weight`/`tare_weight` came back
`62,760`/`41,260` - i.e. נטו was read into the tare field. The three boxes on
that document read `טרה: 21,500 | ברוטו: 62,760 | נטו: 41,260` left-to-right,
so this looks like RTL box-order confusion on this specific layout.
`_flag_gross_tare_mismatch()` cannot catch it: 62760 − 41260 = 21500 exactly
matches the reported quantity, so the arithmetic is self-consistent while
טרה and נטו are swapped. Reproduced identically on the run *before* any
prompt change in this session, so it predates this work. Fixing it means
touching the `quantity`/gross/tare field descriptions, which this file
already flags as the most incident-prone prompt surface in the project -
worth its own change with its own verification, not a drive-by edit.

**Testing done (2026-09-15):** ~60 fabricated-record pure-function checks -
the field-stripping allow-list, all four matching outcomes, duplicate numbers
on *each* side independently, order independence (weighing certificate listed
first), empty-batch no-op, coexistence with the adjacent-page pair check -
plus a `write_records()` run asserting the two hyperlinks exist, differ, and
target the right files, that an ordinary row's extra cell stays empty with no
stray link, that the orphan section lists the unmatched ones and not the
matched one, and a `read_existing_records()` -> `write_records()` round-trip
proving the second link survives re-extension. Plus a
`streamlit.testing.v1.AppTest` run (real `.streamlit/secrets.toml`
`APP_PASSWORD` through the `st.form` gate - note `at.session_state` has no
`.get()`, use `in`/`[]`, and an unauthenticated run renders *nothing*, so
assert the gate passed or every later assertion is vacuous) confirming the
results table renders with exactly the delivery rows, the weighing certificate
absent, and no new UI column. **Verified live** against the two real documents
described above (two `claude-sonnet-5` vision calls), after the first attempt
failed and the field description was rewritten - see that discussion above
before touching the cert_role wording again.

**Heredoc gotcha, for whoever edits these files next:** patching via
`python - <<'PY'` through the Bash tool in this environment is not literal -
a `\n` in the script source collapses to a real newline (so a match string
anchoring on `print(f"\nנשמר: ...")` in main.py silently fails to match), and
an apostrophe inside the body can break shell parsing outright. Write the
patch script to a file and run that instead.

## Field order in FIELD_DEFS is load-bearing, and two prompt lessons (2026-09-15)

Three bugs found while verifying the weighing-certificate feature above
against real documents. All three share one root cause pattern, and it is the
most reusable thing in this file: **under forced tool use with thinking off,
the model answers fields in schema order, and it cannot answer a question that
requires information it doesn't have.** Violate either and it silently returns
a wrong value or a blank - never an error.

### 1. `quantity` returned the טרה instead of the נטו - fixed by REORDERING fields

On `input/תעודת_שקילה_70483.png` (weights row reading
`טרה 21,500 | ברוטו 62,760 | נטו 41,260`), the pipeline reported
`quantity = 21,500` - the tare - with `gross_weight = 62,760` and
`tare_weight = 41,260`. Reproduced on three consecutive runs.

The diagnosis mattered more than the fix, because two rounds of prompt
tightening changed nothing:
- Asked plainly, outside the tool schema ("report each box's label and its
  number, verbatim"), the same model answered **perfectly**:
  `טרה=21,500 / ברוטו=62,760 / נטו=41,260`. So it was not a reading or
  label-association failure.
- Not `image_preprocessing` (identical output with and without) and not our
  code (dumping the raw `tool_use.input` showed `21500` coming straight from
  the model).
- **`FIELD_DEFS` listed `quantity` about twelve fields before
  `gross_weight`/`tare_weight`.** The model had to commit to a quantity
  before it had looked at the labelled ברוטו/טרה boxes, then back-filled
  `tare_weight = 41,260` so that ברוטו − טרה still equalled the quantity it
  had already emitted. Moving those two entries to sit immediately *before*
  `quantity` fixed it outright - `quantity=41,260`, `tare=21,500` - with no
  prompt change at all.

**Rule: a field that is derived from, or constrained by, other fields must be
listed after them in `FIELD_DEFS`.** Everything else that consumes that list
builds a dict keyed by name and is order-independent (`empty_record()`,
`_extract_page`'s record comprehension, `examples_library`), and
`EXCEL_COLUMNS` governs column order separately - so reordering is safe, but
never assume it is inconsequential. There is a comment saying so on the list
itself now.

### 2. Two classification fields asked the model what it could not know

The same mistake, twice, in two different fields - and this file already
recorded it once for `CERT_ROLE_WEIGHING` on 2026-08-30:

- **`cert_role`** - the first version of the weighing-certificate description
  listed "the header company name appears as אתר קולט on ordinary source
  documents" as an identifying sign. The model has one page; it cannot check
  other documents. It returned blank on a textbook-matching document.
- **`reference_number`** - the long-standing description said to fill it
  "only when a separate field explicitly refers to **another document's**
  certificate number", and to leave it blank otherwise. On a delivery note
  with `אסמכתא : 70483` printed plainly in its header, the model left it
  **blank on two runs out of three** - it can't verify what the number refers
  to, so it declined. That silently broke the whole cross-match (no
  `reference_number` → no match attempt → the אסמכתא column fell back to the
  note's own `102050` and its weighing certificate was filed as an orphan).
  Note how this failed: not with an error, but as a plausible-looking row.

Both are now phrased as mechanical instructions - copy what is printed next to
this label; you are not being asked whether it points anywhere; the matching
happens in code. **When adding or editing a field description, check whether
answering it requires anything beyond the page in front of the model. If it
does, the field will quietly come back blank.**

### 3. A symmetric invariant cannot disambiguate - don't ask the model (or the code) to

`gross - tare = net` and `gross - net = tare` are the same equation, so a
tare<->net swap satisfies it exactly. Two consequences, both learned the hard
way here:
- `extractor._flag_gross_tare_mismatch()` provably cannot catch that swap. It
  now checks what it *can* decide - positivity, and that ברוטו is the largest
  of the three (physically mandatory, catches every mislabeling involving the
  gross value) - and its docstring says outright that the swap is out of
  reach.
- **Telling the model to "verify that נטו = ברוטו − טרה" was useless for the
  same reason** - the wrong labeling passes that check too. The instruction
  that helped was the opposite: *this arithmetic cannot confirm your
  assignment, only the printed labels can, and do not infer from magnitude -
  a loaded truck's payload can exceed its own empty weight.*

`pipeline._flag_swapped_tare_net()` is the one check that can catch a swap,
because it brings in evidence from outside the three numbers: the same
vehicle's tare on another document in the batch (a vehicle's tare is roughly
constant, its net load is not). Flags both candidate values in a note, never
auto-corrects, and needs a second document for that vehicle to say anything -
tare weights live on the "מעקב פנימי" sheet, which `read_existing_records()`
does not read, so no cross-run history is available to it.

### Verification standard used here, worth reusing

One passing live run was not treated as evidence, because these failures are
nondeterministic - `reference_number` came back blank on some runs and
correct on others with identical code. The acceptance check
(`quantity=41,260`, `אסמכתא=70,483`, one main row, the weighing certificate
stripped to its number and marked matched, two differently-targeted
hyperlinks, `gross=62,760`, `tare=21,500`) was run **three consecutive times
and passed 12/12 criteria on all three**, with the weighing certificate
deliberately processed first each time. For an extraction-quality change,
re-run the real check several times before believing it.

## Per-batch מרחב/אתר/סוג-פסולת selection replaces extracting them (2026-09-15)

The single biggest change to what the model is asked to do since this project
started. Three fields - `region`, `site`, `waste_type` - are **no longer
extracted from the documents at all**. The user picks them once per upload
batch from closed lists, and they are written onto every row that batch
produces.

**Why this is an improvement and not just a UI feature:** those were exactly
the two fields the "Known limitation: don't guess is prompt-only" section
above documents as this pipeline's most persistent failure - a certificate
often simply does not say which מרחב it belongs to, so the model inferred one
and admitted the guess in `הערות`. That whole class of bug is now gone by
construction rather than caught by validation. `FIELD_DEFS` went from 21
fields to 18, and neither closed list (`WASTE_TYPES`, `REGIONS`) is sent to
the model any more.

### Where each piece lives

- **`app/sites_config.py`** (new) - the two-level מרחב -> אתר map, in display
  order. Reference data with its own update cadence, deliberately not in
  fields.py (the extraction schema). Nothing here reaches the model.
  `fields.REGIONS` is set from `REGION_NAMES` so the dropdown, the summary
  sheet and the results-table dropdown cannot disagree.
- **`pipeline.BatchSelection` + `apply_batch_selection()`** - applied inside
  `process_files()`'s per-record loop, *before* the normalizer and
  quantity-outlier check (the latter is keyed by site, so it needs the real
  one). Unconditional assignment, not fill-if-blank. Applied to `records`
  only - never `skipped` or `weighing_certificates`, neither of which becomes
  a row.
- **`streamlit_app.py`** - the "שיוך האצווה" card, three cascading
  selectboxes, gating both the `file_uploader` and the process button.
- **`main.py`** - optional `--region/--site/--waste-type`. Without them a CLI
  run leaves the three columns blank, because nothing extracts them any
  more; the flags exist so the CLI isn't crippled by this change.

### Three consequences that are easy to get wrong

**1. "No sites" is a complete selection, not a missing field.**
חטיבת הפיתוח has no second level. Gating on `if not site` would make that
מרחב permanently unusable - the button would never enable and the message
would say a site is missing when none exists. Always gate on
`sites_config.region_has_sites(region)`, and use `is_valid_selection()` which
encodes both rules (a known מרחב, plus one of *its own* sites, or no site
when it has none). The UI shows a disabled "(לא רלוונטי למרחב זה)" field for
that case rather than an empty dropdown a user would try to fill.

**2. Changing מרחב must clear the chosen site.** A cascading dropdown invites
exactly one bug: the site from the previous מרחב stays selected,
`is_valid_selection()` correctly rejects the pair, and the upload button goes
dead with no visible reason. `streamlit_app.py` compares the new region
against `st.session_state.batch_region` and resets `batch_site` on a change;
the site widget's key also includes the region name, so Streamlit treats it
as a brand-new widget. Covered by its own AppTest case.

**3. The region rename breaks exact-match aggregation on old rows.** The old
closed list was `["צפון", "דרום", "מרכז", "מטה"]`; the new one is the six real
org names. Rows written before this change keep their old short label (an
explicit requirement - existing projects are not rewritten). The
`סיכום` sheet counts by exact string match, so untreated, every pre-rename
row would fall into the "ללא מרחב מזוהה" catch-all and be reported as having
no region - simply false. Each region row therefore sums its current name
**plus** any legacy label mapping to it (`fields.LEGACY_REGION_TO_CURRENT`),
still using only additive exact-value COUNTIF/SUMIF, per this sheet's
verified-formula rule above. `"מטה"` is spelled identically in both lists and
is deliberately absent from that map.

### site left the fuzzy normalizer, on purpose

`normalize.NORMALIZED_FIELDS` is now `["supplier_or_carrier"]` only. Keeping
`site` would have been actively harmful once it comes from a closed list:
`NameNormalizer` exists to collapse OCR/handwriting spelling variants of a
model-read name, and it would happily rewrite an exact, user-chosen site to a
high-scoring neighbour ("נגב צפוני" vs "נגב מרכזי"), silently corrupting a
value that was never uncertain. `supplier_or_carrier` ("אתר קולט") is still
model-read and still normalized.

### Deliberately NOT changed

- **The results-table `אתר` column stays free text.** Turning it into a
  closed-list dropdown would blank or break free-text site values in existing
  projects. The `מרחב` column there does offer the new list, and
  `_selectbox_options()` already appends any present-but-unlisted value, so a
  historical `"צפון"` still renders and stays editable.
- **Manual-Excel rows don't get the selection.** They come from a
  human-filled sheet carrying its own values; overwriting them would discard
  real typed data, which is the whole premise of `app/excel_input.py`.

### Verification

Two new suites: 40 pure-function checks (the mapping, all the
`is_valid_selection` edge cases, unconditional overwrite, `None` being a
no-op so existing projects are untouched, end-to-end through `process_files`
with extraction mocked to return *wrong* values and be overridden) and 31
`AppTest` checks driving the real widgets (gating with nothing selected,
cascade contents per מרחב, full selection enabling upload, חטיבת הפיתוח
completing without a site, and the switch-region-clears-stale-site case).
Plus a real API run with a selection deliberately contradicting the document
(מרחב צפון/גליל/פסולת בנין on an אוליצקי "עודפי עפר קידוח" certificate), and
Playwright screenshots of the live app in each state.

## ק.מ.מ consolidated monthly summary report (2026-09-15)

A third document type, and the first one this project does **not** send to the
model at all. `app/kmm_report.py` parses it deterministically from the PDF's
own text layer.

**Why not vision.** This report is a digitally generated PDF with a real text
layer - no handwriting, no photograph, no OCR. Reading the text that is
already in the file is faster, free, and *more* accurate than asking a vision
model to look at a picture of it. `pypdfium2` was already a dependency
(extractor.py uses it to rasterize pages), so this added no package.

**One file becomes many rows.** Every other document type here yields one row
per page; this one is structured customer-center -> station -> month ->
material. The real reference file (`input/דוח כמויות מקורות 1-6.26.pdf`,
2 pages) contains **16 stations and 68 data rows totalling 18,670 ק"ג**
(קרטון 53 rows, נייר לבן 15, across 6 months). The report has no printed
total, so those numbers were counted independently from the raw text layer
before the parser existed and are now hard assertions in the test suite - a
future regression that drops or duplicates a row fails loudly instead of
quietly changing a sum.

### The batch-selection conflict, and how it was resolved

The per-batch selection screen (see the section above) sets one
מרחב/אתר/סוג-פסולת for a whole upload. This document breaks that assumption
three ways, and the third is the one that isn't obvious:

1. 16 stations in one file.
2. Two different materials in one file.
3. **12 of the 16 station names are not members of the closed
   `sites_config.REGION_SITES` list at all** (e.g. "מקורות ראש פינה",
   "מקורות אתר ספיר"), and the stations span at least four different מרחבים.

Resolved with the user: **for this document type the document wins for `site`
and `waste_type`, and `region` comes from the selection.** The reasoning is
that the selection exists because ordinary certificates *don't state* these
fields, so the model used to guess them - while this report prints the station
and material explicitly, per row, machine-readably. Applying the selection
here would be following the mechanism against its own purpose and discarding
the only reliable per-row identifying data in the file.

`region` is the one field the report doesn't state per row, so it takes the
selection - the user's explicit choice, accepted with its known downside (the
selected מרחב is wrong for any station belonging to a different one). To
restore the "which rows can I trust" signal that gives up,
`pipeline.apply_kmm_selection()` notes and escalates to "בינונית" any row
whose station name explicitly names a *different* מרחב than the one selected
(10 of 68 rows on the real file). Region is still always filled from the
selection; the note only annotates disagreement.

### Two parsing traps, both of which bit on the first run

**1. RTL text order. Solved with character geometry, not token reversal - see
the dedicated section below**, which supersedes the first implementation
described here. Short version: the extracted character stream is in neither
logical nor reliably visual order, so `_logical_lines()` takes line breaks
from pdfium and character order from each character's x coordinate. The
original approach reversed each line's *tokens*, which fixed word order but
not punctuation, and made marker matching fail outright (`שם תחנה:` arrives
with the colon wedged in the middle, `שם :תחנה`, so matching `"שם תחנה"`
silently never fired - on the first run that flagged **all 68 rows** as
unparsed).

**2. A station block continues across a page break.** Its rows on the next
page have no repeated `שם תחנה` header. Station state was initialized per
page, which orphaned the first 6 rows of page 2 - exactly **1,305 of the
18,670 ק"ג**. Verified directly: page 1's last content line is the `KM113549`
header with zero data rows under it, and page 2's first 6 data rows belong to
it. That state now lives outside the page loop.

Both bugs were caught only because `_unparsed_record()` turns a data-looking
line that fails to parse into a flagged row carrying the raw text, instead of
skipping it. **A silently dropped row in a 68-row report is invisible** - keep
that behavior if this parser is ever reworked.

A third detail in the same family: **the date is printed only when it
changes.** Rows continuing a month leave the cell empty, so the month is
forward-filled from the last row that had one, reset at each new station.

### RTL text extraction: geometry, not token reversal (2026-09-15)

The first implementation reversed each line's tokens. That was replaced after
checking what the PDF actually contains, and the check is the reusable part:
**the extracted character stream is in neither logical nor reliably visual
order.** Printing each character's x coordinate for `מ"שח` shows it:

```
'מ' x=453.6   sorted right-to-left by x:
'"' x=458.7     ש(467) ח(462) "(458.7) מ(453.6)  ->  שח"מ   (correct)
'ש' x=467.0
'ח' x=462.4
```

So word order was never the real problem - it was already right. The defect
was punctuation attaching to the wrong neighbour, which token reversal cannot
reach: `בע"מ` came out `מ"בע`, `ו' סגור` came out `ו סגור'`, `ת"א` came out
`א"ת`, `ב"ש` came out `ש"ב`. **Those last two are Tel Aviv and Be'er Sheva -
wrong data, not cosmetics.**

`_logical_lines()` now does three things:

1. **Line breaks from pdfium's own text stream.** Reconstructing lines by
   clustering characters on their y coordinate was tried first and does NOT
   work here: the characters of one visual line spread over more than 5
   points, which split data rows into fragments and dropped the spaces
   between words (`'י10222026ילי115ייט'`). Don't retry it.
2. **Character order within a line from each character's x coordinate**
   (`get_charbox`) - the true visual left-to-right order.
3. **Visual -> logical**: reverse the line, then re-reverse each left-to-right
   run (`_LTR_RUN`: Latin, digits, and the separators inside them). Without
   step 3, `KM103527`, `1,140` and `20/08/26` come out backwards.

**`python-bidi` is the wrong tool for this and was deliberately not used.** It
implements the Unicode Bidi Algorithm in the logical -> visual direction
(`get_display`); this problem needs the inverse, and the UBA is not trivially
invertible. The page's own coordinates are better evidence than any
reconstruction, and using them added no dependency.

A bonus: lines now arrive in the document's real column order
(`ינו-2026 קרטון 1022 איסוף קרטון לפי קוב 115` - תאריך, סוג החומר, מק"ט,
תאור מוצר, משקל), so `_parse_data_row()` reads columns from the ends inward
instead of unreversing anything, and `_DATE_RE` matches `ינו-2026` rather than
the old `2026-ינו`.

**Result, measured against the committed previous version rather than from
notes** (load the old module via `git show HEAD:app/kmm_report.py` - that
comparison is worth repeating if this layer is ever touched again): 9 of the
16 station names corrected, plus 3 addresses
(`ח"ר טובים 2` -> `ר"ח טובים 2`, `היצירה 2 .ת.א` -> `היצירה 2 א.ת.`,
`וילסון נקרא/6 גם לינקולן` -> `וילסון 6/נקרא גם לינקולן`); the 7 station
names that were already correct unchanged; and **every numeric and structural
field byte-identical** - quantity, waste_type, date, unit,
certificate_or_reference, page_number, 68 rows, 18,670 ק"ג.

One artifact remains and is correct to leave: `שח"מ` may be `שח"ם` with a
final mem in reality, but the PDF contains a regular mem, so the parser is
faithful to the source rather than guessing.

**This is the only place in the project that extracts PDF text.** Every other
document type is rasterized and read by the vision model (extractor.py), and
`excel_input.py` reads xlsx - so no other extracted field is affected by RTL
order, and this fix covers everything that is.

### Isolation from the other mechanisms

- **`reference_number` is deliberately left empty** even though
  `certificate_or_reference` holds the station's customer number
  (`KM103527`). `pipeline._cross_check_weighing_certificates()` indexes
  candidates by `reference_number`, so this is what guarantees a customer
  number can never be mistaken for a תעודת שקילה's printed number. Covered by
  a test that feeds it a weighing certificate numbered `KM103527`.
- **Duplicate detection skips these rows** (`duplicate_check._is_kmm_record`).
  Not an arbitrary exclusion: the check looks for one shipment reported twice
  through two channels, while this reports a month's total per station per
  material. And its rows are *designed* to resemble each other - 53 of 68 are
  the same material at the same station in different months and 115 ק"ג
  recurs constantly - so the date/site/waste_type/quantity heuristic would
  emit a blizzard of false "כפילות אפשרית" notes.
- **None of the vision-error safety nets run** - no NameNormalizer, no
  quantity-history outlier check, no gross/tare arithmetic, no
  date-plausibility flag. They all exist to catch a vision model's mistakes
  and nothing here was read by one; same reasoning `app/excel_input.py`
  already applies to hand-typed rows.
- **`DOCUMENT_TYPE_KMM_SUMMARY_REPORT` is deliberately NOT in
  `DOCUMENT_TYPES`**, which is the closed list the *model* chooses from. The
  model never sees this format, so it must not be offered it.

### Detection

`looks_like_kmm_report()` requires **both** an issuer marker (ק.מ.מ / k.m.m)
and a report-title marker on page 1. One signal is not enough: that same
company already appears in this project's history as an ordinary *invoice*
issuer (see the 2026-09-01 invoice-filtering note above), and a certificate
merely naming it as a supplier must not be misrouted into this parser.
Anything that isn't a readable PDF - a non-PDF, an image, a scanned PDF with
no text layer - returns False and falls through to the normal vision path
rather than erroring. Verified against the other real files in `input/`.

### Known cosmetic artifact, deliberately not fixed

Gershayim inside Hebrew abbreviations land on the wrong side of their token:
station `KM103527` reads `מ"שח מקורות ביצוע מ"בע ו סגור' מ-7` instead of
`שח"ם מקורות ביצוע בע"מ ו' סגור מ-7`. Word order is correct; only the
abbreviations are off, on 3-4 of the 16 station names. **Not auto-corrected on
purpose:** a rule that swaps sides around the gershayim would also turn `ק"ג`
into `ג"ק`, because both placements are legitimate in Hebrew (`בע"מ` puts it
before the last letter, `ק"ג` after the first). A small explicit map of known
abbreviations is the safe fix if this ever matters.

## The blanket RTL rule vs. st.data_editor: a silent display bug (2026-09-15)

**Read this before touching `streamlit_app.py`'s CSS block.** It cost a real
bug that shipped unnoticed through several changes, and the failure mode is
one no test in this repo can catch.

### What happened

`streamlit_app.py` opens its stylesheet with a universal rule:

```css
.stApp, .stApp * { direction: rtl; }
```

`st.data_editor` (and `st.dataframe`) do **not** render cells as DOM text -
they draw them on a `<canvas>` (glide-data-grid), doing their own text layout
and measurement. When that canvas's container inherited `direction: rtl`, the
grid clipped **every cell to its single rightmost character**:

| column | displayed | actual value |
|---|---|---|
| כמות מדווחת | `0` | `115.00` |
| מספר אסמכתא | `2` | `KM111062` |
| מרחב | `מ` | `מרחב מרכז` |
| אתר / מקור | `מ` | `מקורות חבל הירדן מטה אתר אשכול*-חודש` |
| סוג הפסולת | `ק` | `קרטונים` |
| רמת ביטחון | `ב` | `בינונית` |

The rightmost character is the *first* letter of Hebrew text and the *last*
digit of a number, which is why it looks like two different bugs at once.
**The quantity column showed `0` on every single row.**

### Why it survived so long

The data was never wrong. `write_records()` exported full correct values, and
the cell editor showed the full value the moment a cell was opened - only the
collapsed on-screen cells were unreadable. So every functional check passed
while the table was visibly broken.

**Crucially, `AppTest` cannot detect this.** `at.dataframe[0].value` returns
the DataFrame that was handed to the widget, which is correct by construction;
the corruption happens in browser canvas rendering, downstream of anything
AppTest observes. Nine passing suites said nothing. It was found only by
rendering the page in a real browser and looking at it.

### The fix, and why the obvious one doesn't work

```css
[data-testid="stDataFrame"], [data-testid="stDataFrame"] * { direction: ltr; }
```

Restoring `direction: ltr` on the grid container lets glide-data-grid lay out
its own text again. Column order is unaffected - it comes from the DataFrame,
not from CSS - and Hebrew inside cells still renders right-to-left, because
the canvas applies per-string bidi itself.

**`width="large"` is not the fix.** It was tried first: a 287px-wide column
still showed one character. This is not a "text doesn't fit" problem, so
column widths, font sizes and `column_order` are all dead ends.

### How to keep it from coming back

1. **Never let a universal `*` direction/writing-mode rule reach a
   canvas-rendered widget.** If the blanket `.stApp *` rule is kept, the
   `[data-testid="stDataFrame"]` exclusion must be kept with it. Deleting the
   exclusion silently reintroduces the bug - there is no error, no exception,
   and no failing test.
2. **After any change to that CSS block, look at the rendered table**, not
   just the test output. The check is ten seconds: run the app, process
   anything, and confirm `כמות מדווחת` shows a full number rather than a
   single digit. That column is the canary - a bare `0` where a weight
   belongs means this regressed.
3. **Screenshot-verify UI changes generally.** This repo's CSS is extensive
   and mostly targets `data-testid` attributes; AppTest validates behavior and
   values, never appearance. For a visual change, drive a real browser
   (Playwright is installed) and read the image.

## Editable results table (2026-09-15)

`st.data_editor` was already the results table; this change made three
closed-list columns behave like closed lists and locked down what must not be
edited.

- **`אתר / מקור` became a dropdown**, sourced from `sites_config` - the same
  config the upload selection card uses, so the two cannot drift.
- **Options go through `_selectbox_options()`, not the bare closed list.** A
  strictly-closed list would blank two kinds of legitimate existing value the
  moment the table renders: a free-text site on a row written before the
  selection screen existed, and a ק.מ.מ report's station name (12 of the 16
  in the real reference file are not members of `REGION_SITES`). The user
  still cannot *type* a new value - only pick one - which is what "closed
  list" means in the editor.
- **`קובץ מקור` and `מצב` stay read-only.** `source_file` is provenance: it
  carries the `(עמוד N)` page reference, the Excel hyperlink target, and
  `duplicate_check._is_manual_excel_record`'s origin test (by extension).
  There is also a second reason that is easy to miss - **a pandas Styler only
  applies to non-editable columns in `st.data_editor`**, so unlocking
  `קובץ מקור` would silently kill the red/yellow confidence row coloring,
  which this file describes elsewhere as the main UX mechanism of the tool.
  `מצב` is a symbol computed from confidence; editing it would mean nothing.
- **`כמות מדווחת` stays a `TextColumn`, not a `NumberColumn`** - see the
  2026-09-01 manual-Excel note above: a hand-typed quantity can be `כ-1.8`,
  and `NumberColumn` rejects or blanks that.

### The tracking sheet keeps the ORIGINAL confidence

The main table's confidence column is editable (useful for marking a row as
reviewed), but the "מעקב פנימי" sheet is the audit trail of what extraction
concluded and must not be rewritten by an edit - an explicit requirement.
Those two read the same `record["confidence"]`, so editing the table used to
overwrite the audit value; verified, not assumed.

`fields.EXTRACTED_CONFIDENCE_KEY` now freezes it. `pipeline.process_files()`
stamps it once per record **after** the batch-level cross-checks (those
legitimately downgrade a row, and that is part of what extraction concluded)
and before any edit is possible; `excel_input.read_manual_excel()` stamps its
own rows; `excel_writer._tracking_value()` reads the frozen value, falling
back to the live one only for a record that predates the key. Underscore
prefix, so it never reaches a spreadsheet column.

`הערות` is deliberately kept to ~4-5 words (per the field description and
system prompt in `extractor.py`/`fields.py`, 2026-08-23) — e.g. "כתב יד לא
קריא" rather than a full sentence explaining what was inferred and why. This
was a user readability request, not a data-quality fix, so it doesn't change
anything about the guessing behavior above - a short note can still describe
a guessed value, it just won't spell out the reasoning anymore.
