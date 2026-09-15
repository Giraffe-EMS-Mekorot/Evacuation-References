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

`הערות` is deliberately kept to ~4-5 words (per the field description and
system prompt in `extractor.py`/`fields.py`, 2026-08-23) — e.g. "כתב יד לא
קריא" rather than a full sentence explaining what was inferred and why. This
was a user readability request, not a data-quality fix, so it doesn't change
anything about the guessing behavior above - a short note can still describe
a guessed value, it just won't spell out the reasoning anymore.
