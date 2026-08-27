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
certificate committed to the repo to run the real pipeline against.

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

**`app/fields.py` holds two related but distinct lists — don't conflate them.**
- `FIELD_DEFS`: `(field_name, description_for_model, json_type)` tuples for
  what Claude actually extracts from the certificate. Drives the tool-use
  JSON schema in `extractor.py` and the record dict keys used through the
  pipeline. Add/remove a *model-extracted* field here.
- `EXCEL_COLUMNS`: `(record_key, hebrew_header)` tuples defining the exact
  output column order and headers, independent of extraction order. Includes
  keys that aren't in `FIELD_DEFS` at all — `year`/`month` (derived in code,
  see below) and `source_file` (filled in by the pipeline). Add/reorder a
  *spreadsheet* column here — this list is what `excel_writer.py` iterates,
  so column order changes never require touching the extraction code.

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
  but were removed on user feedback (2026-08-23) — not needed. If asked to
  bring vehicle/driver tracking back, it's a `FIELD_DEFS` addition, not a
  restoration of deleted logic (there isn't any to restore).

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
