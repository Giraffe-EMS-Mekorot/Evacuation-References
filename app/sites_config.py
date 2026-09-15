"""The closed, two-level מרחב -> אתר/יחידה org structure the user picks from
before uploading a batch (see streamlit_app.py's selection card).

**This is reference data, not extracted data.** As of 2026-09-15 the model is
no longer asked to read מרחב / אתר-מקור / סוג-פסולת off a certificate at all -
those three columns are filled from the user's per-batch selection instead
(see pipeline.apply_batch_selection and fields.BATCH_SELECTION_KEYS). That
removed three fields from FIELD_DEFS, including the two the "Known
limitation: don't guess is prompt-only" section of CLAUDE.md documented as
this pipeline's most persistent guessing problem - a certificate simply does
not reliably say which מרחב it belongs to, so asking was always the weak part.

Lives in its own module rather than in fields.py because it is organizational
reference data with its own update cadence (sites get added/renamed as the
org changes), while fields.py is the extraction schema. Nothing here reaches
the model.

Ordering is deliberate and preserved (dicts keep insertion order): the
dropdowns render in exactly this order, which is the order the user supplied,
not alphabetical.
"""
from typing import Dict, List

# מרחב -> its אתרים/יחידות. An empty list is meaningful, not a placeholder:
# חטיבת הפיתוח genuinely has no second level, and the UI hides the site
# dropdown entirely for it (see streamlit_app.py) rather than showing an empty
# or disabled control. Any code reading this must therefore treat "no sites"
# as a valid, supported state - see region_has_sites().
REGION_SITES: Dict[str, List[str]] = {
    "מטה": ["מחסן מרכזי (שחם)", "יחדית רכש"],
    "מרחב צפון": ["גליל", "עמקים"],
    "מרחב דרום": ["נגב צפוני", "נגב מרכזי", "ערבה"],
    "מרחב מרכז": ["צפון ירקון", "דרום ירקון", "שפדן"],
    "מרחב המוביל": ["מוביל ארצי", "מובלים", "תחנות"],
    "חטיבת הפיתוח": [],
}

# The closed list of מרחבים, in display order. fields.REGIONS is set from this
# so there is exactly one source of truth - excel_writer.py's summary sheet and
# streamlit_app.py's results-table dropdown both read it from there.
REGION_NAMES: List[str] = list(REGION_SITES)

# Every site name across all מרחבים, for the results-table dropdown (which
# edits an already-written row and so can't know which מרחב the user is
# currently picking under).
ALL_SITE_NAMES: List[str] = [site for sites in REGION_SITES.values() for site in sites]


def sites_for(region: str) -> List[str]:
    """The אתרים under `region`, or [] for an unknown region or one with no
    second level (חטיבת הפיתוח). Returns a copy, so a caller can't mutate the
    config by editing the returned list.
    """
    return list(REGION_SITES.get(region, ()))


def region_has_sites(region: str) -> bool:
    """Whether `region` has a second level to choose from at all.

    False for חטיבת הפיתוח - the one case the site dropdown is hidden and a
    blank site is the correct, complete answer rather than a missing field.
    Callers gating a "ready to upload?" check must use this instead of
    testing the site value for truthiness, or that מרחב can never be used.
    """
    return bool(REGION_SITES.get(region))


def is_valid_selection(region: str, site: str) -> bool:
    """Whether (region, site) is a complete, internally consistent choice:
    a known מרחב, plus one of its own אתרים - or no אתר at all when that מרחב
    has none. Guards against a stale site left over from a previously-chosen
    מרחב (the exact bug a cascading dropdown invites).
    """
    if region not in REGION_SITES:
        return False
    if not region_has_sites(region):
        return not site
    return site in REGION_SITES[region]
