"""Entry point: process every certificate in input/ and write the consolidated Excel file.

Usage:
    python main.py              # process every file in input/
    python main.py --file X.pdf # process a single file in input/ (useful while testing)
"""
import argparse
import sys

from app import config, sites_config
from app.excel_writer import write_records
from app.fields import WASTE_TYPES, WEIGHING_MATCHED_KEY
from app.pipeline import BatchSelection, process_files


def main() -> None:
    parser = argparse.ArgumentParser(description="עיבוד תעודות פינוי פסולת לקובץ Excel מרוכז.")
    parser.add_argument("--file", help="שם קובץ בודד בתוך input/ לעיבוד (ברירת מחדל: כל הקבצים)")
    # מרחב/אתר/סוג פסולת אינם מחולצים מהמסמך יותר (ראו app/pipeline.py's
    # apply_batch_selection) - בלי הדגלים האלה שלוש העמודות האלה יישארו
    # ריקות בהרצת CLI. הדגלים אופציונליים במכוון, כדי שהרצה קיימת לא תישבר.
    parser.add_argument("--region", choices=sites_config.REGION_NAMES, help="מרחב לכל האצווה")
    parser.add_argument("--site", help="אתר/יחידה לכל האצווה (מתוך האתרים של המרחב הנבחר)")
    parser.add_argument("--waste-type", choices=WASTE_TYPES, help="סוג פסולת לכל האצווה")
    args = parser.parse_args()

    selection = None
    if args.region or args.site or args.waste_type:
        if not args.region or not args.waste_type:
            sys.exit("--region ו---waste-type נדרשים יחד כשמשתמשים בבחירת אצווה.")
        site = args.site or ""
        if not sites_config.is_valid_selection(args.region, site):
            allowed = sites_config.sites_for(args.region)
            sys.exit(
                f"שילוב מרחב/אתר לא תקין: {args.region!r} + {site!r}. "
                + (f"אתרים חוקיים: {', '.join(allowed)}" if allowed
                   else f"למרחב {args.region} אין אתרים - אל תעביר --site")
            )
        selection = BatchSelection(region=args.region, site=site, waste_type=args.waste_type)

    if not config.ANTHROPIC_API_KEY:
        sys.exit("ANTHROPIC_API_KEY חסר - ודא שקובץ .env בתיקיית הפרויקט מכיל אותו.")

    config.INPUT_DIR.mkdir(parents=True, exist_ok=True)

    if args.file:
        candidate = config.INPUT_DIR / args.file
        if not candidate.is_file():
            sys.exit(f"הקובץ לא נמצא: {candidate}")
        files = [candidate]
    else:
        files = sorted(
            p
            for p in config.INPUT_DIR.iterdir()
            if p.is_file() and p.suffix.lower() in config.SUPPORTED_EXTENSIONS
        )

    if not files:
        sys.exit(f"לא נמצאו קבצים לעיבוד בתיקיית {config.INPUT_DIR}")

    def on_progress(index, total, path):
        print(f"מעבד ({index}/{total}): {path.name} ...")

    def on_error(path, exc):
        print(f"  שגיאה בעיבוד {path.name}: {exc}")

    result = process_files(files, on_progress=on_progress, on_error=on_error, selection=selection)
    records, skipped = result.records, result.skipped

    if skipped:
        print(f"\n{len(skipped)} עמודים דולגו (לא זוהו כתעודות פינוי):")
        for item in skipped:
            print(f"  {item.get('source_file', '')}: {item.get('notes', '')}")

    # Separate תעודת שקילה pages that matched no תעודת משלוח in this run -
    # they get no row anywhere, so surface the count here (and in the
    # 'מעקב פנימי' sheet write_records() builds from the same list below)
    # rather than letting them disappear silently.
    orphan_weighing = [
        c for c in result.weighing_certificates if not c.get(WEIGHING_MATCHED_KEY)
    ]
    if orphan_weighing:
        print()
        print(f"{len(orphan_weighing)} תעודות שקילה ללא תעודת משלוח תואמת (נרשמו בגיליון 'מעקב פנימי'):")
        for item in orphan_weighing:
            number = item.get('certificate_number', '') or '(לא נקרא)'
            print(f"  {item.get('source_file', '')}: מספר {number}")

    output_path = config.OUTPUT_DIR / config.OUTPUT_FILENAME
    write_records(records, output_path, weighing_certificates=result.weighing_certificates)
    print(f"\nנשמר: {output_path} ({len(records)} תעודות)")


if __name__ == "__main__":
    main()
