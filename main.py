"""Entry point: process every certificate in input/ and write the consolidated Excel file.

Usage:
    python main.py              # process every file in input/
    python main.py --file X.pdf # process a single file in input/ (useful while testing)
"""
import argparse
import sys

from app import config
from app.excel_writer import write_records
from app.pipeline import process_files


def main() -> None:
    parser = argparse.ArgumentParser(description="עיבוד תעודות פינוי פסולת לקובץ Excel מרוכז.")
    parser.add_argument("--file", help="שם קובץ בודד בתוך input/ לעיבוד (ברירת מחדל: כל הקבצים)")
    args = parser.parse_args()

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

    records, skipped = process_files(files, on_progress=on_progress, on_error=on_error)

    if skipped:
        print(f"\n{len(skipped)} עמודים דולגו (לא זוהו כתעודות פינוי):")
        for item in skipped:
            print(f"  {item.get('source_file', '')}: {item.get('notes', '')}")

    output_path = config.OUTPUT_DIR / config.OUTPUT_FILENAME
    write_records(records, output_path)
    print(f"\nנשמר: {output_path} ({len(records)} תעודות)")


if __name__ == "__main__":
    main()
