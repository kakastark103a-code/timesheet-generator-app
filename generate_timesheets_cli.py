#!/usr/bin/env python3
"""
CLI Tool for generating monthly timesheet Excel files per domain with member & comment overrides.

Usage:
  python generate_timesheets_cli.py --month 2026-05 --all --notes-file notes.txt
  python generate_timesheets_cli.py --month 2026-06 --domains cbg ebg
"""

import argparse
import os
from datetime import datetime
from timesheet_generator import generate_domain_timesheet, find_template_path, get_singapore_holidays, DOMAIN_FILE_MAP, DOMAIN_NAMES

def main():
    parser = argparse.ArgumentParser(description="Generate monthly domain timesheets preserving Excel formulas, excluding weekends, and handling member overrides.")
    parser.add_argument("--month", required=True, help="Target month in YYYY-MM format (e.g. 2026-05)")
    parser.add_argument("--domains", nargs="+", help="Specific domain keys (cbg, ebg, identity, provisioning, rwfm)")
    parser.add_argument("--all", action="store_true", help="Generate for all available domains")
    parser.add_argument("--holidays", nargs="*", default=[], help="List of additional custom public holiday dates in YYYY-MM-DD format")
    parser.add_argument("--no-sg-holidays", action="store_true", help="Disable auto-detection of Singapore Public Holidays")
    parser.add_argument("--notes-file", help="Path to text/TSV file containing member comment notes/overrides")
    parser.add_argument("--outdir", default="generated_timesheets", help="Output directory for generated Excel files")
    parser.add_argument("--templatedir", default="timesheets_extracted", help="Template directory containing template files")

    args = parser.parse_args()

    try:
        dt = datetime.strptime(args.month, "%Y-%m")
        year, month = dt.year, dt.month
    except ValueError:
        print("Error: --month must be in YYYY-MM format, e.g. 2026-05")
        return

    if args.all or not args.domains:
        target_keys = list(DOMAIN_FILE_MAP.keys())
    else:
        target_keys = args.domains

    include_sg = not args.no_sg_holidays
    sg_hols = get_singapore_holidays(year, month) if include_sg else {}

    notes_text = ""
    if args.notes_file and os.path.exists(args.notes_file):
        with open(args.notes_file, 'r', encoding='utf-8') as f:
            notes_text = f.read()
        print(f"📖 Loaded comment notes file: {args.notes_file}")

    month_suffix = datetime(year, month, 1).strftime("%b%Y")
    print(f"==================================================")
    print(f"Generating Timesheets for: {datetime(year, month, 1).strftime('%B %Y')}")
    print(f"Auto Singapore Public Holidays: {'ENABLED' if include_sg else 'DISABLED'}")
    if sg_hols:
        for d, name in sg_hols.items():
            print(f"  🇸🇬 {d}: {name}")
    print(f"Custom Holidays: {args.holidays if args.holidays else 'None'}")
    print(f"Target Domains: {', '.join(target_keys)}")
    print(f"==================================================")

    generated = 0
    for key in target_keys:
        template_path = find_template_path(key, args.templatedir)
        if not template_path:
            print(f"⚠️ Warning: Could not find template for domain '{key}' in {args.templatedir}")
            continue

        domain_name = DOMAIN_NAMES.get(key, key.upper())
        out_filename = f"FPT_{domain_name}_Timesheet_{month_suffix}.xlsx"
        out_filepath = os.path.join(args.outdir, out_filename)

        print(f"Processing domain '{domain_name}'...", end="", flush=True)
        generate_domain_timesheet(
            template_path=template_path,
            year=year,
            month=month,
            output_path=out_filepath,
            public_holidays=args.holidays,
            include_sg_holidays=include_sg,
            comment_notes=notes_text
        )
        print(f" ✅ Saved to {out_filepath}")
        generated += 1

    print(f"\n🎉 Finished generating {generated} timesheet(s) in directory '{args.outdir}'.")

if __name__ == '__main__':
    main()
