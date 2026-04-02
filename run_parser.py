"""
run_parser.py — CLI entry point for the ACCE equipment list parser.

Usage:
    python run_parser.py path/to/file.xlsx
    python run_parser.py path/to/file.docx

Format is detected automatically from the file extension.
"""

import sys
import os
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _tech_summary_xlsx(r) -> str:
    parts = []
    if r.flow_rate:       parts.append(f"Q={r.flow_rate}{r.flow_rate_unit}")
    if r.capacity_kw:     parts.append(f"Q={r.capacity_kw}kW")
    if r.pressure:        parts.append(f"P={r.pressure}{r.pressure_unit}")
    if r.volume:          parts.append(f"V={r.volume}{r.volume_unit}")
    if r.dn_mm:           parts.append(f"DN={r.dn_mm}")
    if r.diameter_m:      parts.append(f"Φ={r.diameter_m}m×{r.length_m}m")
    if r.lift_capacity_t: parts.append(f"T={r.lift_capacity_t}t")
    return "  ".join(parts) if parts else "—"


def _tech_summary_docx(r) -> str:
    parts = []
    if r.flow_rate:            parts.append(f"Q={r.flow_rate}{r.flow_rate_unit}")
    if r.pressure:             parts.append(f"P={r.pressure}MPa")
    if r.volume:               parts.append(f"V={r.volume}m3")
    if r.design_temperature:   parts.append(f"Td={r.design_temperature}°C")
    if r.operating_temperature:parts.append(f"To={r.operating_temperature}°C")
    if r.diameter_m:           parts.append(f"d={r.diameter_m}m")
    if r.motor_power_kw:       parts.append(f"N={r.motor_power_kw}kW")
    return "  ".join(parts) if parts else "—"


def run_xlsx(filepath: str):
    from acce_tools.modules.parser_xlsx import parse_xlsx
    records, warnings = parse_xlsx(filepath)

    total   = len(records)
    valid   = sum(1 for r in records if r.is_valid)
    no_type = sum(1 for r in records if not r.acce_item_symbol)

    type_counts: dict[str, int] = {}
    for r in records:
        sym = r.acce_item_symbol or "UNKNOWN"
        type_counts[sym] = type_counts.get(sym, 0) + 1

    print(f"\nSUMMARY")
    print(f"  Total records     : {total}")
    print(f"  Without warnings  : {valid}")
    print(f"  With warnings     : {total - valid}")
    print(f"  Type not resolved : {no_type}")
    print(f"  Skipped rows      : {len(warnings)}")
    print(f"\n  Icarus types:")
    LABELS = {
        "CP":"Centrifugal pumps","FN":"Fans","GC":"Compressors",
        "VT":"Vessels/tanks","HT":"Horizontal vessels","HE":"Heat exchangers",
        "FU":"Furnaces/incinerators","STB":"Boilers","CE":"Cranes","HO":"Hoists",
        "UNKNOWN":"Not resolved"
    }
    for sym, cnt in sorted(type_counts.items()):
        print(f"    {sym:<8} {cnt:>3} pcs  — {LABELS.get(sym, sym)}")

    print(f"\n{'─'*64}")
    print(f"RECORDS")
    print(f"{'─'*64}")
    print(f"  {'Tag':<20} {'Sym':<6} {'Tech specs'}")
    print(f"  {'─'*20} {'─'*6} {'─'*30}")
    for r in records:
        mark = " ⚠" if r.warnings else ""
        print(f"  {r.raw_tag:<20} {r.acce_item_symbol:<6} "
              f"{_tech_summary_xlsx(r)}{mark}")

    _print_issues(records, warnings)


def run_docx(filepath: str):
    from acce_tools.modules.parser_docx import parse_docx
    records, warnings = parse_docx(filepath)

    total   = len(records)
    no_type = sum(1 for r in records if not r.acce_item_symbol)

    type_counts: dict[str, int] = {}
    for r in records:
        sym = r.acce_item_symbol or "UNKNOWN"
        type_counts[sym] = type_counts.get(sym, 0) + 1

    print(f"\nSUMMARY")
    print(f"  Total records    : {total}")
    print(f"  Type not resolved: {no_type}")
    print(f"  Warnings         : {len(warnings)}")
    print(f"\n  Icarus types:")
    LABELS = {
        "CP":"Pumps","FN":"Fans/blowers","GC":"Compressors",
        "VT":"Vessels/reactors/separators","HE":"Heat exchangers",
        "FLR":"Flares","STK":"Stacks/pipes","UNKNOWN":"Not resolved"
    }
    for sym, cnt in sorted(type_counts.items()):
        print(f"    {sym:<8} {cnt:>3} pcs  — {LABELS.get(sym, sym)}")

    print(f"\n{'─'*72}")
    print(f"RECORDS")
    print(f"{'─'*72}")
    print(f"  {'Tag':<20} {'Sym':<6} {'Material':<12} {'Tech specs'}")
    print(f"  {'─'*20} {'─'*6} {'─'*12} {'─'*30}")
    for r in records:
        mark = " ⚠" if r.warnings else ""
        print(f"  {r.user_tag:<20} {r.acce_item_symbol:<6} "
              f"{r.material[:11]:<12} {_tech_summary_docx(r)}{mark}")

    _print_issues(records, warnings)


def _print_issues(records, warnings):
    problem = [r for r in records if r.warnings]
    unknown = [r for r in records if not r.acce_item_symbol]

    if warnings:
        print(f"\n{'─'*72}")
        print(f"SKIPPED ROWS ({len(warnings)})")
        for w in warnings:
            print(f"  Row {w.get('row','?'):<4}  {w.get('tag','—'):<15}  {w['reason']}")

    if problem:
        print(f"\n{'─'*72}")
        print(f"RECORDS WITH WARNINGS ({len(problem)})")
        for r in problem:
            print(f"  {r.raw_tag}")
            for w in r.warnings:
                print(f"    ⚠  {w}")

    if unknown:
        print(f"\n{'─'*72}")
        print(f"TYPE NOT RESOLVED — {len(unknown)} items (require review)")
        for r in unknown:
            print(f"  {r.raw_tag:<15}  {r.description_ru[:40]}")

    print(f"\n{'='*72}")
    has_issues = warnings or problem or unknown
    print(f"Result: {'REQUIRES REVIEW ⚠' if has_issues else 'OK ✓'}")
    print()


def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python run_parser.py path/to/file.xlsx")
        print("  python run_parser.py path/to/file.docx")
        sys.exit(1)

    filepath = sys.argv[1]

    if not os.path.exists(filepath):
        print(f"Error: file not found — {filepath}")
        sys.exit(1)

    ext = os.path.splitext(filepath)[1].lower()

    print(f"\nFile  : {filepath}")
    print(f"Format: {ext.upper()}")
    print("=" * 72)

    if ext == ".xlsx":
        run_xlsx(filepath)
    elif ext == ".docx":
        run_docx(filepath)
    else:
        print(f"Error: unsupported format '{ext}'")
        print("Supported: .xlsx, .docx")
        sys.exit(1)


if __name__ == "__main__":
    main()
