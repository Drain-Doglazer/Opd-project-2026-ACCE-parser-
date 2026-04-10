"""
run_parser.py — CLI entry point for the ACCE equipment list parser.

Usage:
    python run_parser.py path/to/file.xlsx
    python run_parser.py path/to/file.docx
    python run_parser.py path/to/file.pdf

Format is detected automatically from the file extension.
"""

import sys
import os
import re

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
sys.path.insert(0, os.path.join(_HERE, "acce_tools"))


# ── Сводки тех. характеристик по типу файла ───────────────────────

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
    if r.flow_rate:             parts.append(f"Q={r.flow_rate}{r.flow_rate_unit}")
    if r.pressure:              parts.append(f"P={r.pressure}MPa")
    if r.volume:                parts.append(f"V={r.volume}m3")
    if r.design_temperature:    parts.append(f"Td={r.design_temperature}°C")
    if r.operating_temperature: parts.append(f"To={r.operating_temperature}°C")
    if r.diameter_m:            parts.append(f"d={r.diameter_m}m")
    if r.motor_power_kw:        parts.append(f"N={r.motor_power_kw}kW")
    return "  ".join(parts) if parts else "—"


def _tech_summary_pdf(r) -> str:
    """
    Сводка тех. характеристик для PDF.
    Работает для обоих форматов:
    - Equipment List: те же поля что xlsx (Q, P, V, DN, Φ×L)
    - Регламент 5.1/5.2 (колонны TW): P, Td, D, L
    - Регламент 5.3 (теплообменники HE): P, Td, площадь, нагрузка
    """
    parts = []
    # Поля Equipment List
    if r.flow_rate:             parts.append(f"Q={r.flow_rate}{r.flow_rate_unit}")
    if r.capacity_kw:           parts.append(f"Q={r.capacity_kw}kW")
    if r.pressure:
        unit = getattr(r, 'pressure_unit', 'MPa') or 'MPa'
        parts.append(f"P={r.pressure}{unit}")
    if r.volume:                parts.append(f"V={r.volume}m3")
    if r.dn_mm:                 parts.append(f"DN={r.dn_mm}")
    if r.lift_capacity_t:       parts.append(f"T={r.lift_capacity_t}t")
    # Поля регламента
    if r.design_temperature:    parts.append(f"Td={r.design_temperature}°C")
    if r.diameter_m:            parts.append(f"D={r.diameter_m}m")
    if r.length_m and not r.diameter_m:
        parts.append(f"L={r.length_m}m")
    if r.heat_transfer_area_m2: parts.append(f"F={r.heat_transfer_area_m2}m²")
    if r.heat_duty_gcalh:       parts.append(f"Q={r.heat_duty_gcalh}Gcal/h")
    return "  ".join(parts) if parts else "—"


# ── Вывод для каждого формата ─────────────────────────────────────

def run_xlsx(filepath: str):
    from modules.parser_xlsx import parse_xlsx
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

    print(f"\n{'─'*72}")
    print(f"RECORDS")
    print(f"{'─'*72}")
    print(f"  {'Tag':<20} {'Sym':<6} {'Tech specs'}")
    print(f"  {'─'*20} {'─'*6} {'─'*30}")
    for r in records:
        mark = " ⚠" if r.warnings else ""
        print(f"  {r.raw_tag:<20} {r.acce_item_symbol:<6} "
              f"{_tech_summary_xlsx(r)}{mark}")

    _print_issues(records, warnings)


def run_docx(filepath: str):
    from modules.parser_docx import parse_docx
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


def run_pdf(filepath: str):
    from modules.parser_pdf import parse_pdf
    records, warnings = parse_pdf(filepath)

    total   = len(records)
    no_type = sum(1 for r in records if not r.acce_item_symbol)

    # Группировка по разделам (только для регламента)
    sections: dict[str, int] = {}
    for r in records:
        sections[r.source_sheet] = sections.get(r.source_sheet, 0) + 1

    type_counts: dict[str, int] = {}
    for r in records:
        sym = r.acce_item_symbol or "UNKNOWN"
        type_counts[sym] = type_counts.get(sym, 0) + 1

    print(f"\nSUMMARY")
    print(f"  Total records    : {total}")
    print(f"  Type not resolved: {no_type}")
    print(f"  Warnings         : {len(warnings)}")

    # Показываем разделы только если их больше одного
    if len(sections) > 1:
        print(f"\n  By section:")
        SEC_LABELS = {
            "Equipment List": "Equipment List",
            "5.1": "5.1  Rectification columns (TW)",
            "5.2": "5.2  Column equipment (TW)",
            "5.3": "5.3  Heat exchangers (HE)",
        }
        for sec, cnt in sections.items():
            print(f"    {SEC_LABELS.get(sec, sec):<38} {cnt:>3} pcs")

    print(f"\n  Icarus types:")
    LABELS = {
        "TW":"Towers (single/double diameter)",
        "HE":"Heat exchangers",
        "CP":"Centrifugal pumps","FN":"Fans","GC":"Compressors",
        "VT":"Vessels/tanks","FU":"Furnaces","STB":"Boilers",
        "HO":"Hoists","CE":"Cranes","FLR":"Flares","STK":"Stacks",
        "UNKNOWN":"Not resolved"
    }
    for sym, cnt in sorted(type_counts.items()):
        print(f"    {sym:<8} {cnt:>3} pcs  — {LABELS.get(sym, sym)}")

    print(f"\n{'─'*80}")
    print(f"RECORDS")
    print(f"{'─'*80}")
    print(f"  {'Section':<6} {'Tag':<12} {'user_tag':<12} {'Sym':<6} {'Sub':<14} {'Tech specs'}")
    print(f"  {'─'*6} {'─'*12} {'─'*12} {'─'*6} {'─'*14} {'─'*28}")
    for r in records:
        mark = " ⚠" if r.warnings else ""
        print(f"  {r.source_sheet:<6} {r.raw_tag:<12} {r.user_tag:<12} "
              f"{r.acce_item_symbol:<6} {r.acce_item_type:<14} "
              f"{_tech_summary_pdf(r)}{mark}")

    _print_issues(records, warnings)


# ── Общий вывод проблем ───────────────────────────────────────────

def _print_issues(records, warnings):
    problem = [r for r in records if r.warnings]
    unknown = [r for r in records if not r.acce_item_symbol]

    if warnings:
        print(f"\n{'─'*72}")
        print(f"SKIPPED / WARNINGS ({len(warnings)})")
        for w in warnings:
            tag = w.get('tag', w.get('cell', '—'))
            loc = f"Row {w['row']}" if 'row' in w else f"Page {w.get('page','?')}"
            print(f"  {loc:<8}  {tag:<15}  {w['reason']}")

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


# ── Точка входа ───────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python run_parser.py path/to/file.xlsx")
        print("  python run_parser.py path/to/file.docx")
        print("  python run_parser.py path/to/file.pdf")
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
    elif ext == ".pdf":
        run_pdf(filepath)
    else:
        print(f"Error: unsupported format '{ext}'")
        print("Supported: .xlsx, .docx, .pdf")
        sys.exit(1)


if __name__ == "__main__":
    main()