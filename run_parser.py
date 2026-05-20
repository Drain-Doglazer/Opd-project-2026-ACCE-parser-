"""
run_parser.py — CLI entry point for the ACCE equipment list parser.

Usage:
    python run_parser.py <input_file> [--output <output_file>]
                         [--template <template.xlsx>] [--area <area_name>]

Format is detected from the input file extension (.xlsx, .docx, .pdf).
Parser output format is detected from --output extension (.xlsx, .csv, .json).
Default parser output: same directory, same base name, _parsed.xlsx suffix.

When --template is supplied, an additional ACCE-ready import file is produced
alongside the parser output.
"""

import argparse
import logging
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from acce_tools.modules.parser_xlsx import parse as parse_xlsx
from acce_tools.modules.parser_docx import parse as parse_docx
from acce_tools.modules.parser_pdf  import parse as parse_pdf
from acce_tools.modules.classifier  import classify
from acce_tools.modules.writer      import write
from acce_tools.validator.Validator import validate_all

_PARSERS = {
    ".xlsx": parse_xlsx,
    ".docx": parse_docx,
    ".pdf":  parse_pdf,
}


def _default_parsed_path(input_path: str) -> str:
    """Derive a safe parser-output path (avoids overwriting the source)."""
    base, ext = os.path.splitext(input_path)
    if ext.lower() == ".xlsx":
        return base + "_parsed.xlsx"
    return base + ".xlsx"


def _default_acce_path(parsed_path: str) -> str:
    """Derive the ACCE export path from the parser output path."""
    base = os.path.splitext(parsed_path)[0]
    return base.removesuffix("_parsed") + "_acce.xlsx"


def run_pipeline(
    input_path: str,
    template_path: str | None = None,
    area: str = "Main Area",
    output_path: str | None = None,
    parsed_path: str | None = None,
) -> dict:
    """
    Runs the full parse → classify → validate → write → export pipeline.
    Returns a dict with stats and file paths.
    Can be called from Streamlit or any other UI without argparse.
    """
    from pathlib import Path

    input_path = str(input_path)
    ext = Path(input_path).suffix.lower()

    if ext not in _PARSERS:
        raise ValueError(f"Unsupported file type: {ext}")

    stem   = Path(input_path).stem
    parent = Path(input_path).parent
    if parsed_path is None:
        parsed_path = str(parent / f"{stem}_parsed.xlsx")
    if output_path is None:
        output_path = str(parent / f"{stem}_acce.xlsx")

    os.makedirs(os.path.dirname(os.path.abspath(parsed_path)), exist_ok=True)

    records = _PARSERS[ext](input_path)
    records = [classify(r) for r in records]
    records = validate_all(records)
    write(records, parsed_path)

    total    = len(records)
    valid    = sum(1 for r in records if r.is_valid)
    invalid  = total - valid
    warnings = sum(1 for r in records if r.warnings)

    result = {
        "total":       total,
        "valid":       valid,
        "invalid":     invalid,
        "warnings":    warnings,
        "parsed_path": parsed_path,
        "output_path": None,
        "summary":     None,
    }

    if template_path:
        from acce_tools.modules.exporter import export
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        summary = export(
            input_path=parsed_path,
            template_path=template_path,
            output_path=output_path,
            default_area=area,
        )
        result["output_path"] = output_path
        result["summary"]     = summary

    return result


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Parse equipment lists and optionally export to ACCE format.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "input",
        metavar="INPUT_FILE",
        help="Source file to parse (.xlsx, .docx, .pdf)",
    )
    parser.add_argument(
        "--output", "-o",
        metavar="OUTPUT_FILE",
        help=(
            "ACCE-ready export file (requires --template). "
            "Default: <input>_acce.xlsx"
        ),
    )
    parser.add_argument(
        "--template", "-t",
        metavar="TEMPLATE_XLSX",
        help="ACCE master template .xlsx for spreadsheet import",
    )
    parser.add_argument(
        "--area", "-a",
        metavar="AREA_NAME",
        default="Main Area",
        help="Default PARENT_AREA when the source file does not specify one",
    )
    parser.add_argument(
        "--parsed",
        metavar="PARSED_FILE",
        help=(
            "Parser intermediate output file (.xlsx, .csv, .json). "
            "Default: <input>_parsed.xlsx"
        ),
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable debug logging",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.WARNING,
        format="%(levelname)s %(name)s: %(message)s",
    )

    input_path = args.input

    if not os.path.exists(input_path):
        print(f"Error: file not found — {input_path}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in _PARSERS:
        print(
            f"Error: unsupported input format '{ext}'. "
            f"Supported: {', '.join(_PARSERS)}",
            file=sys.stderr,
        )
        sys.exit(1)

    if args.template and not os.path.exists(args.template):
        print(f"Error: template not found — {args.template}", file=sys.stderr)
        sys.exit(1)

    parsed_path = args.parsed or _default_parsed_path(input_path)
    acce_path   = args.output or _default_acce_path(parsed_path)

    result = run_pipeline(
        input_path=input_path,
        template_path=args.template,
        area=args.area or "Main Area",
        output_path=acce_path,
        parsed_path=parsed_path,
    )

    print(f"Parsed:   {result['total']} records from {input_path}")
    print(f"Valid:    {result['valid']}  |  Invalid: {result['invalid']}  |  Warnings: {result['warnings']}")
    print(f"Written:  {result['parsed_path']}")

    summary = result.get("summary")
    if summary:
        print()
        print(f"ACCE export: {result['output_path']}")
        print(f"  Exported:       {summary.exported} records")
        print(f"  Skipped:        {summary.skipped_invalid} (invalid)")
        print(f"  Unknown symbol: {summary.unknown_symbol}")
        if summary.used_default_type:
            print(f"  Default type:   {summary.used_default_type} (type was UNKNOWN)")
        if summary.by_sheet:
            print("  By sheet:")
            for sheet, count in sorted(summary.by_sheet.items()):
                print(f"    {sheet}: {count}")


if __name__ == "__main__":
    main()
