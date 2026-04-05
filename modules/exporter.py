from dataclasses import dataclass

import openpyxl
from openpyxl.styles import Font
from typing import List, Dict
import logging
import os

from .schema import ACCERecord

logger = logging.getLogger(__name__)

SHEET_NAME_MAP = {
    'CP': 'CP',  # Pumps
    'FN': 'FN',  # Fans
    'GC': 'GC',  # Compressors
    'VT': 'VT',  # Vessels/Tanks
    'HT': 'HT',  # Horizontal Tanks (если нужно отдельно, иначе в VT)
    'HE': 'HE',  # Heat Exchangers
    'FU': 'FU',  # Furnaces
    'STB': 'STB',  # Boilers
    'CE': 'CE',  # Cranes
    'HO': 'HO',  # Hoists
    'FLR': 'FLR',  # Flares
    'STK': 'STK',  # Stacks
    'DC': 'DC',  # Dust Collectors
}


@dataclass
class ExportResult:
    exported_count: int
    skipped_count: int
    by_type: Dict[str, int]
    output_path: str
    warnings: List[str]


def _create_dummy_template(template_path: str):
    if os.path.exists(template_path):
        return

    wb = openpyxl.Workbook()

    ws_areas = wb.create_sheet("AREAS")
    headers_areas = ['ACTION', 'AREA_NAME', 'REPORT_GROUP', 'AREA_TYPE']
    for col, h in enumerate(headers_areas, 1):
        ws_areas.cell(1, col, value=h).font = Font(bold=True)
    ws_areas['A2'] = "LAST ROW"

    for symbol in SHEET_NAME_MAP.keys():
        ws = wb.create_sheet(symbol)
        headers_eq = ['ACTION', 'ITEM_SYMBOL', 'ITEM_TYPE', 'USER_TAG', 'PARENT_AREA', 'DESCRIPTION']
        for col, h in enumerate(headers_eq, 1):
            ws.cell(11, col, value=h).font = Font(bold=True, color="000080")

        if symbol == 'CP':
            extra = ['Design temperature', 'Design gauge pressure', 'Driver power']
        elif symbol == 'VT':
            extra = ['Design gauge pressure', 'Design temperature', 'Liquid volume', 'Vessel diameter',
                     'Shell material']
        elif symbol == 'FN':
            extra = ['Actual gas flow rate', 'Fan outlet gauge pressure', 'Driver power']
        elif symbol == 'STB':
            extra = ['Thermal duty', 'Design gauge pressure']
        elif symbol in ['FLR', 'STK']:
            extra = ['Stack height', 'Stack diameter']
        else:
            extra = ['Design temperature', 'Design gauge pressure']

        for i, h in enumerate(extra, 7):
            ws.cell(11, i, value=h).font = Font(bold=True, color="006400")

        ws['A100'] = "LAST ROW"

    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    wb.save(template_path)
    logger.info(f"Dummy template created at {template_path}")


def _ensure_unique_tags(records: List[ACCERecord]) -> List[str]:
    seen_tags = {}
    warnings = []

    tags_list = [rec.user_tag for rec in records if rec.is_valid and rec.user_tag]

    for i, tag in enumerate(tags_list):
        if not tag:
            continue

        base_tag = tag
        counter = 1

        while tag in seen_tags:
            suffix = f"-{chr(64 + counter)}"  # -A, -B...
            max_base_len = 20 - len(suffix)
            new_tag = base_tag[:max_base_len] + suffix
            tag = new_tag
            counter += 1

        seen_tags[tag] = True
        if tag != records[i].user_tag:
            warnings.append(f"Duplicate tag '{records[i].user_tag}' renamed to '{tag}'")
            records[i].user_tag = tag

    return warnings


def export_to_acce(
        records: List[ACCERecord],
        template_path: str,
        output_path: str,
        project_name: str = "TestProject"
) -> ExportResult:
    logger.info(f"Starting export using template: {template_path}")

    if not os.path.exists(template_path):
        _create_dummy_template(template_path)

    valid_records = [r for r in records if r.is_valid]
    skipped_count = len(records) - len(valid_records)

    if not valid_records:
        logger.warning("No valid records to export.")
        pass

    tag_warnings = _ensure_unique_tags(valid_records)

    areas_set = set()
    for rec in valid_records:
        area = rec.parent_area.strip() if rec.parent_area else "DEFAULT"
        areas_set.add(area)

    grouped_records: Dict[str, List[ACCERecord]] = {}
    for rec in valid_records:
        symbol = rec.acce_item_symbol
        if not symbol:
            continue
        if symbol not in grouped_records:
            grouped_records[symbol] = []
        grouped_records[symbol].append(rec)

    sorted_symbols = sorted(grouped_records.keys())

    try:
        wb = openpyxl.load_workbook(template_path)
    except Exception as e:
        raise FileNotFoundError(f"Cannot load template from {template_path}: {e}")

    if 'AREAS' in wb.sheetnames:
        ws_areas = wb['AREAS']
        last_row_idx = None
        for r in range(ws_areas.max_row, 1, -1):
            if str(ws_areas.cell(r, 1).value).upper() == "LAST ROW":
                last_row_idx = r
                break

        if not last_row_idx:
            last_row_idx = ws_areas.max_row + 1
            ws_areas.cell(last_row_idx, 1, value="LAST ROW")

        current_insert_row = last_row_idx
        for area_name in sorted(areas_set):
            ws_areas.insert_rows(current_insert_row, amount=1)
            ws_areas.cell(current_insert_row, 1, value="NEW")
            ws_areas.cell(current_insert_row, 2, value=area_name)
            ws_areas.cell(current_insert_row, 3, value=area_name)
            ws_areas.cell(current_insert_row, 4, value="PROCESS")
            current_insert_row += 1

    by_type_counts = {}

    for symbol in sorted_symbols:
        sheet_name = SHEET_NAME_MAP.get(symbol, symbol)

        if sheet_name not in wb.sheetnames:
            logger.warning(f"Sheet '{sheet_name}' not found in template. Skipping type {symbol}.")
            continue

        ws_eq = wb[sheet_name]

        last_row_idx = None
        for r in range(ws_eq.max_row, 10, -1):
            if str(ws_eq.cell(r, 1).value).upper() == "LAST ROW":
                last_row_idx = r
                break

        if not last_row_idx:
            last_row_idx = ws_eq.max_row + 1
            ws_eq.cell(last_row_idx, 1, value="LAST ROW")

        current_insert_row = last_row_idx
        records_for_type = grouped_records[symbol]

        for rec in records_for_type:
            ws_eq.insert_rows(current_insert_row, amount=1)
            ws_eq.cell(current_insert_row, 1, value=rec.action)
            ws_eq.cell(current_insert_row, 2, value=rec.acce_item_symbol)
            ws_eq.cell(current_insert_row, 3, value=rec.acce_item_type)
            ws_eq.cell(current_insert_row, 4, value=rec.user_tag)
            area_val = rec.parent_area.strip() if rec.parent_area else "DEFAULT"
            ws_eq.cell(current_insert_row, 5, value=area_val)
            desc = rec.description_ru[:60] if rec.description_ru else "No Desc"
            ws_eq.cell(current_insert_row, 6, value=desc)
            press_kpa = rec.pressure * 1000 if rec.pressure else None

            col_idx = 7

            if symbol == 'CP':
                ws_eq.cell(current_insert_row, col_idx, value=rec.design_temperature)
                ws_eq.cell(current_insert_row, col_idx + 1, value=press_kpa)
                ws_eq.cell(current_insert_row, col_idx + 2, value=rec.motor_power_kw)

            elif symbol == 'VT':
                ws_eq.cell(current_insert_row, col_idx, value=press_kpa)
                ws_eq.cell(current_insert_row, col_idx + 1, value=rec.design_temperature)
                ws_eq.cell(current_insert_row, col_idx + 2, value=rec.volume)
                ws_eq.cell(current_insert_row, col_idx + 3, value=rec.diameter_m)
                ws_eq.cell(current_insert_row, col_idx + 4, value=rec.material)

            elif symbol == 'FN':
                p_val = rec.pressure
                if rec.pressure_unit == 'МПа':
                    p_val = rec.pressure * 1_000_000
                ws_eq.cell(current_insert_row, col_idx, value=rec.flow_rate)
                ws_eq.cell(current_insert_row, col_idx + 1, value=p_val)
                ws_eq.cell(current_insert_row, col_idx + 2, value=rec.motor_power_kw)

            elif symbol == 'STB':
                ws_eq.cell(current_insert_row, col_idx, value=rec.capacity_kw)
                ws_eq.cell(current_insert_row, col_idx + 1, value=press_kpa)

            elif symbol in ['FLR', 'STK']:
                h_val = rec.height_m if rec.height_m else rec.length_m
                ws_eq.cell(current_insert_row, col_idx, value=h_val)
                ws_eq.cell(current_insert_row, col_idx + 1, value=rec.diameter_m)

            current_insert_row += 1

        by_type_counts[symbol] = len(records_for_type)

    ws_val = wb.create_sheet("VALIDATION")
    val_headers = ["TAG", "TYPE", "STATUS", "ERRORS", "WARNINGS"]
    for c, h in enumerate(val_headers, 1):
        ws_val.cell(1, c, value=h).font = Font(bold=True)

    row_val = 2
    for rec in records:
        status = "VALID" if rec.is_valid else "INVALID"
        errors_str = "; ".join(rec.errors) if rec.errors else ""
        warns_str = "; ".join(rec.warnings) if rec.warnings else ""

        ws_val.cell(row_val, 1, value=rec.user_tag or rec.raw_tag)
        ws_val.cell(row_val, 2, value=rec.acce_item_symbol)
        ws_val.cell(row_val, 3, value=status)
        ws_val.cell(row_val, 4, value=errors_str)
        ws_val.cell(row_val, 5, value=warns_str)
        row_val += 1

    if 'Contents' in wb.sheetnames:
        ws_cont = wb['Contents']
        wb.move_sheet(ws_cont, 0)
    else:
        ws_cont = wb.create_sheet("Contents", 0)

    ws_cont['A1'] = f"Project: {project_name}"
    ws_cont['A2'] = "Generated by ACCE Exporter"
    ws_cont['A4'] = "Sheet Name"
    ws_cont['B4'] = "Link"

    row_cont = 5
    for sheet_name in wb.sheetnames:
        if sheet_name == "Contents":
            continue
        ws_cont.cell(row_cont, 1, value=sheet_name)
        cell_link = ws_cont.cell(row_cont, 2)
        cell_link.value = f"Go to {sheet_name}"
        cell_link.hyperlink = f"#'{sheet_name}'!A1"
        cell_link.style = "Hyperlink"
        row_cont += 1

    wb.save(output_path)
    logger.info(f"Export saved to {output_path}")

    all_warnings = tag_warnings
    if skipped_count > 0:
        all_warnings.append(f"{skipped_count} records skipped due to validation errors.")

    return ExportResult(
        exported_count=len(valid_records),
        skipped_count=skipped_count,
        by_type=by_type_counts,
        output_path=output_path,
        warnings=all_warnings
    )