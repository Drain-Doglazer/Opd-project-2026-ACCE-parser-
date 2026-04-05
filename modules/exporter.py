import openpyxl
from openpyxl.styles import Font, Alignment
from typing import List
import logging

from .schema import ACCERecord

logger = logging.getLogger(__name__)

SHEET_NAME_MAP = {
    'CP': 'PUMPS',  # Centrifugal Pumps
    'VT': 'VESSELS',  # Vertical Tanks / Vessels (Reactors, Separators)
    'HT': 'HORIZ TANKS',  # Horizontal Tanks
    'HE': 'HEAT EXCHANGERS',  # Heat Exchangers
    'FN': 'FANS',  # Fans / Blowers
    'GC': 'COMPRESSORS',  # Gas Compressors
    'FU': 'FIRED HEATERS',  # Furnaces
    'STB': 'BOILERS',  # Boilers
    'CE': 'CRANES',  # Cranes / Hoists
    'FLR': 'FLARES',  # Flares
    'STK': 'STACKS',  # Stacks
    'DC': 'DUST COLLECTORS',  # Dust Collectors
}

BASE_HEADERS = [
    'ACTION',
    'USER TAG NUMBER',
    'DESCRIPTION',
    'PARENT AREA',
    'ITEM SYMBOL',
    'REFERENCE ID'
]


def _create_contents_sheet(wb: openpyxl.Workbook):
    ws = wb.create_sheet(title="Contents", index=0)

    cell = ws['A1']
    cell.value = "Spreadsheet Contents"
    cell.font = Font(bold=True, size=14)

    ws['A3'] = "Click the links below to navigate to the worksheets."

    row = 5
    for sheet_name in wb.sheetnames:
        if sheet_name == "Contents":
            continue

        cell = ws[f'A{row}']
        cell.value = sheet_name
        cell.hyperlink = f"#'{sheet_name}'!A1"
        cell.style = "Hyperlink"
        cell.font = Font(color="0000FF", underline="single")
        row += 1


def _create_areas_sheet(wb: openpyxl.Workbook, records: List[ACCERecord]):
    ws = wb.create_sheet(title="AREAS")

    headers = ['ACTION', 'AREA NAME', 'REPORT GROUP', 'AREA TYPE', 'LENGTH', 'WIDTH', 'HEIGHT']
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header).font = Font(bold=True)

    unique_areas = set()
    for rec in records:
        area = rec.parent_area.strip() if rec.parent_area else "MAIN_AREA"
        unique_areas.add(area)

    row_idx = 2
    for area_name in sorted(unique_areas):
        ws.cell(row=row_idx, column=1, value="NEW")  # ACTION
        ws.cell(row=row_idx, column=2, value=area_name)  # AREA NAME
        ws.cell(row=row_idx, column=3, value="Main")  # REPORT GROUP
        ws.cell(row=row_idx, column=4, value="GRADE")  # AREA TYPE
        row_idx += 1


def _get_sheet_name(symbol: str) -> str:
    return SHEET_NAME_MAP.get(symbol, f"EQUIP_{symbol}")


def _write_equipment_data(ws, records: List[ACCERecord], symbol_filter: str):
    filtered_records = [
        r for r in records
        if r.acce_item_symbol == symbol_filter and r.is_valid
    ]

    if not filtered_records:
        return

    for col, header in enumerate(BASE_HEADERS, start=1):
        ws.cell(row=1, column=col, value=header).font = Font(bold=True, color="000080")

    extra_headers = []

    if symbol_filter == 'CP':  # Насосы
        extra_headers = ['FLOW_RATE_M3H', 'HEAD_M', 'POWER_KW', 'DESIGN_TEMP_C', 'DESIGN_PRESS_BAR']
    elif symbol_filter in ['VT', 'HT']:  # Емкости
        extra_headers = ['VOLUME_M3', 'DIAMETER_M', 'LENGTH_M', 'DESIGN_TEMP_C', 'DESIGN_PRESS_BAR', 'MATERIAL']
    elif symbol_filter == 'HE':  # Теплообменники
        extra_headers = ['AREA_M2', 'DESIGN_TEMP_SHELL_C', 'DESIGN_PRESS_SHELL_BAR']
    elif symbol_filter == 'FN':  # Вентиляторы
        extra_headers = ['FLOW_RATE_M3H', 'PRESSURE_PA', 'POWER_KW']
    elif symbol_filter == 'GC':  # Компрессоры
        extra_headers = ['FLOW_RATE_NM3H', 'SUCTION_PRESS_BAR', 'DISCHARGE_PRESS_BAR', 'POWER_KW']
    elif symbol_filter == 'FU':  # Печи
        extra_headers = ['HEAT_DUTY_KW', 'DESIGN_TEMP_C']
    elif symbol_filter == 'STB':  # Котлы
        extra_headers = ['CAPACITY_TPH', 'PRESSURE_BAR', 'TEMP_C']
    else:
        extra_headers = ['DESIGN_TEMP_C', 'DESIGN_PRESS_BAR', 'WEIGHT_KG']

    start_extra_col = len(BASE_HEADERS) + 1
    for i, h in enumerate(extra_headers):
        ws.cell(row=1, column=start_extra_col + i, value=h).font = Font(bold=True, color="006400")

    row_idx = 2
    for rec in filtered_records:
        ws.cell(row=row_idx, column=1, value=rec.action)
        ws.cell(row=row_idx, column=2, value=rec.user_tag)

        desc = rec.description_ru[:250] if rec.description_ru else "No Description"
        ws.cell(row=row_idx, column=3, value=desc)

        area = rec.parent_area.strip() if rec.parent_area else "MAIN_AREA"
        ws.cell(row=row_idx, column=4, value=area)

        ws.cell(row=row_idx, column=5, value=rec.acce_item_symbol)
        ws.cell(row=row_idx, column=6, value=rec.raw_tag)

        col_idx = start_extra_col

        press_bar = rec.pressure * 10 if rec.pressure else None

        if symbol_filter == 'CP':
            ws.cell(row=row_idx, column=col_idx, value=rec.flow_rate)  # Flow
            ws.cell(row=row_idx, column=col_idx + 1, value=None)  # Head (нет в схеме пока)
            ws.cell(row=row_idx, column=col_idx + 2, value=rec.motor_power_kw)  # Power
            ws.cell(row=row_idx, column=col_idx + 3, value=rec.design_temperature)  # Temp
            ws.cell(row=row_idx, column=col_idx + 4, value=press_bar)  # Pressure

        elif symbol_filter in ['VT', 'HT']:
            ws.cell(row=row_idx, column=col_idx, value=rec.volume)
            ws.cell(row=row_idx, column=col_idx + 1, value=rec.diameter_m)
            ws.cell(row=row_idx, column=col_idx + 2, value=rec.length_m)
            ws.cell(row=row_idx, column=col_idx + 3, value=rec.design_temperature)
            ws.cell(row=row_idx, column=col_idx + 4, value=press_bar)
            ws.cell(row=row_idx, column=col_idx + 5, value=rec.material)

        elif symbol_filter == 'HE':
            ws.cell(row=row_idx, column=col_idx, value=rec.heat_transfer_area_m2)
            ws.cell(row=row_idx, column=col_idx + 1, value=rec.design_temperature)
            ws.cell(row=row_idx, column=col_idx + 2, value=press_bar)

        elif symbol_filter == 'FN':
            ws.cell(row=row_idx, column=col_idx, value=rec.flow_rate)
            ws.cell(row=row_idx, column=col_idx + 1, value=rec.pressure)
            ws.cell(row=row_idx, column=col_idx + 2, value=rec.motor_power_kw)

        elif symbol_filter == 'GC':
            ws.cell(row=row_idx, column=col_idx, value=rec.flow_rate)
            ws.cell(row=row_idx, column=col_idx + 1, value=rec.pressure)
            ws.cell(row=row_idx, column=col_idx + 2,
                    value=rec.pressure * 1.5 if rec.pressure else None)
            ws.cell(row=row_idx, column=col_idx + 3, value=rec.motor_power_kw)

        elif symbol_filter == 'FU':
            duty_kw = rec.heat_duty_gcalh * 1163 if rec.heat_duty_gcalh else None
            ws.cell(row=row_idx, column=col_idx, value=duty_kw)
            ws.cell(row=row_idx, column=col_idx + 1, value=rec.design_temperature)

        elif symbol_filter == 'STB':
            duty_kw = rec.heat_duty_gcalh * 1163 if rec.heat_duty_gcalh else None
            ws.cell(row=row_idx, column=col_idx, value=duty_kw)
            ws.cell(row=row_idx, column=col_idx + 1, value=press_bar)
            ws.cell(row=row_idx, column=col_idx + 2, value=rec.design_temperature)

        else:
            ws.cell(row=row_idx, column=col_idx, value=rec.design_temperature)
            ws.cell(row=row_idx, column=col_idx + 1, value=press_bar)
            ws.cell(row=row_idx, column=col_idx + 2, value=rec.weight_unit_kg)

        row_idx += 1


def export_to_excel(records: List[ACCERecord], output_path: str):
    logger.info(f"Starting export for {len(records)} records to {output_path}")

    wb = openpyxl.Workbook()

    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    symbols_present = set(
        r.acce_item_symbol
        for r in records
        if r.is_valid and r.acce_item_symbol
    )

    _create_areas_sheet(wb, records)

    for symbol in sorted(symbols_present):
        sheet_name = _get_sheet_name(symbol)
        try:
            ws = wb.create_sheet(title=sheet_name)
            _write_equipment_data(ws, records, symbol)
        except ValueError as e:
            logger.warning(f"Could not create sheet {sheet_name}: {e}")

    _create_contents_sheet(wb)

    wb.save(output_path)
    logger.info(f"Export completed successfully: {output_path}")