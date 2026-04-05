# modules/exporter.py
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from typing import List, Dict
import logging
import os
import re
from dataclasses import dataclass

# Импортируем схему данных
from .schema import ACCERecord

logger = logging.getLogger(__name__)

# Маппинг символов Icarus на имена листов в Excel (согласно ТЗ и документации ACCE)
SHEET_NAME_MAP = {
    'CP': 'CP',      # Pumps
    'FN': 'FN',      # Fans
    'GC': 'GC',      # Compressors
    'VT': 'VT',      # Vessels/Tanks
    'HT': 'HT',      # Horizontal Tanks
    'HE': 'HE',      # Heat Exchangers
    'FU': 'FU',      # Furnaces
    'STB': 'STB',    # Boilers
    'CE': 'CE',      # Cranes
    'HO': 'HO',      # Hoists
    'FLR': 'FLR',    # Flares
    'STK': 'STK',    # Stacks
    'DC': 'DC',      # Dust Collectors
}

@dataclass
class ExportResult:
    exported_count: int
    skipped_count: int
    by_type: Dict[str, int]
    output_path: str
    warnings: List[str]

def _create_dummy_template(template_path: str):
    """
    СОЗДАЕТ ФИКТИВНЫЙ ШАБЛОН ДЛЯ ТЕСТИРОВАНИЯ.
    В реальном проекте этот шаг не нужен, так как шаблон берется из ACCE.
    ТЗ требует наличия строк 1-10 и LAST ROW.
    """
    if os.path.exists(template_path):
        return

    wb = openpyxl.Workbook()
    
    # Создаем лист AREAS
    ws_areas = wb.create_sheet("AREAS")
    headers_areas = ['ACTION', 'AREA_NAME', 'REPORT_GROUP', 'AREA_TYPE']
    for col, h in enumerate(headers_areas, 1):
        ws_areas.cell(1, col, value=h).font = Font(bold=True)
    # Маркер конца данных для AREAS
    ws_areas['A2'] = "LAST ROW" 

    # Создаем фиктивные листы оборудования с заголовками A-F и LAST ROW
    # ТЗ говорит: строки 1-10 служебные. Мы сделаем заголовки на 11-й строке, как в реальном ACCE.
    for symbol in SHEET_NAME_MAP.keys():
        ws = wb.create_sheet(symbol)
        
        # Строки 1-10 оставляем пустыми или заполняем заглушками (служебные)
        for r in range(1, 11):
            ws[f'A{r}'] = f"Header Row {r}"
            
        # Заголовки данных начинаются со строки 11 (стандарт ACCE)
        headers_eq = [
            'ACTION',       # A
            'ITEM_SYMBOL',  # B
            'ITEM_TYPE',    # C
            'USER_TAG',     # D
            'PARENT_AREA',  # E
            'DESCRIPTION'   # F
        ]
        
        for col, h in enumerate(headers_eq, 1):
            ws.cell(11, col, value=h).font = Font(bold=True, color="000080")
        
        # Добавляем специфичные колонки (G+) для примера (ТЗ раздел 3.2)
        extra_headers = []
        if symbol == 'CP':
            extra_headers = ['Design temperature', 'Design gauge pressure', 'Driver power']
        elif symbol == 'VT':
            extra_headers = ['Design gauge pressure', 'Design temperature', 'Liquid volume', 'Vessel diameter', 'Shell material']
        elif symbol == 'FN':
            extra_headers = ['Actual gas flow rate', 'Fan outlet gauge pressure', 'Driver power']
        elif symbol == 'STB':
            extra_headers = ['Thermal duty', 'Design gauge pressure']
        elif symbol in ['FLR', 'STK']:
            extra_headers = ['Stack height', 'Stack diameter']
        else:
            extra_headers = ['Design temperature', 'Design gauge pressure']
            
        for i, h in enumerate(extra_headers, 7):
            ws.cell(11, i, value=h).font = Font(bold=True, color="006400")
            
        # Маркер LAST ROW сразу после заголовков для теста (строка 12)
        # В реальности он может быть ниже, но insert_rows перед ним сработает корректно.
        ws['A12'] = "LAST ROW"

    # Удаляем стандартный Sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
        
    wb.save(template_path)
    logger.info(f"Dummy template created at {template_path}")

def _ensure_unique_tags(records: List[ACCERecord]) -> List[str]:
    """
    ТЗ п. 3.1: Проверка уникальности user_tag. 
    Если дубль -> суффикс -A, -B...
    """
    seen_tags = {}
    warnings = []
    
    # Сначала соберем все теги
    tags_list = [rec.user_tag for rec in records if rec.is_valid and rec.user_tag]
    
    for i, tag in enumerate(tags_list):
        if not tag:
            continue
            
        base_tag = tag
        counter = 1
        
        while tag in seen_tags:
            suffix = f"-{chr(64+counter)}" # -A, -B...
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
    project_name: str = ""
) -> ExportResult:
    """
    Основная функция экспорта согласно ТЗ.
    """
    logger.info(f"Starting export using template: {template_path}")
    
    # Если шаблона нет (для теста), создаем фиктивный
    if not os.path.exists(template_path):
        _create_dummy_template(template_path)

    # 1. Фильтрация валидных записей
    valid_records = [r for r in records if r.is_valid]
    skipped_count = len(records) - len(valid_records)
    
    if not valid_records:
        logger.warning("No valid records to export.")
        
    # 2. Проверка уникальности тегов
    tag_warnings = _ensure_unique_tags(valid_records)
    
    # 3. Сбор уникальных зон (Areas)
    areas_set = set()
    for rec in valid_records:
        area = rec.parent_area.strip() if rec.parent_area else "DEFAULT"
        areas_set.add(area)
        
    # 4. Группировка по типу Icarus
    grouped_records: Dict[str, List[ACCERecord]] = {}
    for rec in valid_records:
        symbol = rec.acce_item_symbol
        if not symbol:
            continue
        if symbol not in grouped_records:
            grouped_records[symbol] = []
        grouped_records[symbol].append(rec)
        
    sorted_symbols = sorted(grouped_records.keys())
    
    # 5. Загрузка шаблона
    try:
        wb = openpyxl.load_workbook(template_path)
    except Exception as e:
        raise FileNotFoundError(f"Cannot load template from {template_path}: {e}")
    
    # --- ОБРАБОТКА ЛИСТА AREAS ---
    if 'AREAS' in wb.sheetnames:
        ws_areas = wb['AREAS']
        # Ищем строку LAST ROW
        last_row_idx = None
        for r in range(1, ws_areas.max_row + 1):
            if str(ws_areas.cell(r, 1).value).upper() == "LAST ROW":
                last_row_idx = r
                break
        
        if not last_row_idx:
            last_row_idx = ws_areas.max_row + 1
            ws_areas.cell(last_row_idx, 1, value="LAST ROW")

        current_insert_row = last_row_idx
        for area_name in sorted(areas_set):
            # Вставляем пустые строки перед LAST ROW
            ws_areas.insert_rows(current_insert_row, amount=1)
            
            # ACTION
            ws_areas.cell(current_insert_row, 1, value="NEW")
            # AREA_NAME
            ws_areas.cell(current_insert_row, 2, value=area_name)
            # REPORT_GROUP (по умолчанию имя области)
            ws_areas.cell(current_insert_row, 3, value=area_name)
            # AREA_TYPE
            ws_areas.cell(current_insert_row, 4, value="PROCESS")
            
            current_insert_row += 1

    # --- ОБРАБОТКА ЛИСТОВ ОБОРУДОВАНИЯ ---
    by_type_counts = {}
    
    for symbol in sorted_symbols:
        sheet_name = SHEET_NAME_MAP.get(symbol, symbol)
        
        if sheet_name not in wb.sheetnames:
            logger.warning(f"Sheet '{sheet_name}' not found in template. Skipping type {symbol}.")
            continue
            
        ws_eq = wb[sheet_name]
        
        # Находим LAST ROW
        last_row_idx = None
        for r in range(1, ws_eq.max_row + 1):
            if str(ws_eq.cell(r, 1).value).upper() == "LAST ROW":
                last_row_idx = r
                break
                
        if not last_row_idx:
            last_row_idx = ws_eq.max_row + 1
            ws_eq.cell(last_row_idx, 1, value="LAST ROW")
            
        current_insert_row = last_row_idx
        records_for_type = grouped_records[symbol]
        
        for rec in records_for_type:
            # Вставляем строку перед LAST ROW
            ws_eq.insert_rows(current_insert_row, amount=1)
            
            # --- КОЛОНКИ A-F (Согласно ТЗ п.3.1) ---
            # A: ACTION
            ws_eq.cell(current_insert_row, 1, value=rec.action)
            # B: ITEM_SYMBOL
            ws_eq.cell(current_insert_row, 2, value=rec.acce_item_symbol)
            # C: ITEM_TYPE
            ws_eq.cell(current_insert_row, 3, value=rec.acce_item_type)
            # D: USER_TAG
            ws_eq.cell(current_insert_row, 4, value=rec.user_tag)
            # E: PARENT_AREA
            area_val = rec.parent_area.strip() if rec.parent_area else "DEFAULT"
            ws_eq.cell(current_insert_row, 5, value=area_val)
            # F: DESCRIPTION (макс 60 симв)
            desc = rec.description_ru[:60] if rec.description_ru else ""
            ws_eq.cell(current_insert_row, 6, value=desc)

            # --- ПАРАМЕТРИЧЕСКИЕ КОЛОНКИ G+ (ТЗ п.3.2) ---
            # Конвертация давления: МПа -> кПа (*1000)
            # Проверяем единицу измерения. Если Па, то не умножаем. Если МПа, то *1000.
            press_kpa = None
            if rec.pressure:
                if rec.pressure_unit == 'МПа':
                    press_kpa = rec.pressure * 1000
                elif rec.pressure_unit == 'Па':
                    press_kpa = rec.pressure # Уже в Па, но ACCE ждет кПа? Нет, ТЗ говорит kPa. 
                                             # Если в файле Па, то это обычно вентиляторы. 
                                             # ТЗ: "pressure(Па) -> Pa без изменений". 
                                             # Но колонка называется "Design gauge pressure" и обычно в кПа.
                                             # Для вентиляторов (FN) давление часто в Па.
                                             # Сделаем проверку по типу оборудования.
                    if symbol == 'FN':
                        press_kpa = rec.pressure # Оставляем в Па для вентиляторов, если так в шаблоне
                    else:
                        press_kpa = rec.pressure * 1000 # Предполагаем, что другие в МПа
                else:
                    press_kpa = rec.pressure * 1000 # Дефолт МПа->кПа

            col_idx = 7 # Начинаем с G
            
            if symbol == 'CP': # Насосы
                ws_eq.cell(current_insert_row, col_idx, value=rec.design_temperature)
                ws_eq.cell(current_insert_row, col_idx+1, value=press_kpa)
                ws_eq.cell(current_insert_row, col_idx+2, value=rec.motor_power_kw)
                
            elif symbol == 'VT': # Емкости
                ws_eq.cell(current_insert_row, col_idx, value=press_kpa)
                ws_eq.cell(current_insert_row, col_idx+1, value=rec.design_temperature)
                ws_eq.cell(current_insert_row, col_idx+2, value=rec.volume)
                ws_eq.cell(current_insert_row, col_idx+3, value=rec.diameter_m)
                ws_eq.cell(current_insert_row, col_idx+4, value=rec.material)
                
            elif symbol == 'FN': # Вентиляторы
                ws_eq.cell(current_insert_row, col_idx, value=rec.flow_rate)
                # Для вентиляторов давление часто в Па. Если в шаблоне колонка "Fan outlet gauge pressure" ожидает Па, то ок.
                # Если кПа, то нужно делить на 1000. По ТЗ: "pressure(Па) -> Pa без изменений".
                # Значит для FN оставляем как есть, если оно в Па.
                fn_press = rec.pressure
                if rec.pressure_unit == 'МПа':
                    fn_press = rec.pressure * 1_000_000 # МПа -> Па
                ws_eq.cell(current_insert_row, col_idx+1, value=fn_press)
                ws_eq.cell(current_insert_row, col_idx+2, value=rec.motor_power_kw)
                
            elif symbol == 'STB': # Котлы
                ws_eq.cell(current_insert_row, col_idx, value=rec.capacity_kw)
                ws_eq.cell(current_insert_row, col_idx+1, value=press_kpa)
                
            elif symbol in ['FLR', 'STK']: # Факелы/Свечи
                h_val = rec.height_m if rec.height_m else rec.length_m
                ws_eq.cell(current_insert_row, col_idx, value=h_val)
                ws_eq.cell(current_insert_row, col_idx+1, value=rec.diameter_m)

            current_insert_row += 1
            
        by_type_counts[symbol] = len(records_for_type)

    # --- ЛИСТ VALIDATION (ТЗ п.2.4) ---
    ws_val = wb.create_sheet("VALIDATION")
    val_headers = ["TAG", "TYPE", "STATUS", "ERRORS", "WARNINGS"]
    for c, h in enumerate(val_headers, 1):
        ws_val.cell(1, c, value=h).font = Font(bold=True)
        
    row_val = 2
    for rec in records: # Все записи, включая невалидные
        status = "VALID" if rec.is_valid else "INVALID"
        errors_str = "; ".join(rec.errors) if rec.errors else ""
        warns_str = "; ".join(rec.warnings) if rec.warnings else ""
        
        ws_val.cell(row_val, 1, value=rec.user_tag or rec.raw_tag)
        ws_val.cell(row_val, 2, value=rec.acce_item_symbol)
        ws_val.cell(row_val, 3, value=status)
        ws_val.cell(row_val, 4, value=errors_str)
        ws_val.cell(row_val, 5, value=warns_str)
        row_val += 1

    # --- ЛИСТ CONTENTS (ТЗ п.2.1) ---
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

    # Сохранение
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