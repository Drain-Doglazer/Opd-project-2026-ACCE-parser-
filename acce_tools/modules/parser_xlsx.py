"""
parser_xlsx.py
==============
Парсер Excel-файлов с перечнями оборудования.


ЧТО ПАРСИМ (нужно ACCE):
  - Тэговый номер оборудования → user_tag
  - Наименование (рус + англ)  → description_ru, description_en
  - Тех. характеристики        → flow_rate, pressure, volume,
                                  capacity_kw, diameter_m, length_m,
                                  dn_mm, lift_capacity_t
  - Количество                 → quantity
  - Вес единицы (кг)           → weight_unit_kg   (нужен ACCE для расчёта)
  - Всего масса (кг / тонны)   → weight_total_kg, weight_total_t
  - ТИТУЛ (область/зона)       → parent_area

ЧТО НЕ ПАРСИМ:
  - Код единичной расценки (колонка E) — российская локальная расценка,
    не используется в Icarus/ACCE. Тип оборудования мы определяем сами
    по наименованию через keyword-матчинг.
  - № п/п — порядковый номер строки, не нужен ACCE.
  - Колонка I (пустая, слитая).

ОПРЕДЕЛЕНИЕ ТИПА ОБОРУДОВАНИЯ (acce_item_symbol):
  Выполняется по ключевым словам из наименования.
  Соответствует символам Icarus (CP, FN, GC, VT, HE, FU, STB, CE, HO).
  Подробнее: Aspen Icarus Reference Guide V15, Chapter 1, pp. 1-3.
"""

import re
import openpyxl
from dataclasses import dataclass, field
from typing import Optional


# ─────────────────────────────────────────────────────────────────
# СХЕМА ДАННЫХ
# Содержит только поля, которые нужны ACCE.
# ─────────────────────────────────────────────────────────────────

@dataclass
class ACCERecord:

    # --- Трассировка источника ---
    source_file:  str = ""   # путь к файлу
    source_sheet: str = ""   # имя листа
    source_row:   int = 0    # номер строки (для отчёта об ошибках)

    # --- Идентификация (нужна ACCE) ---
    raw_tag:         str = ""   # тэговый номер из файла
    user_tag:        str = ""   # очищенный тэг для ACCE (макс. 20 символов)
    description_ru:  str = ""   # наименование по-русски
    description_en:  str = ""   # наименование по-английски
    parent_area:     str = ""   # ТИТУЛ → зона/область в ACCE (AREAS sheet)
    quantity:        int = 1    # количество единиц

    # --- Вес (нужен ACCE для расчёта установки) ---
    weight_unit_kg:  Optional[float] = None   # вес одной единицы, кг
    weight_total_kg: Optional[float] = None   # общий вес, кг
    weight_total_t:  Optional[float] = None   # общий вес, тонны

    # --- Тех. характеристики (нужны Icarus для расчёта) ---
    tech_raw:          str   = ""    # оригинальная строка (для аудита)
    flow_rate:         Optional[float] = None  # Q, м3/ч или т/ч
    flow_rate_unit:    str   = ""
    pressure:          Optional[float] = None  # P, МПа или Па
    pressure_unit:     str   = ""
    volume:            Optional[float] = None  # V, м3
    volume_unit:       str   = ""
    capacity_kw:       Optional[float] = None  # Q когда единица кВт (котлы)
    diameter_m:        Optional[float] = None  # Φ, м (инсинераторы, сосуды)
    length_m:          Optional[float] = None  # L, м (второй размер Φ×L)
    dn_mm:             Optional[int]   = None  # DN, мм (трубопроводная арматура)
    lift_capacity_t:   Optional[float] = None  # T=, т (краны, тали)

    # --- Классификация Icarus ---
    # Символ типа из Icarus Reference Guide V15, Chapter 1, pp. 1-3
    # CP, FN, GC, VT, HE, FU, STB, CE, HO, ...
    acce_item_symbol: str = ""
    acce_item_type:   str = ""   # подтип, например CENTRIF, API 610, CYLINDER

    # --- Управляющие поля ACCE (заполняет экспортёр) ---
    action:     str = "NEW"    # NEW / CHANGE / DELETE / IGNORE

    # --- Результат валидации ---
    is_valid:   bool = False
    errors:     list = field(default_factory=list)
    warnings:   list = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────
# КЛАССИФИКАЦИЯ ОБОРУДОВАНИЯ ПО КЛЮЧЕВЫМ СЛОВАМ
# Источник: Icarus Reference Guide V15, Chapter 1, Table of Equipment
# ─────────────────────────────────────────────────────────────────

# Каждая запись: ([ключевые слова], icarus_symbol, icarus_subtype)
# Проверяются по порядку — первое совпадение побеждает.
KEYWORD_MAP = [
    # Насосы (CP / P / GP)
    (["centrif.*pump", "cp ", "centrifugal pump",
      "насос.*центробежн", "центробежн.*насос",
      "circulating.*pump", "water pump", "насос для"],
     "CP", "CENTRIF"),
    (["gear pump", "шестерёнч.*насос"],
     "GP", "GEAR"),
    (["diaphragm pump", "мембранн.*насос", "мембранный насос"],
     "P", "DIAPHRAGM"),
    (["pump", "насос"],
     "CP", "CENTRIF"),   # общий случай — насос

    # Вентиляторы и нагнетатели (FN)
    (["fan", "blower", "вентилятор", "нагнетатель",
      "induced draft", "forced draft", "дутьевой", "дымосос"],
     "FN", "CENTRIF"),

    # Компрессоры (AC / GC)
    (["air compressor", "воздушн.*компрессор", "компрессор.*воздух"],
     "AC", "CENTRIF M"),
    (["compressor", "компрессор"],
     "GC", "CENTRIF"),

    # Теплообменники (HE / RB)
    (["heat exchanger", "теплообменник", "cooler", "охладитель",
      "condenser", "конденсатор", "air cooler", "воздушн.*охлад"],
     "HE", "FLOAT HEAD"),
    (["reboiler", "рибойлер", "кипятильник"],
     "RB", "THERMOSIPHON"),

    # Печи, огневые нагреватели, инсинераторы (FU)
    (["incinerator", "инсинератор", "furnace", "печь",
      "fired heater", "огневой нагрев", "boiler.*fired",
      "горелка", "burner"],
     "FU", "HEATER"),

    # Котлы (STB)
    (["boiler", "котёл", "котел", "htmboiler", "htm boiler",
      "steam boiler", "паровой котёл"],
     "STB", "BOILER"),

    # Вертикальные сосуды, ёмкости, резервуары (VT)
    (["vertical.*tank", "vertical.*vessel", "vertical.*drum",
      "вертикальн.*ёмкост", "вертикальн.*сосуд",
      "деаэратор", "deaerator",
      "softening device", "умягчитель", "умягчен",
      "tank", "vessel", "drum", "ёмкость", "сосуд",
      "резервуар", "бак"],
     "VT", "CYLINDER"),

    # Горизонтальные сосуды (HT)
    (["horizontal.*tank", "horizontal.*vessel", "horizontal.*drum",
      "горизонтальн.*сосуд", "горизонтальн.*ёмкост"],
     "HT", "HORIZ DRUM"),

    # Краны мостовые (CE)
    (["bridge crane", "overhead crane", "мостовой кран",
      "кран мостовой", "кран-балка"],
     "CE", "BRIDGE CRN"),

    # Тали, подъёмники (HO)
    (["hoist", "electric hoist", "таль", "подъёмник",
      "электрический подъёмник"],
     "HO", "HOIST"),

    # Клапаны — статическое оборудование, не рассчитывается Icarus отдельно
    # Оставляем в VT как «статическое» (ближайший тип)
    (["valve", "клапан", "задвижка"],
     "VT", "CYLINDER"),
]


def _classify(description_en: str, description_ru: str) -> tuple[str, str]:
    """
    Определяет Icarus item symbol и подтип по наименованию.
    Возвращает (symbol, subtype) или ("", "") если не определено.
    """
    text = (description_en + " " + description_ru).lower()
    for keywords, symbol, subtype in KEYWORD_MAP:
        for kw in keywords:
            if re.search(kw, text, re.IGNORECASE):
                return symbol, subtype
    return "", ""


# ─────────────────────────────────────────────────────────────────
# ПАРСИНГ ТЕХ. ХАРАКТЕРИСТИК
# Обрабатывает все паттерны, встреченные в реальных файлах:
#   Q=160m3/h  P=0.8MPa  V=6m3  T=5t H=6m  Φ3.8×25m  DN500
# ─────────────────────────────────────────────────────────────────

_UNITS = r'(m3/h|m3|kW|MPa|kPa|Pa|t/h|kg/h|t|m)?'


def _extract_param(text: str, key: str) -> tuple[Optional[float], str]:
    """Извлекает Q=, P=, V= и т.д. Для диапазонов берёт нижнее значение."""
    pattern = (
        rf'(?<![A-Za-z]){re.escape(key)}\s*=\s*'
        rf'([0-9][0-9,\.]*)'
        rf'(?:\s*[-~]\s*[0-9][0-9,\.]*)?'
        rf'\s*' + _UNITS
    )
    m = re.search(pattern, text, re.IGNORECASE)
    if not m:
        return None, ""
    raw = m.group(1).replace(',', '')
    unit = m.group(2) or ""
    try:
        return float(raw), unit
    except ValueError:
        return None, ""


def _extract_dn(text: str) -> Optional[int]:
    m = re.search(r'DN\s*(\d+)', text, re.IGNORECASE)
    return int(m.group(1)) if m else None


def _extract_phi(text: str) -> tuple[Optional[float], Optional[float]]:
    """Φ3.8×25m → (diameter_m=3.8, length_m=25.0)"""
    m = re.search(r'[ΦΦ]?\s*([\d.]+)\s*[×x]\s*([\d.]+)', text)
    if m:
        return float(m.group(1)), float(m.group(2))
    return None, None


def _extract_lift(text: str) -> Optional[float]:
    """T=5t → 5.0 т (грузоподъёмность крана/тали)"""
    m = re.search(r'T\s*=\s*([\d.]+)\s*t', text)
    return float(m.group(1)) if m else None


def _parse_tech(raw: str) -> dict:
    """
    Разбирает строку тех. характеристик на структурированные поля.
    Возвращает словарь только с найденными значениями.
    """
    if not raw:
        return {}
    result = {}

    q_val, q_unit = _extract_param(raw, 'Q')
    p_val, p_unit = _extract_param(raw, 'P')
    v_val, v_unit = _extract_param(raw, 'V')
    dn             = _extract_dn(raw)
    dia, length    = _extract_phi(raw)
    lift           = _extract_lift(raw)

    if q_val is not None:
        if q_unit and q_unit.lower() == 'kw':
            result['capacity_kw']    = q_val        # котлы: Q в кВт = тепловая мощность
        else:
            result['flow_rate']      = q_val        # насосы/вентиляторы: Q = расход
            result['flow_rate_unit'] = q_unit

    if p_val is not None:
        result['pressure']      = p_val
        result['pressure_unit'] = p_unit

    if v_val is not None:
        result['volume']      = v_val
        result['volume_unit'] = v_unit

    if dn is not None:
        result['dn_mm'] = dn

    if dia is not None:
        result['diameter_m'] = dia
    if length is not None:
        result['length_m'] = length

    if lift is not None:
        result['lift_capacity_t'] = lift

    return result


# ─────────────────────────────────────────────────────────────────
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ─────────────────────────────────────────────────────────────────

def _build_merge_map(sheet) -> dict:
    """
    openpyxl возвращает None для объединённых ячеек (кроме левой верхней).
    Строим словарь (row, col) → значение чтобы ничего не потерять.
    """
    merge_map = {}
    for rng in sheet.merged_cells.ranges:
        val = sheet.cell(row=rng.min_row, column=rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merge_map[(r, c)] = val
    return merge_map


def _cell(sheet, row: int, col: int, mm: dict):
    v = sheet.cell(row=row, column=col).value
    return v if v is not None else mm.get((row, col))


def _str(v) -> str:
    if v is None:
        return ""
    return re.sub(r'\s+', ' ', str(v).strip())


def _float(v) -> Optional[float]:
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "-", "n/a", "tbd", "#REF!", "#VALUE!"):
        return None
    try:
        return float(s)
    except ValueError:
        m = re.search(r'-?\d[\d,\.]*', s)
        if m:
            try:
                return float(m.group().replace(',', ''))
            except ValueError:
                pass
    return None


def _int(v) -> Optional[int]:
    f = _float(v)
    return int(f) if f is not None else None


def _split_bilingual(text: str) -> tuple[str, str]:
    """
    'Насос для воды / Circulating Water Pump' → ('Насос для воды', 'Circulating Water Pump')
    Если разделителя нет — весь текст в русское поле.
    """
    if '/' in text:
        parts = text.split('/', 1)
        return parts[0].strip(), parts[1].strip()
    return text.strip(), ""


_FORBIDDEN = re.compile(r'[!@#$%&*()?/{}[\]*=<>,`~^:;|"\' \\+\s～]')


def _make_user_tag(raw_tag: str, row_num: int) -> str:
    """
    Создаёт валидный ACCE user_tag из сырого тэга.
    Правила ACCE: макс. 20 символов, без пробелов и спецсимволов.
    """
    safe = _FORBIDDEN.sub('-', raw_tag.strip())
    safe = re.sub(r'-{2,}', '-', safe).strip('-')
    safe = safe[:20]
    return safe if safe else f"SEQ-{row_num:04d}"


# ─────────────────────────────────────────────────────────────────
# ИНДЕКСЫ КОЛОНОК (подтверждено по файлу Пример_4.xlsx)
# ─────────────────────────────────────────────────────────────────
COL_SEQ   = 1   # № п/п              — НЕ ПАРСИМ (не нужен ACCE)
COL_AREA  = 2   # ТИТУЛ              → parent_area
COL_TAG   = 3   # ТЭГОВЫЙ НОМЕР      → raw_tag, user_tag
COL_DESC  = 4   # НАИМЕНОВАНИЕ       → description_ru, description_en
COL_CODE  = 5   # КОД ЕД. РАСЦЕНКИ  — НЕ ПАРСИМ (российская расценка)
COL_TECH  = 6   # ТЕХ. ХАРАКТЕРИСТИКИ → flow_rate, pressure, ...
COL_QTY   = 7   # КОЛ.               → quantity
COL_WU    = 8   # ВЕС ЕДИНИЦЫ (кг)  → weight_unit_kg
# COL 9 — пустая объединённая ячейка, пропускаем
COL_WT    = 10  # ВСЕГО МАССА, кг   → weight_total_kg
COL_WTT   = 11  # ВСЕГО МАССА, т    → weight_total_t

HEADER_ROW  = 1   # строка с заголовками
DATA_START  = 3   # первая строка данных (строка 2 — часть объединённого заголовка)
DATA_SHEET  = 'Лист1'


# ─────────────────────────────────────────────────────────────────
# ОСНОВНАЯ ФУНКЦИЯ ПАРСЕРА
# ─────────────────────────────────────────────────────────────────

def parse_xlsx(filepath: str) -> tuple[list[ACCERecord], list[dict]]:
    """
    Читает Excel-файл с перечнем оборудования.

    Возвращает:
        records        — список ACCERecord с данными для ACCE
        parse_warnings — строки, которые были пропущены и причина
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)

    if DATA_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Лист '{DATA_SHEET}' не найден. "
            f"Доступные листы: {wb.sheetnames}"
        )

    ws = wb[DATA_SHEET]
    mm = _build_merge_map(ws)
    records:        list[ACCERecord] = []
    parse_warnings: list[dict]       = []

    for row_idx in range(DATA_START, ws.max_row + 1):

        # Пропускаем полностью пустые строки
        vals = [_cell(ws, row_idx, c, mm) for c in range(1, 12)]
        if all(v is None for v in vals):
            continue

        # Читаем только нужные ACCE поля
        area      = _str(_cell(ws, row_idx, COL_AREA, mm))
        tag       = _str(_cell(ws, row_idx, COL_TAG,  mm))
        desc_raw  = _str(_cell(ws, row_idx, COL_DESC, mm))
        # COL_CODE — код расценки — намеренно не читаем
        tech_raw  = _str(_cell(ws, row_idx, COL_TECH, mm))
        qty       = _int(_cell(ws, row_idx, COL_QTY,  mm))
        w_unit    = _float(_cell(ws, row_idx, COL_WU,  mm))
        w_total   = _float(_cell(ws, row_idx, COL_WT,  mm))
        w_total_t = _float(_cell(ws, row_idx, COL_WTT, mm))

        # Пропускаем итоговые/суммарные строки (нет тэга и нет порядкового номера)
        seq = _int(_cell(ws, row_idx, COL_SEQ, mm))
        if not tag and not seq:
            parse_warnings.append({
                "sheet": DATA_SHEET, "row": row_idx,
                "reason": "Нет тэга и порядкового номера — "
                          "вероятно итоговая строка, пропущена"
            })
            continue

        # Предупреждение если нет тэга
        row_warns = []
        if not tag:
            row_warns.append("Отсутствует тэговый номер оборудования")
        if not desc_raw:
            row_warns.append("Отсутствует наименование")

        # Разбиваем двуязычное наименование
        desc_ru, desc_en = _split_bilingual(desc_raw)

        # Определяем тип оборудования по наименованию (не по коду расценки)
        symbol, subtype = _classify(desc_en, desc_ru)
        if not symbol:
            row_warns.append(
                f"Не удалось определить тип Icarus по наименованию: "
                f"'{desc_en or desc_ru}'"
            )

        # Разбираем тех. характеристики
        tech = _parse_tech(tech_raw)

        # Строим запись
        record = ACCERecord(
            source_file  = filepath,
            source_sheet = DATA_SHEET,
            source_row   = row_idx,

            raw_tag         = tag,
            user_tag        = _make_user_tag(tag, row_idx),
            description_ru  = desc_ru,
            description_en  = desc_en,
            parent_area     = area,
            quantity        = qty or 1,

            weight_unit_kg  = w_unit,
            weight_total_kg = w_total,
            weight_total_t  = w_total_t,

            tech_raw          = tech_raw,
            flow_rate         = tech.get('flow_rate'),
            flow_rate_unit    = tech.get('flow_rate_unit', ''),
            pressure          = tech.get('pressure'),
            pressure_unit     = tech.get('pressure_unit', ''),
            volume            = tech.get('volume'),
            volume_unit       = tech.get('volume_unit', ''),
            capacity_kw       = tech.get('capacity_kw'),
            diameter_m        = tech.get('diameter_m'),
            length_m          = tech.get('length_m'),
            dn_mm             = tech.get('dn_mm'),
            lift_capacity_t   = tech.get('lift_capacity_t'),

            acce_item_symbol = symbol,
            acce_item_type   = subtype,

            action    = "NEW",
            is_valid  = len(row_warns) == 0,
            warnings  = row_warns,
        )

        records.append(record)

        if row_warns:
            parse_warnings.append({
                "sheet": DATA_SHEET, "row": row_idx,
                "tag": tag,
                "reason": "; ".join(row_warns)
            })

    return records, parse_warnings