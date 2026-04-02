"""
parser_docx.py — Парсер DOCX-файлов с перечнями технологического оборудования.

Принцип работы:
  1. docx → читаем таблицу напрямую через python-docx (без конвертации в PDF)
  2. Разворачиваем строки с несколькими тэгами в отдельные записи
  3. Парсим тех. характеристики через русскоязычные regex-паттерны
  4. Классифицируем по префиксу тэга → Icarus item symbol
  5. Транслитерируем тэг → ACCE user_tag

НЕ ПАРСИМ:
  - Среда (рабочая среда — не нужна Icarus для расчёта стоимости)

ПАРСИМ:
  - Тэг → raw_tag, user_tag
  - Наименование → description_ru
  - Кол-во → quantity (используется для разворачивания)
  - Материал → material (нужен Icarus: shell_material)
  - Тех. характеристики → design_temperature, operating_temperature,
    pressure, volume, diameter_m, length_m, flow_rate,
    motor_power_kw, heat_transfer_area_m2, weight_unit_kg, height_m

Источник классификации: Icarus Reference Guide V15, Chapter 1, pp. 1-3
"""

import re
from dataclasses import dataclass, field
from typing import Optional

from docx import Document


# ══════════════════════════════════════════════════════════════════
# ТРАНСЛИТЕРАЦИЯ
# Российские тэги → латиница для ACCE user_tag
# ══════════════════════════════════════════════════════════════════

_CYR_TO_LAT = {
    'А':'A','Б':'B','В':'V','Г':'G','Д':'D','Е':'E','Ё':'Yo',
    'Ж':'Zh','З':'Z','И':'I','Й':'Y','К':'K','Л':'L','М':'M',
    'Н':'N','О':'O','П':'P','Р':'R','С':'S','Т':'T','У':'U',
    'Ф':'F','Х':'Kh','Ц':'Ts','Ч':'Ch','Ш':'Sh','Щ':'Sch',
    'Ъ':'','Ы':'Y','Ь':'','Э':'E','Ю':'Yu','Я':'Ya',
    'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'yo',
    'ж':'zh','з':'z','и':'i','й':'y','к':'k','л':'l','м':'m',
    'н':'n','о':'o','п':'p','р':'r','с':'s','т':'t','у':'u',
    'ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh','щ':'sch',
    'ъ':'','ы':'y','ь':'','э':'e','ю':'yu','я':'ya',
}

_FORBIDDEN = re.compile(r'[!@#$%&*()?/{}[\]*=<>,`~^:;|"\' \\+\s～]')


def _transliterate(text: str) -> str:
    """Транслитерирует кириллицу в латиницу."""
    return ''.join(_CYR_TO_LAT.get(ch, ch) for ch in text)


def _make_user_tag(raw_tag: str) -> str:
    """
    Создаёт валидный ACCE user_tag из российского тэга.
    Р-001 → R-001, Км-002 → Km-002, ДЕ-001 → DE-001
    Правила ACCE: макс. 20 символов, без пробелов и спецсимволов.
    """
    lat = _transliterate(raw_tag.strip())
    # Убираем запрещённые символы (кроме дефиса и цифр)
    safe = _FORBIDDEN.sub('-', lat)
    safe = re.sub(r'-{2,}', '-', safe).strip('-')[:20]
    return safe if safe else f"TAG-{raw_tag[:10]}"


# ══════════════════════════════════════════════════════════════════
# КЛАССИФИКАЦИЯ ПО ПРЕФИКСУ ТЭГА
# Источник: Icarus Reference Guide V15, Chapter 1, pp. 1-3
#
# Префикс тэга в российских проектах — надёжный идентификатор типа.
# Это надёжнее чем keyword-матчинг по наименованию.
# ══════════════════════════════════════════════════════════════════

# (icarus_symbol, icarus_subtype)
_PREFIX_MAP = {
    # Реакторы — вертикальные сосуды давления
    'Р':   ('VT',  'CYLINDER'),
    # Сепараторы — вертикальные сосуды
    'С':   ('VT',  'CYLINDER'),
    # Теплообменники, испарители, холодильники, подогреватели
    'Т':   ('HE',  'FIXED T S'),
    # Газодувки — вентиляторы/нагнетатели
    'Г':   ('FN',  'CENTRIF'),
    # Компрессоры — поршневые (наиболее распространены в российских проектах)
    'Км':  ('GC',  'RECIP MOTR'),
    # Ёмкости (сырьевые, буферные, дренажные, аварийные)
    'Е':   ('VT',  'CYLINDER'),
    # Дренажные ёмкости
    'ДЕ':  ('VT',  'CYLINDER'),
    # Насосы — центробежные по умолчанию
    'Н':   ('CP',  'CENTRIF'),
    # Факелы
    'Ф':   ('FLR', 'SELF SUPP'),
    # Свечи рассеивания — дымовые трубы/стеки
    'СР':  ('STK', 'STACK'),
}

# Длинные префиксы должны проверяться первыми
_PREFIX_ORDER = ['ДЕ', 'Км', 'СР', 'Р', 'С', 'Т', 'Г', 'Е', 'Н', 'Ф']


def _classify_by_tag(raw_tag: str) -> tuple[str, str]:
    """
    Определяет Icarus item symbol и подтип по префиксу тэга.
    Возвращает (symbol, subtype) или ('', '') если не определено.
    """
    tag = raw_tag.strip()
    for prefix in _PREFIX_ORDER:
        if tag.startswith(prefix):
            return _PREFIX_MAP.get(prefix, ('', ''))
    return '', ''


# ══════════════════════════════════════════════════════════════════
# ПАРСИНГ ТЕХ. ХАРАКТЕРИСТИК (русскоязычные паттерны)
# ══════════════════════════════════════════════════════════════════

# Общий паттерн: "Ключевое слово – значение единица"
# Разделитель: –, —, -, =
_SEP = r'\s*[–—\-=]\s*'
_NUM = r'([\d\s,\.]+)'     # число с возможными пробелами и запятыми
_NUM_MINUS = r'((?:минус\s+)?[\d\s,\.]+)'  # "минус 40" тоже число


def _num(s: str) -> Optional[float]:
    """Парсит число из строки, включая 'минус 40' → -40.0"""
    if s is None:
        return None
    s = s.strip()
    negative = bool(re.match(r'минус', s, re.I))
    digits = re.sub(r'[^\d,\.]', '', s.replace(',', '.'))
    if not digits:
        return None
    try:
        val = float(digits)
        return -val if negative else val
    except ValueError:
        return None


def _parse_tech(text: str) -> dict:
    """
    Разбирает строку тех. характеристик на типизированные поля.

    Обрабатывает все паттерны из реального файла:
      Трасч. – 450 °С          → design_temperature
      Трасч = 450 °С           → design_temperature  (без точки и с =)
      Траб. – 45 °С            → operating_temperature
      Траб. – минус 40°С       → operating_temperature = -40
      Ррасч. – 4,9 МПа (изб.) → pressure (расчётное)
      Рраб. – 0,07 МПа (изб.) → pressure (рабочее)
      Рнагнетания – 4,46 МПа  → pressure (нагнетание для насосов)
      нагнетание – 0,021 МПа  → pressure (для газодувок/компрессоров)
      Объем – 20,0 м3          → volume
      Диаметр – 450 мм         → diameter_m
      Длина (высота) – 3950 мм → length_m
      Высота ствола – 10 м     → height_m
      Диаметр ствола – 50 мм  → diameter_m
      Производительность – 470 нм3/ч → flow_rate (нм3/ч)
      Производительность – 0,045 м3/ч → flow_rate (м3/ч)
      Производительность – 170 мл/ч   → flow_rate (мл/ч)
      Мощность – 9,0 кВт       → motor_power_kw
      Площадь теплообмена – 3,3 м2 → heat_transfer_area_m2
      Масса – 2 т              → weight_unit_kg (× 1000)
      Тепловая нагрузка – 0,010 Гкал/ч → heat_duty_gcalh
    """
    if not text:
        return {}

    r = {}

    def _find(pattern: str) -> Optional[float]:
        m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if m:
            return _num(m.group(1))
        return None

    # ── Температуры ───────────────────────────────────────────────

    # Расчётная температура: Трасч. / Трасч (с точкой или без, с – или =)
    # Берём ПЕРВОЕ вхождение — это рабочая среда (не теплоноситель)
    t_calc = re.search(
        r'Трасч[\.]*\s*[–—\-=]\s*' + _NUM_MINUS + r'\s*°?С',
        text, re.IGNORECASE)
    if t_calc:
        r['design_temperature'] = _num(t_calc.group(1))

    # Рабочая температура: Траб.
    t_work = re.search(
        r'Траб[\.]*\s*[–—\-=]\s*' + _NUM_MINUS + r'\s*°?С',
        text, re.IGNORECASE)
    if t_work:
        r['operating_temperature'] = _num(t_work.group(1))

    # ── Давление ─────────────────────────────────────────────────
    # Приоритет: Ррасч > Рнагнетания > нагнетание > Рраб
    # Берём первое расчётное давление (рабочая среда, не теплоноситель)

    p_calc = re.search(
        r'Ррасч[\.]*\s*[–—\-=]\s*' + _NUM + r'\s*МПа',
        text, re.IGNORECASE)
    p_nagn = re.search(
        r'Рнагнетани[яе]\s*[–—\-=]\s*' + _NUM + r'\s*МПа',
        text, re.IGNORECASE)
    p_nagn2 = re.search(
        r'нагнетание\s*[–—\-=]\s*' + _NUM + r'\s*МПа',
        text, re.IGNORECASE)
    p_work = re.search(
        r'Рраб[\.]*\s*[–—\-=]\s*' + _NUM + r'\s*МПа',
        text, re.IGNORECASE)

    if p_calc:
        r['pressure'] = _num(p_calc.group(1))
    elif p_nagn:
        r['pressure'] = _num(p_nagn.group(1))
    elif p_nagn2:
        r['pressure'] = _num(p_nagn2.group(1))
    elif p_work:
        r['pressure'] = _num(p_work.group(1))

    # ── Объём ─────────────────────────────────────────────────────
    vol = re.search(r'Объем\s*[–—\-=]\s*' + _NUM + r'\s*м3', text, re.IGNORECASE)
    if vol:
        r['volume'] = _num(vol.group(1))

    # ── Размеры ───────────────────────────────────────────────────
    # Диаметр (приоритет: "Диаметр ствола", потом просто "Диаметр")
    d_stv = re.search(r'Диаметр\s+ствола\s*[–—\-=]\s*' + _NUM + r'\s*(мм|м)\b',
                      text, re.IGNORECASE)
    d_gen = re.search(r'Диаметр\s*[–—\-=]\s*' + _NUM + r'\s*(мм|м)\b',
                      text, re.IGNORECASE)
    dm = d_stv or d_gen
    if dm:
        val = _num(dm.group(1))
        unit = dm.group(2).lower()
        if val is not None:
            r['diameter_m'] = val / 1000 if unit == 'мм' else val

    # Длина / высота (основной размер)
    ln = re.search(
        r'(?:Длина|Высота\s+ствола|Длина\s*\(высота\))\s*[–—\-=]\s*'
        + _NUM + r'\s*(мм|м)\b',
        text, re.IGNORECASE)
    if ln:
        val = _num(ln.group(1))
        unit = ln.group(2).lower()
        if val is not None:
            r['length_m'] = val / 1000 if unit == 'мм' else val

    # ── Производительность / расход ───────────────────────────────
    # нм3/ч, м3/ч, мл/ч — всё в flow_rate с единицей
    prod = re.search(
        r'Производительность\s*[–—\-=]\s*' + _NUM
        + r'\s*(нм3/ч|м3/ч|мл/ч|м3/час)',
        text, re.IGNORECASE)
    if prod:
        val = _num(prod.group(1))
        unit = prod.group(2).lower()
        # мл/ч → м3/ч для единообразия (÷ 1 000 000)
        if unit == 'мл/ч' and val is not None:
            val = val / 1_000_000
            unit = 'м3/ч'
        r['flow_rate']      = val
        r['flow_rate_unit'] = unit

    # ── Мощность ─────────────────────────────────────────────────
    pw = re.search(r'Мощность\s*[–—\-=]\s*' + _NUM + r'\s*кВт', text, re.IGNORECASE)
    if pw:
        r['motor_power_kw'] = _num(pw.group(1))

    # ── Площадь теплообмена ───────────────────────────────────────
    area = re.search(
        r'Площадь\s+теплообмена\s*[–—\-=]\s*' + _NUM + r'\s*м2',
        text, re.IGNORECASE)
    if area:
        r['heat_transfer_area_m2'] = _num(area.group(1))

    # ── Тепловая нагрузка ─────────────────────────────────────────
    heat = re.search(
        r'Тепловая\s+нагрузка\s*[–—\-=]\s*' + _NUM + r'\s*Гкал/ч',
        text, re.IGNORECASE)
    if heat:
        r['heat_duty_gcalh'] = _num(heat.group(1))

    # ── Масса ─────────────────────────────────────────────────────
    mass = re.search(r'Масса\s*[–—\-=]\s*' + _NUM + r'\s*т\.?', text, re.IGNORECASE)
    if mass:
        val = _num(mass.group(1))
        if val is not None:
            r['weight_unit_kg'] = val * 1000   # т → кг

    return r


# ══════════════════════════════════════════════════════════════════
# РАЗВОРАЧИВАНИЕ ТЭГОВ
# "Р-001 Р-002 Р-003 Р-004" → ['Р-001', 'Р-002', 'Р-003', 'Р-004']
# ══════════════════════════════════════════════════════════════════

def _split_tags(raw: str) -> list[str]:
    """
    Разбивает ячейку с одним или несколькими тэгами на список.
    Тэги разделены пробелами или переносами строк.
    Паттерн тэга: одна или две заглавные кириллические буквы + дефис + цифры.
    """
    # [А-ЯЁ]{1,2}  — одна или две заглавные (Р, ДЕ, СР)
    # [а-яё]?      — опциональная строчная (Км → К + м)
    # -\d+         — дефис и цифры
    tags = re.findall(r'[А-ЯЁ]{1,2}[а-яё]?-\d+', raw)
    return tags if tags else []


# ══════════════════════════════════════════════════════════════════
# СХЕМА ДАННЫХ (расширенная для docx-специфичных полей)
# ══════════════════════════════════════════════════════════════════

@dataclass
class ACCERecord:
    # Трассировка
    source_file:  str = ""
    source_sheet: str = ""   # для docx — имя документа
    source_row:   int = 0

    # Идентификация
    raw_tag:        str = ""
    user_tag:       str = ""
    description_ru: str = ""
    description_en: str = ""    # в docx обычно пуст
    parent_area:    str = ""
    quantity:       int = 1
    material:       str = ""    # материал корпуса — нужен Icarus

    # Вес
    weight_unit_kg:  Optional[float] = None
    weight_total_kg: Optional[float] = None
    weight_total_t:  Optional[float] = None

    # Тех. характеристики
    tech_raw:              str   = ""
    design_temperature:    Optional[float] = None   # Трасч, °С
    operating_temperature: Optional[float] = None   # Траб, °С
    pressure:              Optional[float] = None   # МПа
    pressure_unit:         str   = "МПа"
    volume:                Optional[float] = None   # м3
    diameter_m:            Optional[float] = None   # м
    length_m:              Optional[float] = None   # м
    flow_rate:             Optional[float] = None   # м3/ч или нм3/ч
    flow_rate_unit:        str   = ""
    motor_power_kw:        Optional[float] = None   # кВт
    heat_transfer_area_m2: Optional[float] = None   # м2 (для HE)
    heat_duty_gcalh:       Optional[float] = None   # Гкал/ч

    # Тип Icarus
    acce_item_symbol: str = ""
    acce_item_type:   str = ""

    # ACCE управляющее поле
    action:   str  = "NEW"
    is_valid: bool = False
    errors:   list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════
# ИНДЕКСЫ КОЛОНОК (подтверждено по Samples_3.docx)
# ══════════════════════════════════════════════════════════════════
COL_NUM   = 0   # №
COL_TAG   = 1   # Индекс по схеме
COL_NAME  = 2   # Наименование оборудования
COL_QTY   = 3   # Кол-во, шт
# COL 4 — Среда — НЕ ПАРСИМ
COL_MAT   = 5   # Материал
COL_TECH  = 6   # Техническая характеристика

HEADER_ROWS = 2   # первые 2 строки — заголовки


# ══════════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ
# ══════════════════════════════════════════════════════════════════

def parse_docx(filepath: str) -> tuple[list[ACCERecord], list[dict]]:
    """
    Читает DOCX-файл с перечнем технологического оборудования.

    Логика:
      1. Открываем docx напрямую через python-docx
      2. Читаем первую таблицу (основной перечень)
      3. Каждую строку разворачиваем по количеству тэгов:
         - Несколько тэгов в ячейке → отдельная запись на каждый тэг
         - qty > 1 с одним тэгом → отдельная запись на каждую единицу
           с суффиксом -1, -2, ... к user_tag
      4. Парсим тех. характеристики
      5. Классифицируем по префиксу тэга

    Возвращает:
        records        — список ACCERecord
        parse_warnings — строки с проблемами
    """
    doc = Document(filepath)

    if not doc.tables:
        raise ValueError(f"В файле {filepath} не найдено таблиц")

    table = doc.tables[0]
    records:        list[ACCERecord] = []
    parse_warnings: list[dict]       = []
    doc_name = filepath.split('/')[-1]

    for row_idx, row in enumerate(table.rows):
        if row_idx < HEADER_ROWS:
            continue

        cells = [c.text.strip().replace('\n', ' ') for c in row.cells]

        # Пропускаем пустые строки
        if not any(cells):
            continue

        raw_tags_cell = cells[COL_TAG] if len(cells) > COL_TAG else ""
        name_raw      = cells[COL_NAME] if len(cells) > COL_NAME else ""
        qty_raw       = cells[COL_QTY]  if len(cells) > COL_QTY  else ""
        material      = cells[COL_MAT]  if len(cells) > COL_MAT  else ""
        tech_raw      = cells[COL_TECH] if len(cells) > COL_TECH else ""

        # Парсим количество
        try:
            qty = int(re.search(r'\d+', qty_raw).group()) if qty_raw else 1
        except (AttributeError, ValueError):
            qty = 1

        # Получаем список тэгов из ячейки
        tags = _split_tags(raw_tags_cell)

        if not tags:
            parse_warnings.append({
                "row": row_idx,
                "cell": raw_tags_cell,
                "reason": "Тэг не распознан — строка пропущена"
            })
            continue

        # Парсим тех. характеристики один раз для всей строки
        tech = _parse_tech(tech_raw)

        # ── Разворачивание ────────────────────────────────────────
        # Случай 1: несколько тэгов в ячейке (Р-001 Р-002 Р-003 Р-004)
        #   → qty уже указан в таблице (= числу тэгов)
        #   → создаём по одной записи на каждый тэг
        #
        # Случай 2: один тэг + qty > 1 (Н-001, qty=2)
        #   → создаём qty записей с суффиксами -1, -2, ...

        if len(tags) > 1:
            # Несколько тэгов — каждый получает свою запись
            items_to_create = [(tag, 1) for tag in tags]
        elif qty > 1:
            # Один тэг + qty > 1 — разворачиваем с суффиксами
            items_to_create = [(tags[0], 1)] * qty
        else:
            items_to_create = [(tags[0], 1)]

        for item_idx, (raw_tag, _) in enumerate(items_to_create):
            # Генерируем user_tag
            base_tag = _make_user_tag(raw_tag)

            if qty > 1 and len(tags) == 1:
                # Добавляем суффикс для отличия дубликатов
                user_tag = f"{base_tag}-{item_idx + 1}"[:20]
            else:
                user_tag = base_tag

            # Классификация по префиксу тэга
            symbol, subtype = _classify_by_tag(raw_tag)

            row_warns = []
            if not symbol:
                row_warns.append(
                    f"Не удалось классифицировать тэг '{raw_tag}'"
                )

            record = ACCERecord(
                source_file  = filepath,
                source_sheet = doc_name,
                source_row   = row_idx,

                raw_tag         = raw_tag,
                user_tag        = user_tag,
                description_ru  = name_raw,
                parent_area     = "",        # заполняется вручную или из проекта
                quantity        = 1,         # уже развёрнуто

                material        = material,

                weight_unit_kg  = tech.get('weight_unit_kg'),

                tech_raw               = tech_raw,
                design_temperature     = tech.get('design_temperature'),
                operating_temperature  = tech.get('operating_temperature'),
                pressure               = tech.get('pressure'),
                pressure_unit          = "МПа",
                volume                 = tech.get('volume'),
                diameter_m             = tech.get('diameter_m'),
                length_m               = tech.get('length_m'),
                flow_rate              = tech.get('flow_rate'),
                flow_rate_unit         = tech.get('flow_rate_unit', ''),
                motor_power_kw         = tech.get('motor_power_kw'),
                heat_transfer_area_m2  = tech.get('heat_transfer_area_m2'),
                heat_duty_gcalh        = tech.get('heat_duty_gcalh'),

                acce_item_symbol = symbol,
                acce_item_type   = subtype,

                action    = "NEW",
                is_valid  = len(row_warns) == 0,
                warnings  = row_warns,
            )

            records.append(record)

            if row_warns:
                parse_warnings.append({
                    "row": row_idx, "tag": raw_tag,
                    "reason": "; ".join(row_warns)
                })

    return records, parse_warnings