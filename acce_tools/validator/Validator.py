"""
validator.py — Модуль валидации записей ACCERecord.

Три блока (каскад):
  Блок 1 — Полнота (Группа A): тэг, зона, тип Icarus, уникальность
  Блок 2 — Диапазоны Icarus (Группа B): сверка параметров с лимитами
  Блок 3 — Логика (Группа C): согласованность полей между собой

Конвертация единиц выполняется перед блоком 2.
Пустые поля в блоке 2 пропускаются.
is_valid = True только если errors пуст.
"""

import math
import re
from typing import Optional


# ══════════════════════════════════════════════════════════════════
# КОНВЕРТАЦИЯ ЕДИНИЦ → Icarus SI
# Источник: Icarus Reference Guide V15, единицы измерения
# ══════════════════════════════════════════════════════════════════

def _to_kpa(value: Optional[float], unit: str) -> Optional[float]:
    """Давление → кПа."""
    if value is None:
        return None
    u = unit.upper().replace(" ", "")
    if u in ("МПА", "MPA"):
        return value * 1000
    if u in ("КПА", "KPA"):
        return value
    if u in ("ПА", "PA"):
        return value / 1000
    # Если единица не указана — считаем что уже в МПа (docx-парсер)
    return value * 1000


def _to_celsius(value: Optional[float]) -> Optional[float]:
    """Температура уже в °С — без изменений."""
    return value


def _to_m3h(value: Optional[float], unit: str) -> Optional[float]:
    """Расход → м3/ч."""
    if value is None:
        return None
    u = unit.lower().replace(" ", "")
    if "мл" in u:
        return value / 1_000_000   # мл/ч → м3/ч
    if "нм3" in u or "nm3" in u:
        return value               # нм3/ч ≈ м3/ч (условно)
    return value                   # уже м3/ч


# ══════════════════════════════════════════════════════════════════
# РЕЕСТР ЛИМИТОВ ICARUS
# Источник: Aspen Icarus Reference Guide V15
# Формат: (min, max) в единицах Icarus SI
# None = нет ограничения
# ══════════════════════════════════════════════════════════════════

ICARUS_LIMITS = {

    # ── Насосы центробежные (CP) ──────────────────────────────────
    # Icarus Ref V15, Chapter 7 — Pumps
    "CP": {
        "CENTRIF": {
            "design_temperature": (None, 450),   # °С
            "pressure_kpa":       (None, None),  # нет верхнего лимита
            "flow_rate_sg":       (0.2,  5.0),   # уд. вес жидкости
        },
        "API 610": {
            "design_temperature": (None, 450),
        },
        "ANSI": {
            "design_temperature": (None, 260),
        },
        "MAG DRIVE": {
            "design_temperature": (None, 230),
            "pressure_kpa":       (None, 1895),
        },
        "ANSI PLAST": {
            "design_temperature": (None, 107),
        },
    },

    # ── Вентиляторы (FN) ─────────────────────────────────────────
    # Icarus Ref V15, Chapter 3 — Compressors, pp. 3-12
    "FN": {
        "CENTRIF": {
            "flow_rate_m3h": (1200, 1_529_000),  # м3/ч
            "pressure_kpa":  (0,    None),        # Па → кПа (нет верхнего лимита — зависит от типа)
        },
        "VANEAXIAL": {
            "flow_rate_m3h": (3950, 76_450),
        },
        "ROT BLOWER": {
            "flow_rate_m3h": (170,  6_700),
            "pressure_kpa":  (15,   100),
        },
    },

    # ── Газовые компрессоры (GC) ──────────────────────────────────
    # Icarus Ref V15, Chapter 3 — Compressors, pp. 3-5
    "GC": {
        "CENTRIF": {
            "flow_rate_m3h":      (102,   509_700),
            "design_temperature": (-125,  90),
            "pressure_kpa":       (None,  34_470),
        },
        "RECIP MOTR": {
            "flow_rate_m3h":      (None,  339_000),
            "design_temperature": (-125,  90),
            "pressure_kpa":       (None,  41_300),
        },
        "CENTRIF IG": {
            "flow_rate_m3h":      (850,   118_900),
            "design_temperature": (0,     90),
            "pressure_kpa":       (None,  1_480),
        },
    },

    # ── Вертикальные сосуды / реакторы / сепараторы (VT) ─────────
    # Icarus Ref V15, Chapter 10 — Vessels, pp. 10-21
    "VT": {
        "CYLINDER": {
            "design_temperature": (None, 340),   # °С, ферритный мат.
        },
        "JACKETED": {
            "design_temperature": (None, 340),
        },
    },

    # ── Теплообменники (HE) ───────────────────────────────────────
    # Icarus Ref V15, Chapter 5 — Heat Transfer, pp. 5-2
    "HE": {
        "FLOAT HEAD": {
            "design_temperature":    (None, 340),   # °С
            "pressure_kpa":          (None, None),
            # tube OD, tube length — проверяем если заданы
        },
        "FIXED T S": {
            "design_temperature":    (None, 340),
        },
        "U TUBE": {
            "design_temperature":    (None, 340),
        },
        "KETTLE REBOIL": {
            "design_temperature":    (None, 340),
        },
    },

    # ── Печи / инсинераторы (FU) ──────────────────────────────────
    # Icarus Ref V15, Chapter 5, pp. 5-41
    "FU": {
        "HEATER": {
            "capacity_kw":           (None, 145_000),  # кВт (145 МВт)
            "pressure_kpa":          (None, 41_000),
            "design_temperature":    (None, 815),
        },
    },

    # ── Колонны (TW) ─────────────────────────────────────────────
    # Icarus Ref V15, Chapter 9 — Towers
    "TW": {
        "TRAYED": {
            "design_temperature": (None, 340),
            "pressure_kpa":       (None, None),
        },
        "PACKED": {
            "design_temperature": (None, 340),
            "pressure_kpa":       (None, None),
        },
    },

    # ── Котлы (STB) ───────────────────────────────────────────────
    # Icarus Ref V15, Chapter 15
    "STB": {
        "BOILER": {
            # capacity_kw обязателен — проверяется в блоке 3
            "pressure_kpa": (None, None),
        },
    },

    # ── Факелы (FLR) ─────────────────────────────────────────────
    "FLR": {
        "SELF SUPP": {
            "length_m": (10, 60),
        },
    },

    # ── Свечи (STK) ──────────────────────────────────────────────
    "STK": {
        "STACK": {
            "length_m": (10, 60),
        },
    },
}

# Обязательные числовые поля для конкретных типов
# Если поле пустое — это ошибка уровня B (не просто пропуск)
REQUIRED_FIELDS = {
    "STB": ["capacity_kw"],   # котёл без мощности — нельзя рассчитать
    "FLR": ["length_m", "diameter_m"],
    "STK": ["length_m", "diameter_m"],
}


# ══════════════════════════════════════════════════════════════════
# БЛОК 1 — ПОЛНОТА (Группа A)
# ══════════════════════════════════════════════════════════════════

def _check_completeness(record, seen_tags: set) -> bool:
    """
    Проверяет наличие обязательных полей.
    Возвращает False если найдена критическая ошибка → ранний выход.
    """
    errors = record.errors

    # A1: тэг присутствует
    if not record.user_tag or not record.user_tag.strip():
        errors.append("[A1] user_tag отсутствует")
        return False   # дальше нет смысла проверять

    # A2: длина тэга
    if len(record.user_tag) > 20:
        errors.append(f"[A2] user_tag '{record.user_tag}' длиннее 20 символов "
                      f"({len(record.user_tag)})")

    # A3: без пробелов
    if re.search(r'\s', record.user_tag):
        errors.append(f"[A3] user_tag '{record.user_tag}' содержит пробел")

    # A4: запрещённые символы
    forbidden = re.search(r'[!@#$%&*()?/{}[\]*=<>,`~^:;|"\'+\\]', record.user_tag)
    if forbidden:
        errors.append(f"[A4] user_tag содержит запрещённый символ: "
                      f"'{forbidden.group()}'")

    # A5: уникальность
    if record.user_tag in seen_tags:
        errors.append(f"[A5] user_tag '{record.user_tag}' дублируется")
    else:
        seen_tags.add(record.user_tag)

    # A6: зона не указана — предупреждение, экспортёр подставит default_area
    if not record.parent_area or not record.parent_area.strip():
        record.warnings.append("[A6] parent_area не указана")

    # A7: тип Icarus определён
    if not record.acce_item_symbol:
        errors.append("[A7] acce_item_symbol не определён — запись не экспортируется")
        return False   # ранний выход — дальше некорректно проверять

    # A8–A13: тип-специфические поля — предупреждения, не ошибки
    sym = record.acce_item_symbol
    if sym == 'CP' and record.flow_rate is None:
        record.warnings.append("[A8] CP: flow_rate не указан")
    if sym in ('CP', 'GC', 'FN') and record.motor_power_kw is None:
        record.warnings.append(f"[A9] {sym}: motor_power_kw не указан")
    if sym in ('GC', 'FU', 'STB') and record.capacity_kw is None:
        record.warnings.append(f"[A10] {sym}: capacity_kw не указан")
    if sym == 'HE' and record.heat_transfer_area_m2 is None:
        record.warnings.append("[A11] HE: heat_transfer_area_m2 не указана")
    if sym in ('VT', 'TW') and all(
        v is None for v in [record.diameter_m, record.length_m, record.volume]
    ):
        record.warnings.append(f"[A12] {sym}: diameter_m/length_m/volume не указаны")
    if (sym == 'CE' and record.acce_item_type == 'CRANE'
            and record.lift_capacity_t is None):
        record.warnings.append("[A13] CE CRANE: lift_capacity_t не указана")

    return True


# ══════════════════════════════════════════════════════════════════
# БЛОК 2 — ДИАПАЗОНЫ ICARUS (Группа B)
# ══════════════════════════════════════════════════════════════════

def _check_range(field_name: str, value: Optional[float],
                 limits: dict, record_errors: list, code: str):
    """Проверяет одно поле по словарю лимитов."""
    if value is None:
        return   # пустое поле пропускаем
    if field_name not in limits:
        return
    lo, hi = limits[field_name]
    if lo is not None and value < lo:
        record_errors.append(
            f"[{code}] {field_name}={value} < минимума {lo} (Icarus V15)"
        )
    if hi is not None and value > hi:
        record_errors.append(
            f"[{code}] {field_name}={value} > максимума {hi} (Icarus V15)"
        )


def _check_ranges(record) -> None:
    """
    Сверяет параметры записи с лимитами из ICARUS_LIMITS.
    Сначала конвертирует единицы в Icarus SI.
    Пустые поля пропускаются.
    """
    sym  = record.acce_item_symbol
    sub  = record.acce_item_type or ""
    errs = record.errors

    # Получаем лимиты для этого типа и подтипа
    type_limits = ICARUS_LIMITS.get(sym, {})
    # Ищем точное совпадение подтипа, иначе берём первый доступный
    limits = type_limits.get(sub) or (
        next(iter(type_limits.values())) if type_limits else {}
    )

    if not limits:
        return   # нет лимитов для этого типа — пропускаем

    # Конвертируем и проверяем
    code = f"B-{sym}"

    _check_range("design_temperature",
                 _to_celsius(getattr(record, 'design_temperature', None)),
                 limits, errs, code)

    _check_range("pressure_kpa",
                 _to_kpa(getattr(record, 'pressure', None),
                         getattr(record, 'pressure_unit', None) or "МПа"),
                 limits, errs, code)

    _check_range("flow_rate_m3h",
                 _to_m3h(getattr(record, 'flow_rate', None),
                         getattr(record, 'flow_rate_unit', None) or ""),
                 limits, errs, code)

    _check_range("capacity_kw",
                 getattr(record, 'capacity_kw', None),
                 limits, errs, code)

    _check_range("length_m",
                 getattr(record, 'length_m', None),
                 limits, errs, code)

    # Рекомендуемые поля — отсутствие является предупреждением, не ошибкой
    required = REQUIRED_FIELDS.get(sym, [])
    for field in required:
        val = getattr(record, field, None)
        if val is None:
            record.warnings.append(f"[W-{sym}] Поле '{field}' не заполнено для {sym}")

    # Для вентиляторов: давление в Па нужно конвертировать отдельно
    # (парсер xlsx хранит его в Pa, не в MPa)
    if sym == "FN" and record.pressure is not None:
        unit = (record.pressure_unit or "").upper()
        if unit in ("PA", "ПА"):
            _check_range("pressure_kpa",
                         record.pressure / 1000,   # Па → кПа
                         limits, errs, code)


# ══════════════════════════════════════════════════════════════════
# БЛОК 3 — ЛОГИКА (Группа C)
# ══════════════════════════════════════════════════════════════════

def _check_logic(record) -> None:
    """Проверяет физическую согласованность полей."""
    errs  = record.errors
    warns = record.warnings

    # C1: рабочая температура ≤ расчётной
    if (getattr(record, 'operating_temperature', None) is not None and
            record.design_temperature is not None):
        if record.operating_temperature > getattr(record, 'design_temperature', record.operating_temperature + 1):
            errs.append(
                f"[C1] Траб={record.operating_temperature}°С > "
                f"Трасч={record.design_temperature}°С — физически невозможно"
            )

    # C2: вес × количество = общий вес (±1%)
    if (getattr(record, 'weight_unit_kg', None) is not None and
            getattr(record, 'weight_total_kg', None) is not None and
            record.quantity and record.quantity > 0):
        expected = record.weight_unit_kg * record.quantity
        if expected > 0:
            delta = abs(expected - record.weight_total_kg) / expected
            if delta > 0.01:
                warns.append(
                    f"[C2] Вес: {record.weight_unit_kg} кг × {record.quantity} шт "
                    f"= {expected:.0f} кг, но указано {record.weight_total_kg:.0f} кг "
                    f"(расхождение {delta*100:.1f}%)"
                )

    # C3: объём из размеров vs заявленный объём (±10%)
    if (getattr(record, 'volume', None) is not None and
            record.diameter_m is not None and
            record.length_m is not None and
            record.volume > 0):
        calc_vol = math.pi / 4 * record.diameter_m ** 2 * record.length_m
        delta = abs(calc_vol - record.volume) / record.volume
        if delta > 0.10:
            errs.append(
                f"[C3] Объём из D={record.diameter_m}м × L={record.length_m}м "
                f"= {calc_vol:.2f}м3, но заявлено {record.volume}м3 "
                f"(расхождение {delta*100:.0f}%, допуск ±10%)"
            )

    # C4: котёл без тепловой мощности — предупреждение
    if record.acce_item_symbol == "STB" and record.capacity_kw is None:
        warns.append("[W4] STB: capacity_kw не заполнено — ACCE использует значение по умолчанию")

    # C5: нет описания
    if not record.description_ru and not record.description_en:
        warns.append("[C5] Отсутствует наименование оборудования")


# ══════════════════════════════════════════════════════════════════
# ГЛАВНАЯ ФУНКЦИЯ
# ══════════════════════════════════════════════════════════════════

def validate_all(records: list) -> list:
    """
    Запускает трёхблочный каскад валидации для всего списка.

    Порядок:
      1. Блок 1 (полнота) — ранний выход при критической ошибке
      2. Блок 2 (диапазоны Icarus) — только если прошёл блок 1
      3. Блок 3 (логика) — всегда, если нет критических ошибок блока 1

    is_valid = True только если errors пуст.
    """
    seen_tags: set[str] = set()

    for record in records:
        # Сбрасываем предыдущие результаты (на случай повторной валидации)
        record.errors   = []
        record.warnings = []
        record.is_valid = False

        # ── Блок 1: Полнота ───────────────────────────────────────
        ok = _check_completeness(record, seen_tags)
        if not ok:
            # Критическая ошибка — дальше не проверяем эту запись
            record.is_valid = False
            continue

        # ── Блок 2: Диапазоны Icarus ─────────────────────────────
        _check_ranges(record)

        # ── Блок 3: Логика ────────────────────────────────────────
        _check_logic(record)

        # ── Финальный флаг ────────────────────────────────────────
        record.is_valid = len(record.errors) == 0

    return records