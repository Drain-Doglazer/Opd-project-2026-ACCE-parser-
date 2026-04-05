# modules/schema.py
from dataclasses import dataclass, field
from typing import Optional, List


@dataclass
class ACCERecord:
    # ── Source tracing ────────────────────────────────────────────────────────
    source_file: str = ""  # path to the original xlsx/docx/pdf
    source_sheet: str = ""  # sheet name or page number
    source_row: int = 0  # row number in the source file

    # ── Identity fields ───────────────────────────────────────────────────────
    seq_number: int = 0  # № п/п
    raw_tag: str = ""  # Исходный тег (напр. "Р-001")
    user_tag: str = ""  # Очищенный тег для ACCE (транслитерация, макс 20 симв)
    description_ru: str = ""  # Наименование на русском
    description_en: str = ""  # Наименование на английском (если есть)

    # ── Classification ────────────────────────────────────────────────────────
    price_code: str = ""  # Код расценки (напр. "6852") - используется парсером
    acce_item_symbol: str = ""  # Символ типа оборудования Icarus (CP, VT, HE...) - ВАЖНО для экспортера
    acce_item_type: str = ""  # Подтип (CENTRIF, CYLINDER...) - ВАЖНО для экспортера

    # ── Common Parameters ─────────────────────────────────────────────────────
    quantity: int = 1
    weight_unit_kg: Optional[float] = None
    material: str = ""  # Материал конструкции (напр. "12ХМ") - ВАЖНО для экспортера

    # ── Technical Characteristics ─────────────────────────────────────────────
    tech_raw: str = ""  # Исходная строка характеристик

    # Flow / Capacity
    flow_rate: Optional[float] = None
    flow_rate_unit: str = ""  # м3/ч, нм3/ч и т.д.
    capacity_kw: Optional[float] = None  # Тепловая нагрузка или мощность (кВт)

    # Pressure & Temperature
    pressure: Optional[float] = None
    pressure_unit: str = "МПа"  # По умолчанию МПа
    design_temperature: Optional[float] = None  # Расчетная температура (°C) - ВАЖНО для экспортера
    operating_temperature: Optional[float] = None  # Рабочая температура (°C)

    # Geometry
    volume: Optional[float] = None  # м3
    diameter_m: Optional[float] = None  # м
    length_m: Optional[float] = None  # м
    height_m: Optional[float] = None  # м (для свечей/факелов)
    dn_mm: Optional[int] = None  # DN номинальный диаметр

    # Electrical / Heat Exchange
    motor_power_kw: Optional[float] = None  # Мощность электродвигателя (кВт) - ВАЖНО для экспортера
    heat_transfer_area_m2: Optional[float] = None  # Площадь теплообмена (м2) - ВАЖНО для экспортера
    lift_capacity_t: Optional[float] = None  # Грузоподъемность (т)

    # ── ACCE Import Control Fields ────────────────────────────────────────────
    action: str = "NEW"
    parent_area: str = ""  # Название области (Area Name)

    # ── Validation State ──────────────────────────────────────────────────────
    is_valid: bool = False
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)