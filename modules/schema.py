from dataclasses import dataclass, field
from typing import Optional, List


@dataclass
class ACCERecord:
    # --- Источник данных ---
    source_file: str = ""
    source_sheet: str = ""
    source_row: int = 0

    # --- Идентификация ---
    seq_number: int = 0
    raw_tag: str = ""  # Исходный тег (напр. "Р-001")
    user_tag: str = ""  # Очищенный тег для ACCE (транслит, без пробелов, макс 20 симв)
    description_ru: str = ""  # Описание на русском
    description_en: str = ""  # Описание на английском (если есть)

    # --- Классификация ---
    price_code: str = ""  # Код расценки (напр. "6852")
    acce_item_symbol: str = ""  # Символ Icarus (CP, VT, HE, FN, STB, FU, CE...)
    acce_item_type: str = ""  # Подтип (CENTRIF, CYLINDER...)

    # --- Общие параметры ---
    quantity: int = 1
    weight_unit_kg: Optional[float] = None
    material: str = ""  # Материал (напр. "12ХМ", "Carbon Steel")

    # --- Технические характеристики ---
    tech_raw: str = ""  # Исходная строка характеристик

    # Потоковые
    flow_rate: Optional[float] = None
    flow_rate_unit: str = ""  # м3/ч, нм3/ч

    # Давление и температура
    pressure: Optional[float] = None
    pressure_unit: str = "МПа"  # Важно для конвертации! По умолчанию МПа
    design_temperature: Optional[float] = None  # °C
    operating_temperature: Optional[float] = None  # °C

    # Геометрия
    volume: Optional[float] = None  # м3
    diameter_m: Optional[float] = None  # м
    length_m: Optional[float] = None  # м
    height_m: Optional[float] = None  # м (для свечей/факелов)

    # Электрика / Тепло
    motor_power_kw: Optional[float] = None  # кВт
    heat_transfer_area_m2: Optional[float] = None  # м2
    heat_duty_gcalh: Optional[float] = None  # Гкал/ч
    capacity_kw: Optional[float] = None  # кВт (для котлов)

    # --- Данные для ACCE Import ---
    action: str = "NEW"
    parent_area: str = ""  # Название области (Area Name)

    # --- Валидация ---
    is_valid: bool = False
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)