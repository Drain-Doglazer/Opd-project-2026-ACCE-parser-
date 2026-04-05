from dataclasses import dataclass, field
from typing import Optional

@dataclass
class ACCERecord:
    # ── Source tracing ────────────────────────────────────────────────────────
    source_file:  str = ""   # path to the original xlsx
    source_sheet: str = ""   # sheet name
    source_row:   int = 0    # row number in the sheet (for error messages)

    # ── Identity fields (from the equipment list) ─────────────────────────────
    seq_number:       int    = 0    # № п/п
    raw_tag:          str    = ""   # ТАГОВЫЙ НОМЕР ОБОРУДОВАНИЯ (original)
    description_ru:   str    = ""   # НАИМЕНОВАНИЕ ОБОРУДОВАНИЯ (Russian)
    description_en:   str    = ""   # English part extracted from bilingual cell
    price_code:       str    = ""   # код ед.расценки  e.g. "6831", "6880.12"
    quantity:         int    = 1    # КОЛ.
    weight_unit_kg:   Optional[float] = None   # ВЕС ЕДИНИЦЫ (кг)
    weight_total_kg:  Optional[float] = None   # Всего масса, кг
    weight_total_t:   Optional[float] = None   # Всего масса, ТОННЫ
    title:            str    = ""   # ТИТУЛ (area / contract section)

    # ── Technical characteristics (parsed from ТЕХ. ХАРАКТЕРИСТИКИ) ──────────
    tech_raw:         str    = ""           # full original string, kept for audit
    flow_rate:        Optional[float] = None  # Q value
    flow_rate_unit:   str    = ""
    pressure:         Optional[float] = None  # P value
    pressure_unit:    str    = ""
    volume:           Optional[float] = None  # V value
    volume_unit:      str    = ""
    capacity_kw:      Optional[float] = None  # Q when unit is kW (thermal duty)
    diameter_m:       Optional[float] = None  # Φ dimension or DN
    length_m:         Optional[float] = None  # second dimension in Φ×L
    dn_mm:            Optional[int]   = None  # DN nominal diameter
    lift_capacity_t:  Optional[float] = None  # T= for hoists/cranes

    # ── ACCE import control fields (filled by Module 3 — Standardiser) ────────
    action:           str    = "NEW"
    user_tag:         str    = ""   # generated unique tag for ACCE
    parent_area:      str    = ""   # mapped from title
    acce_equip_type:  str    = ""   # mapped from price_code to ACCE model code

    # ── Validation state (filled by Module 2 — Validator) ────────────────────
    is_valid:         bool   = False
    errors:           list   = field(default_factory=list)
    warnings:         list   = field(default_factory=list)