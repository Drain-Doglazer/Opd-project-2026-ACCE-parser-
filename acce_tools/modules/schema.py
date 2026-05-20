"""schema.py — Canonical ACCERecord dataclass shared across all parser modules."""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ACCERecord:
    # ── Source tracing ────────────────────────────────────────────────────────
    source_file:  str = ""
    source_sheet: str = ""
    source_row:   int = 0

    # ── Identity ─────────────────────────────────────────────────────────────
    raw_tag:        str = ""
    user_tag:       str = ""
    description_ru: str = ""
    description_en: str = ""
    parent_area:    str = ""
    quantity:       int = 1
    material:       str = ""

    # ── Weight ───────────────────────────────────────────────────────────────
    weight_unit_kg:  Optional[float] = None
    weight_total_kg: Optional[float] = None
    weight_total_t:  Optional[float] = None

    # ── Technical characteristics ─────────────────────────────────────────────
    tech_raw:              str            = ""
    flow_rate:             Optional[float] = None
    flow_rate_unit:        str            = ""
    pressure:              Optional[float] = None
    pressure_unit:         str            = ""
    volume:                Optional[float] = None
    volume_unit:           str            = ""
    capacity_kw:           Optional[float] = None
    motor_power_kw:        Optional[float] = None
    diameter_m:            Optional[float] = None
    length_m:              Optional[float] = None
    dn_mm:                 Optional[int]   = None
    lift_capacity_t:       Optional[float] = None
    design_temperature:    Optional[float] = None
    operating_temperature: Optional[float] = None
    heat_transfer_area_m2: Optional[float] = None
    heat_duty_gcalh:       Optional[float] = None

    # ── Equipment flags ───────────────────────────────────────────────────────
    vfd:              bool = False
    explosion_proof:  bool = False
    operation_status: str  = "DUTY"

    # ── Classification (NOT filled by parsers — belongs to a classifier step) ─
    acce_item_symbol:      str = ""
    acce_item_type:        str = ""
    classification_source: str = ""  # e.g. "tech:flow_rate+pressure", "kw:ru:насос", "prefix:Н-"

    # ── ACCE import control ───────────────────────────────────────────────────
    action: str = "NEW"

    # ── Validation state ──────────────────────────────────────────────────────
    is_valid: bool = False
    errors:   list = field(default_factory=list)
    warnings: list = field(default_factory=list)
