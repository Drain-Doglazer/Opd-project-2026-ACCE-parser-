"""classifier.py — Classifies equipment type and populates
acce_item_symbol, acce_item_type, and classification_source for each ACCERecord."""

from __future__ import annotations

from .schema import ACCERecord

# ── Symbol classification rules ───────────────────────────────────────────────
# Each entry: (symbol, ru_keywords, en_keywords, tag_prefixes)
# Evaluated in order; first match wins.
_SYMBOL_RULES: list[tuple[str, list[str], list[str], list[str]]] = [
    (
        "CP",
        ["насос", "агрегат насосный", "насосный агрегат"],
        ["pump"],
        ["Н-", "H-", "CP", "P-"],
    ),
    (
        "GC",
        ["компрессор", "газовый компрессор", "нагнетатель"],
        ["compressor", "gas compressor"],
        ["КМ-", "GC", "K-"],
    ),
    (
        "FN",
        ["вентилятор", "дымосос", "воздуходувка", "газодувка"],
        ["fan", "blower", "ventilator", "induced draft", "forced draft"],
        ["FN", "В-", "D-", "ГД-", "ВД-"],
    ),
    # ── FU before HE: "печь" must win over the generic "heater" EN keyword ─────
    (
        "FU",
        ["печь", "огневой нагреватель", "термическое окисление",
         "термоокислитель", "инсинератор", "горелка"],
        ["furnace", "fired heater", "thermal oxidizer", "pyrolysis",
         "incinerator", "burner"],
        ["П-", "FU", "F-", "IC-"],
    ),
    # ── HE before VT so "E-" prefix hits HE, not VT ──────────────────────────
    (
        "HE",
        ["теплообменник", "холодильник", "конденсатор", "испаритель",
         "рекуператор", "нагреватель", "воздушный холодильник",
         "кипятильник", "ребойлер"],
        ["exchanger", "cooler", "condenser", "evaporator",
         "heater", "air cooler", "reboiler", "air cooled"],
        ["Т-", "HE", "E-", "АВО-"],
    ),
    (
        "VT",
        # NOTE: "сепаратор" removed — magnetic/vibration separators are UNKNOWN
        # NOTE: "колонна аппарат" removed — columns go to TW
        ["ёмкость", "емкость", "сосуд", "резервуар", "бак",
         "барабан", "реактор", "автоклав", "силос"],
        ["vessel", "tank", "drum", "reactor", "silo",
         "hopper", "receiver", "autoclave"],
        ["Е-", "V-", "R-", "Р-", "TK-"],
    ),
    (
        "TW",
        ["колонна", "абсорбер", "десорбер", "ректификатор", "скруббер",
         "регенератор"],
        ["tower", "column", "absorber", "stripper", "scrubber",
         "regenerator"],
        ["К-", "KA-", "TW", "Col-", "C-"],
    ),
    (
        "STB",
        ["котёл", "котел", "парогенератор", "бойлер"],
        ["boiler", "steam generator", "waste heat boiler"],
        ["STB", "Б-"],
    ),
    (
        "CE",
        ["кран", "конвейер", "элеватор", "транспортёр", "электроталь"],
        ["crane", "conveyor", "belt conveyor", "elevator", "bucket elevator",
         "chain block", "monorail"],
        ["CE", "К-мост", "К-козл", "GT-", "CV-", "EL-"],
    ),
    (
        "HO",
        ["таль", "тельфер", "лебёдка"],
        ["hoist", "winch"],
        ["HO"],
    ),
    (
        "FLR",
        ["факел", "факельная установка"],
        ["flare", "flare stack"],
        ["FLR"],
    ),
    (
        "STK",
        ["труба дымовая", "дымовая труба"],
        ["stack", "chimney"],
        ["STK"],
    ),
    # ── Explicit UNKNOWN rules — stop false positives ─────────────────────────
    # These match BEFORE the catch-all keyword rules would misfire.
    (
        "UNKNOWN",
        # Magnetic/vibration/vibrating separators are NOT pressure vessels
        ["электромагнитный сепаратор", "магнитный сепаратор",
         "вибрационный грохот", "грохот"],
        ["magnetic separator", "electromagnetic separator",
         "vibrating screen", "vibrating separator"],
        ["MS-", "VS-"],
    ),
    (
        "UNKNOWN",
        ["затвор", "задвижка", "клапан", "заслонка"],
        ["gate", "valve", "damper", "shutter"],
        ["ZSH-", "V-1", "DA-", "FL-"],
    ),
    (
        "UNKNOWN",
        ["питатель", "дозатор"],
        ["feeder", "doser", "dosing"],
        ["FE-"],
    ),
    (
        "UNKNOWN",
        ["дробилка", "мельница", "измельчитель"],
        ["crusher", "mill", "grinder"],
        ["CS-", "ML-"],
    ),
    (
        "UNKNOWN",
        ["пылеуловитель", "циклон", "фильтр рукавный", "фильтр мешочный"],
        ["dust collector", "cyclone", "bag filter", "pulse filter",
         "vacuum filter"],
        ["DC-", "F-"],
    ),
    (
        "UNKNOWN",
        ["желоб", "лоток", "течка"],
        ["chute", "trough"],
        ["CH-"],
    ),
    (
        "UNKNOWN",
        ["весы", "система взвешивания", "взвешивание"],
        ["scale", "weighing", "weigh"],
        ["WT-"],
    ),
    (
        "UNKNOWN",
        ["пробоотборник", "пробозаборник"],
        ["sampler", "sampling"],
        ["SM-"],
    ),
    (
        "UNKNOWN",
        ["смеситель", "мешалка", "миксер"],
        ["mixer", "kneader", "blender"],
        ["MX-"],
    ),
    (
        "UNKNOWN",
        ["пресс", "экструдер"],
        ["press", "extruder"],
        ["EP-"],
    ),
    (
        "UNKNOWN",
        ["тележка", "транспортная тележка", "автоматическая тележка"],
        ["transfer car", "vehicle", "agv", "carrier"],
        ["MP-", "MV-", "MX-"],
    ),
    (
        "UNKNOWN",
        ["упаковочное", "упаковка"],
        ["packaging", "packing machine"],
        ["PM-"],
    ),
]

# ── Type-selection keywords ───────────────────────────────────────────────────
_CP_GEAR_KW:     list[str] = ["шестер", "gear"]
_CP_SCREW_KW:    list[str] = ["винт", "screw"]

_GC_RECIPROC_KW: list[str] = ["поршн", "reciproc"]
_GC_SCREW_KW:    list[str] = ["винт", "screw"]
_GC_AXIAL_KW:    list[str] = ["осев", "axial"]

_FN_AXIAL_KW:    list[str] = ["осев", "axial"]

_VT_SPHERE_KW:   list[str] = ["шар", "sphere"]
_VT_CONE_KW:     list[str] = ["конус", "cone", "силос", "silo"]
_VT_VERT_KW:     list[str] = ["верт", "vert", "реактор", "reactor",
                               "автоклав", "autoclave"]
_VT_HORIZ_KW:    list[str] = ["барабан", "drum", "ёмкость", "емкость",
                               "резервуар", "бак", "vessel", "tank",
                               "receiver", "hopper"]

_HE_AIR_KW:      list[str] = ["возд", "air cool", "аво", "avo"]
_HE_KETTLE_KW:   list[str] = ["кетл", "kettle", "кипятильник", "reboiler",
                               "ребойлер"]
_HE_UTUBE_KW:    list[str] = ["u-tube", "u tube"]
_HE_FIXED_KW:    list[str] = ["фикс", "fixed"]

_TW_PACKED_KW:   list[str] = ["насадк", "pack", "скруббер", "scrubber"]

_STB_VERT_KW:    list[str] = ["верт", "vert"]

_CE_BELT_KW:     list[str] = ["конвей", "belt", "транспортёр"]
_CE_ELEVATOR_KW: list[str] = ["элев", "elevator", "bucket"]
_CE_CRANE_KW:    list[str] = ["кран", "crane", "таль", "chain block",
                               "monorail", "электроталь"]

# Motors above this threshold (kW) are classified as large compressors/turbines.
_GC_MOTOR_KW_THRESHOLD = 500


# ── Private helpers ───────────────────────────────────────────────────────────

def _desc(record: ACCERecord) -> str:
    """Description text only — raw_tag excluded to prevent accidental keyword matches."""
    return f"{record.description_ru} {record.description_en}".lower()


def _has_kw(text: str, keywords: list[str]) -> bool:
    return any(kw.lower() in text for kw in keywords)


# ── Step 1: technical-field priority ─────────────────────────────────────────

def _symbol_by_tech(record: ACCERecord) -> tuple[str, str] | tuple[None, None]:
    if record.heat_transfer_area_m2 is not None:
        return "HE", "tech:heat_transfer_area_m2"
    if record.heat_duty_gcalh is not None:
        return "HE", "tech:heat_duty_gcalh"
    if record.lift_capacity_t is not None:
        return "CE", "tech:lift_capacity_t"
    if record.motor_power_kw is not None and record.motor_power_kw > _GC_MOTOR_KW_THRESHOLD:
        return "GC", "tech:motor_power_kw"
    if record.flow_rate is not None and record.pressure is not None:
        # FIX E4: fan descriptions beat the generic flow+pressure→CP rule
        if not _has_kw(_desc(record), ["вентилятор", "fan", "blower", "дымосос",
                                        "воздуходувка", "induced draft",
                                        "forced draft", "газодувка"]):
            return "CP", "tech:flow_rate+pressure"
    return None, None


# ── Step 2: keyword and prefix matching ──────────────────────────────────────

def _symbol_by_kw(record: ACCERecord) -> tuple[str, str] | tuple[None, None]:
    desc = _desc(record)
    tag  = record.raw_tag.strip().lower()
    for symbol, ru_kw, en_kw, tag_prefixes in _SYMBOL_RULES:
        for kw in ru_kw:
            if kw.lower() in desc:
                return symbol, f"kw:ru:{kw}"
        for kw in en_kw:
            if kw.lower() in desc:
                # FIX E2: "compressor" in vessel-name context (e.g. "Compressor KO Drum")
                # If RU description has vessel terms, the word is context, not type.
                if symbol == "GC" and kw.lower() in ("compressor", "gas compressor"):
                    if _has_kw(record.description_ru.lower(),
                                ["емкость", "ёмкость", "сосуд", "бак",
                                 "барабан", "резервуар", "аккумулятор"]):
                        continue
                return symbol, f"kw:en:{kw}"
        for prefix in tag_prefixes:
            if tag.startswith(prefix.lower()):
                return symbol, f"prefix:{prefix}"
    return None, None


# ── Item-type selection ───────────────────────────────────────────────────────

def _item_type(symbol: str, record: ACCERecord) -> str:
    desc = _desc(record)

    if symbol == "CP":
        if record.pressure is not None and record.pressure > 10:
            return "RECIPROC"
        if _has_kw(desc, _CP_GEAR_KW):
            return "GEAR"
        if _has_kw(desc, _CP_SCREW_KW):
            return "SCREW"
        return "CENTRIF"

    if symbol == "GC":
        if _has_kw(desc, _GC_RECIPROC_KW):
            return "RECIPROC"
        if _has_kw(desc, _GC_SCREW_KW):
            return "SCREW"
        if _has_kw(desc, _GC_AXIAL_KW):
            return "AXIAL"
        return "CENTRIF"

    if symbol == "FN":
        if _has_kw(desc, _FN_AXIAL_KW):
            return "AXIAL"
        return "CENTRIF"

    if symbol == "VT":
        if _has_kw(desc, _VT_SPHERE_KW):
            return "SPHERE"
        if _has_kw(desc, _VT_CONE_KW):
            return "CONE ROOF"
        if _has_kw(desc, _VT_VERT_KW):
            return "VERT DRUM"
        if _has_kw(desc, _VT_HORIZ_KW):
            return "HORIZ DRUM"
        return "HORIZ DRUM"

    if symbol == "HE":
        if _has_kw(desc, _HE_AIR_KW):
            return "AIR COOLED"
        if _has_kw(desc, _HE_KETTLE_KW):
            return "KETTLE"
        if _has_kw(desc, _HE_UTUBE_KW):
            return "U TUBE"
        if _has_kw(desc, _HE_FIXED_KW):
            return "FIXED TS"
        return "FLOAT HEAD"

    if symbol == "TW":
        if _has_kw(desc, _TW_PACKED_KW):
            return "PACKED"
        return "TRAYED"

    if symbol == "STB":
        if _has_kw(desc, _STB_VERT_KW):
            return "VERT"
        return "HORIZ"

    if symbol == "CE":
        if record.lift_capacity_t is not None:
            return "CRANE"
        if _has_kw(desc, _CE_CRANE_KW):
            return "CRANE"
        if _has_kw(desc, _CE_BELT_KW):
            return "BELT"
        if _has_kw(desc, _CE_ELEVATOR_KW):
            return "ELEVATOR"
        return "CRANE"

    # FU, FLR, STK, HO, UNKNOWN — no subtypes
    return ""


# ── Public interface ──────────────────────────────────────────────────────────

def classify(record: ACCERecord) -> ACCERecord:
    """Classifies equipment type and populates
    acce_item_symbol, acce_item_type, and classification_source."""
    symbol, source = _symbol_by_tech(record)
    if symbol is None:
        symbol, source = _symbol_by_kw(record)

    if symbol is None:
        record.acce_item_symbol = "UNKNOWN"
        record.acce_item_type = "UNKNOWN"
        record.classification_source = ""
        return record

    record.acce_item_symbol = symbol
    record.acce_item_type = _item_type(symbol, record)
    record.classification_source = source
    return record