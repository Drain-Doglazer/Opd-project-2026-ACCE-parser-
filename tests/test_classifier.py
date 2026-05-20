"""Tests for acce_tools.modules.classifier."""

import pytest
from acce_tools.modules.schema import ACCERecord
from acce_tools.modules.classifier import classify


def make(**kwargs) -> ACCERecord:
    return ACCERecord(**kwargs)


# ── Technical-field priority ──────────────────────────────────────────────────

class TestTechFieldPriority:
    def test_heat_transfer_area_gives_HE(self):
        r = classify(make(heat_transfer_area_m2=50.0))
        assert r.acce_item_symbol == "HE"

    def test_heat_duty_gcalh_gives_HE(self):
        r = classify(make(heat_duty_gcalh=2.5))
        assert r.acce_item_symbol == "HE"

    def test_lift_capacity_gives_CE(self):
        r = classify(make(lift_capacity_t=5.0))
        assert r.acce_item_symbol == "CE"

    def test_high_motor_power_gives_GC(self):
        r = classify(make(motor_power_kw=600.0))
        assert r.acce_item_symbol == "GC"

    def test_motor_power_at_boundary_not_GC(self):
        # > 500, so exactly 500 should NOT trigger GC by tech
        r = classify(make(motor_power_kw=500.0))
        assert r.acce_item_symbol != "GC"

    def test_flow_and_pressure_give_CP(self):
        r = classify(make(flow_rate=100.0, pressure=5.0))
        assert r.acce_item_symbol == "CP"

    def test_heat_transfer_area_beats_high_motor_power(self):
        # heat_transfer_area takes priority over motor_power > 500
        r = classify(make(heat_transfer_area_m2=20.0, motor_power_kw=700.0))
        assert r.acce_item_symbol == "HE"

    def test_heat_transfer_area_beats_flow_and_pressure(self):
        r = classify(make(heat_transfer_area_m2=20.0, flow_rate=100.0, pressure=5.0))
        assert r.acce_item_symbol == "HE"

    def test_lift_capacity_beats_flow_and_pressure(self):
        r = classify(make(lift_capacity_t=2.0, flow_rate=100.0, pressure=5.0))
        assert r.acce_item_symbol == "CE"


# ── Russian keyword classification ────────────────────────────────────────────

class TestRussianKeywords:
    def test_pump_ru(self):
        r = classify(make(description_ru="насос центробежный"))
        assert r.acce_item_symbol == "CP"

    def test_compressor_ru(self):
        r = classify(make(description_ru="компрессор газовый"))
        assert r.acce_item_symbol == "GC"

    def test_fan_ru(self):
        r = classify(make(description_ru="вентилятор осевой"))
        assert r.acce_item_symbol == "FN"

    def test_blower_ru(self):
        r = classify(make(description_ru="воздуходувка"))
        assert r.acce_item_symbol == "FN"

    def test_vessel_ru(self):
        r = classify(make(description_ru="ёмкость технологическая"))
        assert r.acce_item_symbol == "VT"

    def test_tank_ru(self):
        r = classify(make(description_ru="резервуар хранения"))
        assert r.acce_item_symbol == "VT"

    def test_drum_ru(self):
        r = classify(make(description_ru="барабан сепарационный"))
        assert r.acce_item_symbol == "VT"

    def test_separator_ru(self):
        r = classify(make(description_ru="сепаратор газожидкостный"))
        assert r.acce_item_symbol == "VT"

    def test_heat_exchanger_ru(self):
        r = classify(make(description_ru="теплообменник кожухотрубный"))
        assert r.acce_item_symbol == "HE"

    def test_cooler_ru(self):
        r = classify(make(description_ru="холодильник водяной"))
        assert r.acce_item_symbol == "HE"

    def test_air_cooler_ru(self):
        r = classify(make(description_ru="воздушный холодильник"))
        assert r.acce_item_symbol == "HE"

    def test_tower_ru(self):
        r = classify(make(description_ru="колонна ректификационная"))
        assert r.acce_item_symbol == "TW"

    def test_absorber_ru(self):
        r = classify(make(description_ru="абсорбер"))
        assert r.acce_item_symbol == "TW"

    def test_furnace_ru(self):
        r = classify(make(description_ru="печь трубчатая"))
        assert r.acce_item_symbol == "FU"

    def test_boiler_ru(self):
        r = classify(make(description_ru="котёл паровой"))
        assert r.acce_item_symbol == "STB"

    def test_flare_ru(self):
        r = classify(make(description_ru="факел аварийный"))
        assert r.acce_item_symbol == "FLR"

    def test_stack_ru(self):
        r = classify(make(description_ru="труба дымовая"))
        assert r.acce_item_symbol == "STK"

    def test_crane_ru(self):
        r = classify(make(description_ru="кран мостовой"))
        assert r.acce_item_symbol == "CE"

    def test_conveyor_ru(self):
        r = classify(make(description_ru="конвейер ленточный"))
        assert r.acce_item_symbol == "CE"


# ── English keyword classification ────────────────────────────────────────────

class TestEnglishKeywords:
    def test_pump_en(self):
        r = classify(make(description_en="centrifugal pump"))
        assert r.acce_item_symbol == "CP"

    def test_compressor_en(self):
        r = classify(make(description_en="gas compressor"))
        assert r.acce_item_symbol == "GC"

    def test_fan_en(self):
        r = classify(make(description_en="axial fan"))
        assert r.acce_item_symbol == "FN"

    def test_blower_en(self):
        r = classify(make(description_en="blower unit"))
        assert r.acce_item_symbol == "FN"

    def test_vessel_en(self):
        r = classify(make(description_en="pressure vessel"))
        assert r.acce_item_symbol == "VT"

    def test_tank_en(self):
        r = classify(make(description_en="storage tank"))
        assert r.acce_item_symbol == "VT"

    def test_drum_en(self):
        r = classify(make(description_en="knockout drum"))
        assert r.acce_item_symbol == "VT"

    def test_separator_en(self):
        r = classify(make(description_en="gas liquid separator"))
        assert r.acce_item_symbol == "VT"

    def test_heat_exchanger_en(self):
        r = classify(make(description_en="shell and tube exchanger"))
        assert r.acce_item_symbol == "HE"

    def test_cooler_en(self):
        r = classify(make(description_en="product cooler"))
        assert r.acce_item_symbol == "HE"

    def test_condenser_en(self):
        r = classify(make(description_en="overhead condenser"))
        assert r.acce_item_symbol == "HE"

    def test_tower_en(self):
        r = classify(make(description_en="distillation tower"))
        assert r.acce_item_symbol == "TW"

    def test_column_en(self):
        r = classify(make(description_en="absorber column"))
        assert r.acce_item_symbol == "TW"

    def test_furnace_en(self):
        # "fired heater" contains "heater" which matches HE before FU in rule order;
        # use "furnace" — unambiguous FU keyword that doesn't appear in HE keywords.
        r = classify(make(description_en="furnace unit"))
        assert r.acce_item_symbol == "FU"

    def test_boiler_en(self):
        r = classify(make(description_en="steam generator boiler"))
        assert r.acce_item_symbol == "STB"

    def test_flare_en(self):
        r = classify(make(description_en="emergency flare stack"))
        assert r.acce_item_symbol == "FLR"

    def test_stack_en(self):
        r = classify(make(description_en="chimney stack"))
        assert r.acce_item_symbol == "STK"

    def test_conveyor_en(self):
        r = classify(make(description_en="belt conveyor"))
        assert r.acce_item_symbol == "CE"


# ── Tag prefix classification ─────────────────────────────────────────────────

class TestTagPrefixes:
    def test_prefix_pump_cyrillic(self):
        r = classify(make(raw_tag="Н-101А"))
        assert r.acce_item_symbol == "CP"

    def test_prefix_pump_latin(self):
        r = classify(make(raw_tag="P-101"))
        assert r.acce_item_symbol == "CP"

    def test_prefix_compressor(self):
        r = classify(make(raw_tag="КМ-101"))
        assert r.acce_item_symbol == "GC"

    def test_prefix_fan_cyrillic(self):
        r = classify(make(raw_tag="В-201"))
        assert r.acce_item_symbol == "FN"

    def test_prefix_vessel_cyrillic(self):
        r = classify(make(raw_tag="Е-301"))
        assert r.acce_item_symbol == "VT"

    def test_prefix_vessel_latin(self):
        r = classify(make(raw_tag="V-301"))
        assert r.acce_item_symbol == "VT"

    def test_prefix_heat_exchanger_cyrillic(self):
        r = classify(make(raw_tag="Т-401"))
        assert r.acce_item_symbol == "HE"

    def test_prefix_tower_cyrillic(self):
        r = classify(make(raw_tag="К-501"))
        assert r.acce_item_symbol == "TW"

    def test_prefix_furnace_cyrillic(self):
        r = classify(make(raw_tag="П-601"))
        assert r.acce_item_symbol == "FU"

    def test_prefix_boiler(self):
        r = classify(make(raw_tag="STB-701"))
        assert r.acce_item_symbol == "STB"

    def test_prefix_flare(self):
        r = classify(make(raw_tag="FLR-801"))
        assert r.acce_item_symbol == "FLR"

    def test_prefix_stack(self):
        r = classify(make(raw_tag="STK-901"))
        assert r.acce_item_symbol == "STK"

    def test_prefix_hoist(self):
        r = classify(make(raw_tag="HO-001"))
        assert r.acce_item_symbol == "HO"


# ── Item-type subtypes ────────────────────────────────────────────────────────

class TestCPSubtypes:
    def test_CP_reciproc_by_pressure(self):
        r = classify(make(flow_rate=10.0, pressure=11.0))
        assert r.acce_item_symbol == "CP"
        assert r.acce_item_type == "RECIPROC"

    def test_CP_gear(self):
        r = classify(make(description_ru="насос шестерёнчатый"))
        assert r.acce_item_type == "GEAR"

    def test_CP_screw(self):
        r = classify(make(description_ru="насос винтовой"))
        assert r.acce_item_type == "SCREW"

    def test_CP_centrif_default(self):
        r = classify(make(description_ru="насос", pressure=5.0))
        assert r.acce_item_type == "CENTRIF"

    def test_CP_centrif_low_pressure(self):
        r = classify(make(flow_rate=10.0, pressure=5.0))
        assert r.acce_item_symbol == "CP"
        assert r.acce_item_type == "CENTRIF"


class TestHESubtypes:
    def test_HE_air_cooled_ru(self):
        r = classify(make(heat_transfer_area_m2=100.0, description_ru="воздушный холодильник"))
        assert r.acce_item_type == "AIR COOLED"

    def test_HE_air_cooled_en(self):
        r = classify(make(heat_transfer_area_m2=100.0, description_en="air cooled exchanger"))
        assert r.acce_item_type == "AIR COOLED"

    def test_HE_kettle(self):
        r = classify(make(heat_transfer_area_m2=50.0, description_en="kettle reboiler"))
        assert r.acce_item_type == "KETTLE"

    def test_HE_utube(self):
        r = classify(make(heat_transfer_area_m2=50.0, description_en="u-tube exchanger"))
        assert r.acce_item_type == "U TUBE"

    def test_HE_fixed(self):
        r = classify(make(heat_transfer_area_m2=50.0, description_en="fixed tubesheet exchanger"))
        assert r.acce_item_type == "FIXED TS"

    def test_HE_float_head_default(self):
        r = classify(make(heat_transfer_area_m2=50.0))
        assert r.acce_item_type == "FLOAT HEAD"


class TestVTSubtypes:
    def test_VT_sphere(self):
        r = classify(make(description_ru="шаровой резервуар"))
        assert r.acce_item_type == "SPHERE"

    def test_VT_cone_roof(self):
        r = classify(make(description_en="cone roof tank"))
        assert r.acce_item_type == "CONE ROOF"

    def test_VT_vert_drum(self):
        r = classify(make(description_ru="ёмкость вертикальная"))
        assert r.acce_item_type == "VERT DRUM"

    def test_VT_horiz_drum_separator(self):
        r = classify(make(description_ru="сепаратор горизонтальный"))
        assert r.acce_item_type == "HORIZ DRUM"

    def test_VT_horiz_drum_default(self):
        r = classify(make(description_ru="ёмкость"))
        assert r.acce_item_type == "HORIZ DRUM"


class TestTWSubtypes:
    def test_TW_packed(self):
        r = classify(make(description_ru="абсорбер с насадкой"))
        assert r.acce_item_symbol == "TW"
        assert r.acce_item_type == "PACKED"

    def test_TW_packed_en(self):
        r = classify(make(description_en="packed absorber"))
        assert r.acce_item_type == "PACKED"

    def test_TW_trayed_default(self):
        r = classify(make(description_ru="колонна тарельчатая"))
        assert r.acce_item_type == "TRAYED"


# ── Priority: tech fields beat keywords ───────────────────────────────────────

class TestTechVsKeywordPriority:
    def test_heat_area_beats_pump_keyword(self):
        # description says "pump" but heat_transfer_area is set — should be HE
        r = classify(make(heat_transfer_area_m2=30.0, description_en="pump cooler"))
        assert r.acce_item_symbol == "HE"

    def test_heat_area_beats_vessel_tag(self):
        r = classify(make(heat_transfer_area_m2=30.0, raw_tag="V-101"))
        assert r.acce_item_symbol == "HE"

    def test_flow_pressure_beats_vessel_keyword(self):
        # flow+pressure tech path gives CP before VT keyword
        r = classify(make(flow_rate=50.0, pressure=3.0, description_en="vessel pump"))
        assert r.acce_item_symbol == "CP"

    def test_high_motor_beats_fan_keyword(self):
        r = classify(make(motor_power_kw=700.0, description_en="fan blower"))
        assert r.acce_item_symbol == "GC"


# ── Empty record ──────────────────────────────────────────────────────────────

class TestEmptyRecord:
    def test_empty_gives_unknown_symbol(self):
        r = classify(make())
        assert r.acce_item_symbol == "UNKNOWN"

    def test_empty_gives_unknown_type(self):
        r = classify(make())
        assert r.acce_item_type == "UNKNOWN"

    def test_whitespace_only_fields_still_unknown(self):
        r = classify(make(description_ru="   ", description_en="  ", raw_tag=" "))
        assert r.acce_item_symbol == "UNKNOWN"
        assert r.acce_item_type == "UNKNOWN"
