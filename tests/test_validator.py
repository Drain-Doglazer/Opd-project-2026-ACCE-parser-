"""Tests for acce_tools.validator.Validator."""

import pytest
from acce_tools.modules.schema import ACCERecord
from acce_tools.validator.Validator import validate_all


def rec(**kwargs) -> ACCERecord:
    """Minimal valid record; keyword args override defaults."""
    defaults = dict(
        user_tag="P-101",
        parent_area="U-100",
        acce_item_symbol="CP",
        acce_item_type="CENTRIF",
        description_ru="насос",
    )
    defaults.update(kwargs)
    return ACCERecord(**defaults)


# ── Block A: completeness ─────────────────────────────────────────────────────

class TestBlockA:
    def test_valid_record_passes(self):
        r = validate_all([rec()])[0]
        assert r.is_valid
        assert r.errors == []

    def test_A1_missing_tag(self):
        r = validate_all([rec(user_tag="")])[0]
        assert not r.is_valid
        assert any("A1" in e for e in r.errors)

    def test_A1_whitespace_only_tag(self):
        r = validate_all([rec(user_tag="   ")])[0]
        assert not r.is_valid
        assert any("A1" in e for e in r.errors)

    def test_A2_tag_too_long(self):
        r = validate_all([rec(user_tag="A" * 21)])[0]
        assert not r.is_valid
        assert any("A2" in e for e in r.errors)

    def test_A2_tag_exactly_20_ok(self):
        r = validate_all([rec(user_tag="A" * 20)])[0]
        assert r.is_valid

    def test_A3_tag_has_space(self):
        r = validate_all([rec(user_tag="P 101")])[0]
        assert not r.is_valid
        assert any("A3" in e for e in r.errors)

    def test_A4_tag_forbidden_char(self):
        r = validate_all([rec(user_tag="P/101")])[0]
        assert not r.is_valid
        assert any("A4" in e for e in r.errors)

    def test_A5_first_duplicate_ok_second_not(self):
        r1 = rec(user_tag="P-101")
        r2 = rec(user_tag="P-101")
        validate_all([r1, r2])
        assert r1.is_valid
        assert not r2.is_valid
        assert any("A5" in e for e in r2.errors)

    def test_A5_unique_tags_both_ok(self):
        r1 = rec(user_tag="P-101")
        r2 = rec(user_tag="P-102")
        validate_all([r1, r2])
        assert r1.is_valid
        assert r2.is_valid

    def test_A6_missing_area(self):
        r = validate_all([rec(parent_area="")])[0]
        assert not r.is_valid
        assert any("A6" in e for e in r.errors)

    def test_A7_missing_symbol_causes_early_exit(self):
        r = validate_all([rec(acce_item_symbol="")])[0]
        assert not r.is_valid
        assert any("A7" in e for e in r.errors)
        # Early exit: blocks B and C must not run
        assert not any("B-" in e for e in r.errors)
        assert not any(e.startswith("[C") for e in r.errors)


# ── Block B: Icarus ranges ────────────────────────────────────────────────────

class TestBlockB:
    def test_CP_ANSI_temp_within_limit_ok(self):
        r = validate_all([rec(acce_item_type="ANSI", design_temperature=260.0)])[0]
        assert r.is_valid

    def test_CP_ANSI_temp_over_limit(self):
        r = validate_all([rec(acce_item_type="ANSI", design_temperature=261.0)])[0]
        assert not r.is_valid
        assert any("B-CP" in e for e in r.errors)

    def test_CP_ANSI_PLAST_temp_over_limit(self):
        r = validate_all([rec(acce_item_type="ANSI PLAST", design_temperature=110.0)])[0]
        assert not r.is_valid
        assert any("B-CP" in e for e in r.errors)

    def test_CP_MAG_DRIVE_pressure_over_limit(self):
        # 3 MPa = 3000 kPa, limit is 1895 kPa
        r = validate_all([rec(acce_item_type="MAG DRIVE",
                               pressure=3.0, pressure_unit="МПа")])[0]
        assert not r.is_valid
        assert any("B-CP" in e for e in r.errors)

    def test_CP_MAG_DRIVE_pressure_within_limit_ok(self):
        # 1.8 MPa = 1800 kPa < 1895 kPa
        r = validate_all([rec(acce_item_type="MAG DRIVE",
                               pressure=1.8, pressure_unit="МПа")])[0]
        assert r.is_valid

    def test_FN_CENTRIF_flow_too_low(self):
        r = ACCERecord(
            user_tag="FN-101", parent_area="U-100",
            acce_item_symbol="FN", acce_item_type="CENTRIF",
            description_ru="вентилятор",
            flow_rate=100.0, flow_rate_unit="м3/ч",
        )
        validate_all([r])
        assert not r.is_valid
        assert any("B-FN" in e for e in r.errors)

    def test_FN_CENTRIF_flow_within_range_ok(self):
        r = ACCERecord(
            user_tag="FN-102", parent_area="U-100",
            acce_item_symbol="FN", acce_item_type="CENTRIF",
            description_ru="вентилятор",
            flow_rate=5000.0, flow_rate_unit="м3/ч",
        )
        validate_all([r])
        assert r.is_valid

    def test_STB_missing_required_capacity_kw(self):
        r = ACCERecord(
            user_tag="STB-101", parent_area="U-100",
            acce_item_symbol="STB", acce_item_type="BOILER",
            description_ru="котёл",
        )
        validate_all([r])
        assert not r.is_valid
        assert any("B-STB" in e for e in r.errors)

    def test_none_values_skip_range_check(self):
        # No pressure or temp set → no B errors
        r = validate_all([rec(acce_item_type="CENTRIF")])[0]
        assert r.is_valid
        assert not any("B-" in e for e in r.errors)

    def test_unknown_symbol_no_limits_ok(self):
        # HO has no Icarus limits defined → Block B skips quietly
        r = ACCERecord(
            user_tag="HO-101", parent_area="U-100",
            acce_item_symbol="HO", acce_item_type="",
            description_ru="таль",
        )
        validate_all([r])
        assert r.is_valid


# ── Block C: logic checks ─────────────────────────────────────────────────────

class TestBlockC:
    def test_C1_op_temp_exceeds_design(self):
        r = validate_all([rec(
            design_temperature=100.0,
            operating_temperature=150.0,
        )])[0]
        assert not r.is_valid
        assert any("C1" in e for e in r.errors)

    def test_C1_op_temp_equal_to_design_ok(self):
        r = validate_all([rec(
            design_temperature=100.0,
            operating_temperature=100.0,
        )])[0]
        assert r.is_valid

    def test_C1_op_temp_below_design_ok(self):
        r = validate_all([rec(
            design_temperature=200.0,
            operating_temperature=150.0,
        )])[0]
        assert r.is_valid

    def test_C2_weight_mismatch_is_warning(self):
        # 1000 kg * 2 = 2000, but 3000 given → 50% delta > 1% threshold
        r = validate_all([rec(
            weight_unit_kg=1000.0,
            weight_total_kg=3000.0,
            quantity=2,
        )])[0]
        assert r.is_valid  # C2 is a warning, not an error
        assert any("C2" in w for w in r.warnings)

    def test_C2_weight_match_no_warning(self):
        r = validate_all([rec(
            weight_unit_kg=1000.0,
            weight_total_kg=2000.0,
            quantity=2,
        )])[0]
        assert not any("C2" in w for w in r.warnings)

    def test_C3_volume_inconsistent_with_dimensions(self):
        # pi/4 * 1^2 * 1 ≈ 0.785 m3, claimed 5 m3 → 84% delta > 10% threshold
        r = validate_all([rec(
            diameter_m=1.0,
            length_m=1.0,
            volume=5.0,
        )])[0]
        assert not r.is_valid
        assert any("C3" in e for e in r.errors)

    def test_C3_volume_consistent_with_dimensions_ok(self):
        import math
        diam, length = 1.0, 2.0
        vol = math.pi / 4 * diam ** 2 * length  # ≈ 1.571
        r = validate_all([rec(diameter_m=diam, length_m=length, volume=vol)])[0]
        assert r.is_valid

    def test_C4_STB_without_capacity_kw(self):
        r = ACCERecord(
            user_tag="STB-201", parent_area="U-200",
            acce_item_symbol="STB", acce_item_type="BOILER",
            description_ru="котёл",
        )
        validate_all([r])
        assert not r.is_valid
        assert any("C4" in e for e in r.errors)

    def test_C4_STB_with_capacity_kw_ok(self):
        r = ACCERecord(
            user_tag="STB-202", parent_area="U-200",
            acce_item_symbol="STB", acce_item_type="BOILER",
            description_ru="котёл",
            capacity_kw=5000.0,
        )
        validate_all([r])
        assert r.is_valid

    def test_C5_no_description_is_warning(self):
        r = validate_all([rec(description_ru="", description_en="")])[0]
        assert r.is_valid  # warning only, not an error
        assert any("C5" in w for w in r.warnings)

    def test_C5_ru_description_suppresses_warning(self):
        r = validate_all([rec(description_ru="насос", description_en="")])[0]
        assert not any("C5" in w for w in r.warnings)


# ── Cascade and reset behaviour ───────────────────────────────────────────────

class TestCascade:
    def test_seen_tags_reset_between_validate_all_calls(self):
        r = rec(user_tag="P-101")
        validate_all([r])
        assert r.is_valid
        # Second independent call — seen_tags is fresh, no duplicate
        validate_all([r])
        assert r.is_valid

    def test_errors_cleared_on_re_validation(self):
        r = rec(user_tag="")
        validate_all([r])
        assert not r.is_valid
        r.user_tag = "P-101"
        validate_all([r])
        assert r.is_valid
        assert r.errors == []

    def test_multiple_errors_accumulated(self):
        # Long tag (A2) + space in tag (A3) — both should appear
        r = validate_all([rec(user_tag="P " + "A" * 20)])[0]
        codes = [e for e in r.errors if "A2" in e or "A3" in e]
        assert len(codes) == 2
