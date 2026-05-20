"""Universal mapping from ACCERecord fields to ACCE SlotNames."""

# Maps internal field name → ACCE slot name (row 4 of each sheet).
# During export, the slot is silently skipped if it doesn't exist in
# the target sheet's col_index — so this map works for all equipment types.
UNIVERSAL_MAP: dict[str, str] = {
    # Identity fields — present in all sheets
    "user_tag":               "CpUserTag",
    "parent_area":            "ParentArea",
    "action":                 "Action",
    "description_en":         "CpItemDescription",
    "quantity":               "CpNumberItems",
    "operation_status":       "CpApplication",

    # Process conditions
    "design_temperature":     "CpDesignTemperature",
    "operating_temperature":  "CpOperatingTemperature",
    "pressure":               "CpDesignGaugePressure",

    # Geometry
    "diameter_m":             "CpVesselDiameter",
    "length_m":               "CpStraightSideLength",
    "volume":                 "CpLiquidVolume",
    "dn_mm":                  "CpNominalDiameter",

    # Power
    "motor_power_kw":         "CpDriverPower",
    "capacity_kw":            "CpCapacity",

    # Heat transfer
    "heat_transfer_area_m2":  "CpHeatTransferArea",
    "heat_duty_gcalh":        "CpHeatDuty",

    # Flow — pumps/vessels: CpLiquidFlowrate; fans: CpActualGasFlowrate
    "flow_rate":              "CpLiquidFlowrate",
    "flow_rate_fan":          "CpActualGasFlowrate",

    # Pressure overrides — wrong slot silently skipped per sheet
    "pressure_fan":           "CpFanOutletGaugePressure",
    "pressure_steam":         "CpSteamGaugePressure",

    # Mechanical handling
    "lift_capacity_t":        "CpCraneCapacity",

    # Equipment flags (written when slot exists in sheet)
    "weight_unit_kg":         "CpWeight",
    "vfd":                    "CpVFD",
    "explosion_proof":        "CpExplosionProof",
}

# Maps UNIVERSAL_MAP key → actual column name in the parser output xlsx.
# Needed because writer.py uses shorter header names for some fields.
RECORD_COLUMN_MAP: dict[str, str] = {
    "description_en":         "DESCRIPTION",
    "diameter_m":             "DIAMETER",
    "length_m":               "LENGTH",
    "heat_transfer_area_m2":  "HEAT_TRANSFER_AREA",
    "heat_duty_gcalh":        "HEAT_DUTY",
    "weight_unit_kg":         "WEIGHT_KG",
    "flow_rate_fan":          "FLOW_RATE",
    "pressure_fan":           "PRESSURE",
    "pressure_steam":         "PRESSURE",
    "vfd":                    "VFD",
    "explosion_proof":        "EXPLOSION_PROOF",
}

# Slots that are set explicitly in the mandatory-fields block of the exporter;
# they are skipped in the generic UNIVERSAL_MAP iteration to avoid overwriting.
MANDATORY_SLOTS: frozenset[str] = frozenset(
    {"Action", "CpUserTag", "CpItemDescription", "ParentArea", "CpNumberItems", "CpApplication"}
)
