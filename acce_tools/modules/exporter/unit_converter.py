"""Converts field values from parser units to ACCE units."""

from __future__ import annotations

import logging
import math
from typing import Callable

logger = logging.getLogger(__name__)

# All keys are uppercase so they match the normalised (upper-stripped) units
# that convert() looks up.
CONVERTERS: dict[tuple[str, str], Callable[[float], float]] = {
    # Pressure
    ("MPA",    "PA"):     lambda x: round(x * 1_000_000, 1),
    ("MPA",    "KPAG"):   lambda x: round(x * 1000, 1),
    ("MPA",    "KPA"):    lambda x: round(x * 1000, 1),
    ("MPA",    "BAR"):    lambda x: round(x * 10, 3),
    ("KPA",    "KPAG"):   lambda x: round(x, 1),
    ("BAR",    "KPAG"):   lambda x: round(x * 100, 1),
    ("PA",     "KPAG"):   lambda x: round(x / 1000, 3),

    # Volumetric flow
    ("M3/H",   "L/S"):    lambda x: round(x / 3.6, 4),
    ("M3/H",   "ACMH"):   lambda x: round(x, 2),
    ("M3/H",   "M3/H"):   lambda x: round(x, 2),
    ("M3/S",   "L/S"):    lambda x: round(x * 1000, 4),

    # Mass flow
    ("T/H",    "T/H"):    lambda x: round(x, 2),
    ("T/H",    "L/S"):    lambda x: round(x / 3.6, 4),   # approx for water density 1 t/m3
    ("KG/H",   "L/S"):    lambda x: round(x / 3600, 4),  # approx for water

    # Gas flow (normal conditions)
    ("NM3/H",  "NM3/H"):  lambda x: round(x, 2),

    # Power
    ("KW",     "KW"):     lambda x: round(x, 2),
    ("KW",     "KCAL/H"): lambda x: round(x * 859.845, 1),

    # Heat duty
    ("GCAL/H", "KCAL/H"): lambda x: round(x * 1_000_000, 1),
    ("GCAL/H", "GCAL/H"): lambda x: round(x, 4),

    # Dimensions
    ("M",      "M"):      lambda x: round(x, 4),
    ("MM",     "MM"):     lambda x: round(x, 1),
    ("M2",     "M2"):     lambda x: round(x, 2),
    ("M3",     "M3"):     lambda x: round(x, 3),

    # Mass / lifting capacity
    ("T",      "T"):      lambda x: round(x, 2),
    ("T",      "TONNE"):  lambda x: round(x, 2),    # metric tonne = tonne
    ("TONNE",  "T"):      lambda x: round(x, 2),
    ("KG",     "T"):      lambda x: round(x / 1000, 3),
    ("KG",     "TONNE"):  lambda x: round(x / 1000, 3),

    # Temperature (identity — no conversion needed, just normalise)
    ("DEG C",  "DEG C"):  lambda x: round(x, 1),
    ("C",      "DEG C"):  lambda x: round(x, 1),
}


def convert(
    value: float | None,
    from_unit: str,
    to_unit: str,
) -> float | None:
    """Return *value* converted from *from_unit* to *to_unit*.

    Returns None when value is None or NaN.
    Returns value unchanged when the conversion pair is unknown.
    """
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None

    from_norm = (from_unit or "").strip().upper()
    to_norm = (to_unit or "").strip().upper()

    if not from_norm or not to_norm:
        return value

    if from_norm == to_norm:
        return value

    converter = CONVERTERS.get((from_norm, to_norm))
    if converter is not None:
        return converter(float(value))

    logger.warning(
        "No converter for (%r → %r); writing value as-is", from_unit, to_unit
    )
    return value
