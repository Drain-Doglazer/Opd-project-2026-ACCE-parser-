"""spec_parser.py — Parses ТЕХ.ХАРАКТЕРИСТИКИ / SPEC column
into ACCERecord numeric field values."""

import re


def parse_spec(spec: str) -> dict:
    """
    Parses a SPEC cell from equipment list into field dict.

    Examples of inputs found in real files:
      'GZGF503 Q=20t/h'          → flow_rate=20, unit='t/h'
      'XP2-200 Q=40000 m3/h'     → flow_rate=40000, unit='m3/h'
      'CD1-6D Q=3t'              → lift_capacity_t=3.0
      '∅1.8m V=10m3'             → diameter_m=1.8, volume=10.0
      'N-TGD315 C=46.75m Q=20t/h'→ length_m=46.75, flow_rate=20
      'DN250'                    → dn_mm=250.0
      'DN200 Q=5t/h'             → dn_mm=200.0, flow_rate=5.0
      'Q=4000kW P=0.8MPa'        → capacity_kw=4000, pressure=0.8
      'Q=160m3/h P=0.8MPa'       → flow_rate=160, pressure=0.8
      'Q=5000-5500m3/h P=2500Pa' → flow_rate=5000, pressure=0.0025MPa
    """
    if not spec or str(spec).strip() in ('--', 'nan', '', '-', 'None'):
        return {}

    s   = str(spec).strip()
    out = {}

    # 1. Flow rate t/h
    m = re.search(r'Q\s*=\s*([\d,.\s]+)\s*t/h', s, re.IGNORECASE)
    if m:
        out['flow_rate']      = _num(m.group(1))
        out['flow_rate_unit'] = 't/h'

    # 2. Flow rate m3/h  (overwrites t/h if both present)
    m = re.search(r'Q\s*=\s*([\d,.\s-]+?)\s*(?:nm)?m3/h', s, re.IGNORECASE)
    if m:
        # Range like "5000-5500" → take first number
        val = re.split(r'[-–]', m.group(1))[0]
        unit = 'nm3/h' if 'nm3' in s.lower() else 'm3/h'
        out['flow_rate']      = _num(val)
        out['flow_rate_unit'] = unit

    # 3. Lift capacity Q=3t  (no /h — negative lookahead)
    if 'flow_rate' not in out:
        m = re.search(r'Q\s*=\s*([\d.]+)\s*t\b(?!\s*/h)', s, re.IGNORECASE)
        if m:
            out['lift_capacity_t'] = float(m.group(1))

    # 4. Capacity in kW (boilers, heaters)  Q=4000kW
    m = re.search(r'Q\s*=\s*([\d.]+)\s*kW', s, re.IGNORECASE)
    if m:
        out['capacity_kw'] = float(m.group(1))

    # 5. Volume  V=10m3
    m = re.search(r'V\s*=\s*([\d.]+)\s*m3', s, re.IGNORECASE)
    if m:
        out['volume'] = float(m.group(1))

    # 6. Diameter  ∅1.8m  or  Φ1.8m
    m = re.search(r'[∅ØΦφ]\s*([\d.]+)\s*m\b', s)
    if m:
        out['diameter_m'] = float(m.group(1))

    # 7. Conveyor length  C=46.75m
    m = re.search(r'\bC\s*=\s*([\d.]+)\s*m\b', s, re.IGNORECASE)
    if m:
        out['length_m'] = float(m.group(1))

    # 8. Nominal diameter  DN250
    m = re.search(r'\bDN\s*(\d+)', s, re.IGNORECASE)
    if m:
        out['dn_mm'] = float(m.group(1))

    # 9. Pressure MPa
    m = re.search(r'P\s*=\s*([\d.]+)\s*MPa', s, re.IGNORECASE)
    if m:
        out['pressure']      = float(m.group(1))
        out['pressure_unit'] = 'MPa'

    # 10. Pressure Pa  (convert to MPa)
    m = re.search(r'P\s*=\s*([\d,.\s-]+?)\s*Pa\b', s, re.IGNORECASE)
    if m and 'pressure' not in out:
        val = _num(re.split(r'[-–]', m.group(1))[0])
        out['pressure']      = round(val / 1_000_000, 8)
        out['pressure_unit'] = 'MPa'

    return out


def _num(s: str) -> float:
    """'5,000' or '5 000' → 5000.0"""
    return float(str(s).replace(',', '').replace(' ', '').strip())
