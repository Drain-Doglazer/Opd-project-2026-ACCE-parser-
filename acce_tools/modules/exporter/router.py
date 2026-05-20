"""Routes each ACCERecord to the correct ACCE sheet name."""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)

# (acce_item_symbol, acce_item_type) → sheet name as it appears in master_template.xlsx
ROUTE_MAP: dict[tuple[str, str], str] = {
    # ── Pumps ─────────────────────────────────────────────────────────────────
    ('CP', 'CENTRIF'):      'DCP GEN SERV',
    ('CP', 'RECIPROC'):     'DCP GEN SERV',
    ('CP', 'GEAR'):         'DCP GEN SERV',
    ('CP', 'SCREW'):        'DCP GEN SERV',
    ('CP', 'API610'):       'DCP API 610',
    ('CP', 'API 610'):      'DCP API 610',

    # ── Compressors ───────────────────────────────────────────────────────────
    ('GC', 'CENTRIF'):      'DGC CENTRIF',
    ('GC', 'RECIPROC'):     'DGC CENTRIF',
    ('GC', 'SCREW'):        'DGC CENTRIF',
    ('GC', 'AXIAL'):        'DGC CENTRIF',

    # ── Fans ──────────────────────────────────────────────────────────────────
    ('FN', 'CENTRIF'):      'EFN CENTRIF',
    ('FN', 'AXIAL'):        'EFN CENTRIF',

    # ── Heat exchangers ───────────────────────────────────────────────────────
    ('HE', 'FLOAT HEAD'):   'DHE FLOAT HEAD',
    ('HE', 'FIXED TS'):     'DHE FLOAT HEAD',
    ('HE', 'U TUBE'):       'DHE FLOAT HEAD',
    ('HE', 'KETTLE'):       'DHE FLOAT HEAD',
    ('HE', 'AIR COOLED'):   'DHE AIR COOLER',

    # ── Vessels ───────────────────────────────────────────────────────────────
    ('VT', 'HORIZ DRUM'):   'DHT HORIZ DRUM',
    ('VT', 'VERT DRUM'):    'DVT CYLINDER',
    ('VT', 'SPHERE'):       'DVT CYLINDER',
    ('VT', 'CONE ROOF'):    'DVT CYLINDER',
    ('VT', 'FLT ROOF'):     'DVT CYLINDER',

    # ── Towers ────────────────────────────────────────────────────────────────
    ('TW', 'TRAYED'):       'DTW TRAYED',
    ('TW', 'PACKED'):       'DTW PACKED',

    # ── Fired heaters ─────────────────────────────────────────────────────────
    ('FU', 'FIREBOX'):      'EFU BOX',
    ('FU', 'PYROLYSIS'):    'EFU BOX',
    ('FU', ''):             'EFU BOX',

    # ── Boilers ───────────────────────────────────────────────────────────────
    ('STB', 'HORIZ'):       'ESTBBOILER',
    ('STB', 'VERT'):        'ESTBBOILER',
    ('STB', ''):            'ESTBBOILER',

    # ── Cranes / Conveyors ────────────────────────────────────────────────────
    ('CE', 'CRANE'):        'ECE BRIDGE CRN',
    ('CE', 'BELT'):         'DCO OPEN BELT',
    ('CE', 'ELEVATOR'):     'DCO OPEN BELT',
    ('CE', ''):             'ECE BRIDGE CRN',

    # ── Hoists ────────────────────────────────────────────────────────────────
    ('HO', ''):             'ECE BRIDGE CRN',
}


SYMBOL_DEFAULTS: dict[str, str] = {
    'CP':  'CENTRIF',
    'GC':  'CENTRIF',
    'FN':  'CENTRIF',
    'VT':  'HORIZ DRUM',
    'HE':  'FLOAT HEAD',
    'TW':  'TRAYED',
    'STB': 'HORIZ',
    'FU':  'FIREBOX',
    'CE':  'CRANE',
    'HO':  '',
}


def route(
    symbol: str,
    item_type: str,
    available_sheets: set[str],
) -> tuple[str | None, str]:
    """Return *(sheet_name, routing_note)* for the given symbol + type.

    sheet_name is None when no suitable sheet can be found.
    """
    sym = (symbol or '').strip().upper()
    typ = (item_type or '').strip().upper()

    if not sym or sym == 'UNKNOWN':
        return None, 'symbol UNKNOWN'

    # Exact match
    key = (sym, typ)
    if key in ROUTE_MAP:
        sheet = ROUTE_MAP[key]
        if sheet in available_sheets:
            return sheet, f'exact:{key}'
        return None, f"sheet '{sheet}' not in template"

    # Fallback to SYMBOL_DEFAULTS for this symbol
    default_type = SYMBOL_DEFAULTS.get(sym, '')
    fallback_key = (sym, default_type)
    if fallback_key in ROUTE_MAP:
        sheet = ROUTE_MAP[fallback_key]
        if sheet in available_sheets:
            return sheet, f'fallback:{fallback_key}'
        return None, f"sheet '{sheet}' not in template"

    return None, f'no route for ({sym!r}, {typ!r})'
