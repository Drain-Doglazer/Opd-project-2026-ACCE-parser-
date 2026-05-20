"""parser_pdf.py — Reads a PDF equipment list and returns raw ACCERecord objects.

Supports three PDF formats:
  1. Equipment List (one A3 page, 23+ columns)
  2. Process Regulations (sections 5.1 / 5.2 / 5.3)
  3. Technology Specifications per equipment type (Samples_8 format)

Single responsibility: extract data from PDF tables and return ACCERecord objects.
Classification (acce_item_symbol / acce_item_type) is NOT performed here.
"""

import re
from typing import Optional
import pdfplumber

from .schema import ACCERecord
from .spec_parser import parse_spec


# ── Transliteration ───────────────────────────────────────────────
_CYR = {
    'А':'A','Б':'B','В':'V','Г':'G','Д':'D','Е':'E','Ё':'Yo','Ж':'Zh',
    'З':'Z','И':'I','Й':'Y','К':'K','Л':'L','М':'M','Н':'N','О':'O',
    'П':'P','Р':'R','С':'S','Т':'T','У':'U','Ф':'F','Х':'Kh','Ц':'Ts',
    'Ч':'Ch','Ш':'Sh','Щ':'Sch','Ъ':'','Ы':'Y','Ь':'','Э':'E','Ю':'Yu','Я':'Ya',
    'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'yo','ж':'zh',
    'з':'z','и':'i','й':'y','к':'k','л':'l','м':'m','н':'n','о':'o',
    'п':'p','р':'r','с':'s','т':'t','у':'u','ф':'f','х':'kh','ц':'ts',
    'ч':'ch','ш':'sh','щ':'sch','ъ':'','ы':'y','ь':'','э':'e','ю':'yu','я':'ya',
}
_FORB = re.compile(r'[!@#$%&*()?/{}[\]*=<>,`~^:;|"\' \\+\s～]')


def _translit(t: str) -> str:
    return ''.join(_CYR.get(c, c) for c in t)


def _make_tag(raw: str) -> str:
    """Creates a valid ACCE user_tag from a raw tag string (max 20 chars)."""
    lat = _translit(raw.strip())
    s = _FORB.sub('-', lat)
    s = re.sub(r'-{2,}', '-', s).strip('-')[:20]
    return s or f"TAG-{raw[:10]}"


# ── Utilities ─────────────────────────────────────────────────────

def _c(cell) -> str:
    if cell is None:
        return ""
    return re.sub(r'\s+', ' ', str(cell).replace('\n', ' ')).strip()


def _num(s) -> Optional[float]:
    if not s:
        return None
    s = str(s).strip()
    neg = bool(re.match(r'минус', s, re.I))
    d = re.sub(r'[^\d,.]', '', s).replace(',', '.')
    if not d:
        return None
    try:
        v = float(d)
        return -v if neg else v
    except Exception:
        return None


def _split_bi(t: str) -> tuple[str, str]:
    if '/' in t:
        p = t.split('/', 1)
        return p[0].strip(), p[1].strip()
    return t.strip(), ""


def _parse_qty(raw: str) -> int:
    raw = _c(raw)
    m = re.search(r'(\d+)\s*(?:блок|бло)\S*\s+по\s+(\d+)', raw, re.I)
    if m:
        return int(m.group(1)) * int(m.group(2))
    m = re.search(r'\d+', raw)
    return int(m.group()) if m else 1


def _maxnum(*vals) -> Optional[float]:
    v = [x for x in vals if x is not None]
    return max(v) if v else None


# ── Equipment List tech extraction (latin notation: Q=, P=, V=, DN, Φ×L) ─────

def _parse_range_first(text: str) -> Optional[float]:
    """Return first number from range strings like '5,000-5,500' or '2,500'.
    Strips comma-as-thousands-separator (X,XXX pattern) before parsing."""
    first = re.split(r'\s*[-–]\s*', text.strip())[0].strip()
    # Remove thousands commas: "5,000" → "5000" (comma followed by exactly 3 digits)
    cleaned = re.sub(r',(\d{3})(?!\d)', r'\1', first)
    cleaned = cleaned.replace(',', '.')  # remaining commas are decimal separators
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def _parse_tech_el(raw: str) -> dict:
    r: dict = {}
    if not raw:
        return r
    # Q in kW → capacity (boilers, fired heaters)
    m = re.search(r'Q\s*=\s*([\d,]+)\s*kW', raw, re.I)
    if m:
        v = _parse_range_first(m.group(1))
        if v is not None:
            r['capacity_kw'] = v
        # continue — also try to parse P below
    else:
        # Q in volumetric / mass flow (only when not kW)
        m = re.search(r'Q\s*=\s*([\d,. -]+?)\s*(m3/h|m³/h|nm3/h|t/h|kg/h)', raw, re.I)
        if m:
            v = _parse_range_first(m.group(1))
            if v is not None:
                r['flow_rate'] = v
                r['flow_rate_unit'] = m.group(2).strip()
    # P → pressure, always normalised to MPa
    m = re.search(r'P\s*=\s*([\d,. -]+?)\s*(MPa|kPa|Pa|bar)\b', raw, re.I)
    if m:
        v = _parse_range_first(m.group(1))
        unit = m.group(2).upper()
        if v is not None:
            if unit == 'PA':
                v = v / 1_000_000
            elif unit == 'KPA':
                v = v / 1_000
            elif unit == 'BAR':
                v = v * 0.1
            # MPa stays as-is
            r['pressure'] = v
            r['pressure_unit'] = 'MPa'
    # V → volume (m3)
    m = re.search(r'V\s*=\s*([\d.,]+)\s*m3', raw, re.I)
    if m:
        v = _parse_range_first(m.group(1))
        if v is not None:
            r['volume'] = v
    # T= → lift capacity (t)
    m = re.search(r'\bT\s*=\s*([\d.]+)\s*t\b', raw, re.I)
    if m:
        r['lift_capacity_t'] = float(m.group(1))
    # DN → nominal diameter
    m = re.search(r'DN\s*(\d+)', raw, re.I)
    if m:
        r['dn_mm'] = int(m.group(1))
    # Φ×L → diameter × length (m)
    m = re.search(r'[ΦΦ]?\s*([\d.]+)\s*[×x]\s*([\d.]+)', raw)
    if m:
        r['diameter_m'] = float(m.group(1))
        r['length_m'] = float(m.group(2))
    return r


# ── Regulations tech extraction (Russian patterns) ────────────────

def _parse_tech_reg(text: str) -> dict:
    r: dict = {}
    t = text.replace('\n', ' ')

    diams = []
    for m in re.finditer(r'D\s*(?:верх|нижн|средн)?[\s.]*(?:части\s*)?=\s*([\d\s]+)\s*мм', t, re.I):
        v = _num(m.group(1).replace(' ', ''))
        if v:
            diams.append(v)
    if diams:
        r['diameter_m'] = max(diams) / 1000

    m = re.search(r'Нцил[\s.]*=\s*([\d\s]+)\s*мм', t, re.I)
    if m:
        r['length_m'] = _num(m.group(1).replace(' ', '')) / 1000

    m = re.search(r'[Пп]оверхност[ьи][\s\w]*теплообмена[\s\S]*?([\d\s]+(?:[,.][\d]+)?)\s*м[²2]', t)
    if m:
        v = _num(m.group(1).replace(' ', ''))
        if v:
            r['heat_transfer_area_m2'] = v

    m = re.search(r'[Тт]епловая нагрузка[\s\S]*?([\d\s]+(?:[,.][\d]+)?)\s*Мкал', t)
    if m:
        v = _num(m.group(1).replace(' ', ''))
        if v:
            r['heat_duty_gcalh'] = v / 1000

    pvals = [_num(m.group(1)) for m in
             re.finditer(r'(?:трубное|межтрубное)\s+пространство\s*[–\-]\s*([\d,]+)', t, re.I)]
    pvals = [p for p in pvals if p]
    if pvals:
        r['pressure'] = max(pvals)
        r['pressure_unit'] = "МПа"

    blk = re.search(r'[Тт]емпература расчетная.*?°?С[:\s]*(.*?)(?:Давление|\Z)', t, re.DOTALL)
    if blk:
        tvals = [_num(m.group(1)) for m in re.finditer(r'([\d,]+)', blk.group(1))]
        tvals = [v for v in tvals if v and 0 < v < 1000]
        if tvals:
            r['design_temperature'] = max(tvals)

    return r


# ── Condition value parser for table 5.1 ─────────────────────────

def _parse_cond_val(s: str) -> Optional[float]:
    """Parses a condition value from table 5.1, handles ranges and OCR artifacts."""
    if not s:
        return None
    s = s.strip()
    if 'рт' in s.lower():
        return None
    parts = re.split(r'\s*[-–]\s*', s)
    vals = []
    for part in parts:
        part = part.strip().lstrip(',').rstrip(',')
        cleaned = re.sub(r'[^\d,.]', '', part)
        if not cleaned:
            continue
        if cleaned.startswith('0') and ',' not in cleaned and '.' not in cleaned and len(cleaned) >= 2:
            cleaned = '0.' + cleaned[1:]
        cleaned = cleaned.replace(',', '.')
        try:
            vals.append(float(cleaned))
        except Exception:
            pass
    return max(vals) if vals else None


def _assign_pt(c16: str, c17: str, c18: str, c19: str) -> tuple[Optional[float], Optional[float]]:
    """
    Separates pressure and temperature from table 5.1 columns.
    Returns (pressure_MPa, temperature_C). Handles vacuum columns (mm Hg).
    """
    if 'рт' in c16.lower() or 'рт' in c17.lower():
        t = _maxnum(_parse_cond_val(c18), _parse_cond_val(c19))
        return None, t

    v16 = _parse_cond_val(c16)
    v17 = _parse_cond_val(c17)
    v18 = _parse_cond_val(c18)
    v19 = _parse_cond_val(c19)

    first = v16 if v16 is not None else v17
    if first is not None and first > 20:
        t = _maxnum(v16, v17)
        p = _maxnum(v18, v19)
    else:
        p = _maxnum(v16, v17)
        t = _maxnum(v18, v19)
    return p, t


def _parse_pv_pair(text: str) -> Optional[float]:
    """Returns max of all numbers found in a 'Top X Bottom Y' string."""
    vals = [_num(n) for n in re.findall(r'[\d,\.]+', text.replace(' ', ''))]
    vals = [v for v in vals if v is not None]
    return max(vals) if vals else None


# ── Format 1: Equipment List ──────────────────────────────────────

def _parse_el_page(page, filepath: str) -> list[ACCERecord]:
    records: list[ACCERecord] = []
    tables = page.extract_tables()
    if not tables:
        return records
    for ri, row in enumerate(tables[0]):
        if len(row) < 14:
            continue
        tag = _c(row[1])
        if not tag or not re.search(r'[\w]+-[\w]+', tag):
            continue
        name = _c(row[2])
        ru, en = _split_bi(name)
        tech = _c(row[3])
        qty = _parse_qty(_c(row[4]))
        wt = _num(_c(row[5]))
        oper = "SPAR" if re.search(r'РЕЗЕРВ|STANDBY|SPAR', _c(row[9]).upper()) else "DUTY"
        pwr_raw = _c(row[10])
        if pwr_raw and pwr_raw.strip() not in ('--', ''):
            _m = re.search(r'(\d[\d.]*)', pwr_raw)
            pwr = float(_m.group(1)) if _m else None
        else:
            pwr = None
        vfd = bool(re.search(r'ДА|YES', _c(row[11]).upper()))
        expl = bool(re.search(r'ДА|YES', _c(row[13]).upper()))
        td = _parse_tech_el(tech)
        rec = ACCERecord(
            source_file=filepath, source_sheet="Equipment List", source_row=ri,
            raw_tag=tag, user_tag=_make_tag(tag),
            description_ru=ru, description_en=en,
            quantity=qty, weight_unit_kg=wt, tech_raw=tech,
            flow_rate=td.get('flow_rate'), flow_rate_unit=td.get('flow_rate_unit', ''),
            pressure=td.get('pressure'), pressure_unit=td.get('pressure_unit', ''),
            volume=td.get('volume'), diameter_m=td.get('diameter_m'),
            length_m=td.get('length_m'), dn_mm=td.get('dn_mm'),
            capacity_kw=td.get('capacity_kw'), motor_power_kw=pwr,
            lift_capacity_t=td.get('lift_capacity_t'),
            vfd=vfd, explosion_proof=expl, operation_status=oper,
            action="NEW", is_valid=True)
        for f, v in parse_spec(tech).items():
            if hasattr(rec, f) and getattr(rec, f) is None:
                setattr(rec, f, v)
        records.append(rec)
    return records


# ── Format 2: Process Regulations (sections 5.1 / 5.2 / 5.3) ─────

def _detect_section(text: str, cur: Optional[str]) -> Optional[str]:
    if re.search(r'5\.1.{0,10}Характеристика.{0,30}ректификацион', text, re.I):
        return "5.1"
    if re.search(r'5\.2.{0,10}Характеристика.{0,30}колонного', text, re.I):
        return "5.2"
    if re.search(r'5\.3.{0,10}Характеристика.{0,30}теплообменного', text, re.I):
        return "5.3"
    return cur


def _row51(row: list, ri: int, fp: str) -> Optional[ACCERecord]:
    if len(row) < 22:
        return None
    tag = _c(row[2])
    if not tag or not re.search(r'[А-ЯЁA-Z]-\d{2,3}', tag):
        return None
    name = _c(row[3])
    qty = _parse_qty(_c(row[4]))
    pressure, temp = _assign_pt(_c(row[16]), _c(row[17]), _c(row[18]), _c(row[19]))
    tech = _c(row[20])
    mat = _c(row[21])
    td = _parse_tech_reg(tech)
    return ACCERecord(
        source_file=fp, source_sheet="5.1", source_row=ri,
        raw_tag=tag, user_tag=_make_tag(tag), description_ru=name,
        quantity=qty, material=mat, tech_raw=tech,
        pressure=pressure, pressure_unit="МПа",
        design_temperature=temp,
        diameter_m=td.get('diameter_m'), length_m=td.get('length_m'),
        action="NEW", is_valid=True)


def _row52(row: list, ri: int, fp: str) -> Optional[ACCERecord]:
    if len(row) < 17:
        return None
    tag = _c(row[2])
    if not tag or not re.search(r'[А-ЯЁA-Z]-\d{2,3}', tag):
        return None
    name = _c(row[3])
    qty = _parse_qty(_c(row[4]))
    tech = _c(row[14])
    mat = _c(row[16])
    td = _parse_tech_reg(tech)
    return ACCERecord(
        source_file=fp, source_sheet="5.2", source_row=ri,
        raw_tag=tag, user_tag=_make_tag(tag), description_ru=name,
        quantity=qty, material=mat, tech_raw=tech,
        pressure=_parse_pv_pair(_c(row[11])), pressure_unit="МПА",
        design_temperature=_parse_pv_pair(_c(row[8])),
        diameter_m=td.get('diameter_m'), length_m=td.get('length_m'),
        action="NEW", is_valid=True)


def _row53(row: list, ri: int, fp: str) -> Optional[ACCERecord]:
    if len(row) < 16:
        return None
    tag = _c(row[2])
    if not tag or not re.search(r'[А-ЯЁ]-\d{3}', tag):
        return None
    name = _c(row[3])
    qty = _parse_qty(_c(row[4]))
    tech = _c(row[15])
    mat = _c(row[16]) if len(row) > 16 else ""
    td = _parse_tech_reg(tech)
    bm = re.search(r'[А-ЯЁ]-\d{3}', tag)
    btag = bm.group() if bm else tag
    w = []
    if not td.get('heat_transfer_area_m2'):
        w.append("Нет площади теплообмена в тех. описании")
    return ACCERecord(
        source_file=fp, source_sheet="5.3", source_row=ri,
        raw_tag=btag, user_tag=_make_tag(btag), description_ru=name,
        quantity=qty, material=mat, tech_raw=tech,
        pressure=td.get('pressure'), pressure_unit="МПа",
        design_temperature=td.get('design_temperature'),
        heat_transfer_area_m2=td.get('heat_transfer_area_m2'),
        heat_duty_gcalh=td.get('heat_duty_gcalh'),
        action="NEW", is_valid=len(w) == 0, warnings=w)


# ── Format 3: Technology Specifications (Samples_8) ──────────────

_SPECS_COL_MAP: dict[str, dict] = {
    'reactors':    {'mat': 9,  'T': 11, 'P': 13, 'D': 6,  'V': 4},
    'vessels':     {'mat': 8,  'T': 10, 'P': 12, 'D': 6,  'V': 4, 'L': 7},
    'hx':          {'mat': 10, 'T': 12, 'P': 14, 'F': 7,  'D': 8},
    'pumps':       {'mat': 10, 'T': -1, 'P':  9},
    'compressors': {'mat': 12, 'T': -1, 'P':  8},
    'misc':        {'mat': 12, 'T': -1, 'P': -1},
}

_SHEET_LABELS: dict[str, str] = {
    'reactors': 'Реакторы', 'vessels': 'Ёмкости',
    'hx': 'Теплообменники', 'pumps': 'Насосы',
    'compressors': 'Компрессоры', 'misc': 'Прочее',
}


def _parse_specs_page(page, filepath: str) -> list[ACCERecord]:
    """Parses one page of the Technology Specifications format."""
    records: list[ACCERecord] = []
    text = page.extract_text() or ''
    tbs  = page.extract_tables()
    if not tbs:
        return records

    table_type = ''
    for ttype, pat in [
        ('reactors',    r'Таблица\s+\d+\s*[–\-]\s*Реактор'),
        ('vessels',     r'Таблица\s+\d+\s*[–\-]\s*Емкостн'),
        ('hx',          r'Таблица\s+\d+\s*[–\-]\s*Теплообменн'),
        ('pumps',       r'Таблица\s+\d+\s*[–\-]\s*Насос'),
        ('compressors', r'Таблица\s+\d+\s*[–\-]\s*Компрессор'),
        ('misc',        r'Таблица\s+\d+\s*[–\-]\s*Проч'),
    ]:
        if re.search(pat, text, re.I):
            table_type = ttype
            break
    if not table_type:
        return records

    table = tbs[0]
    cm    = _SPECS_COL_MAP.get(table_type, {})
    DATA_START = 3
    sheet_label = _SHEET_LABELS.get(table_type, 'Specs')

    def _gcol(row: list, ci: int) -> str:
        if ci < 0 or ci >= len(row):
            return ''
        return _c(row[ci])

    main_row_data = None
    sub_rows: list = []

    def _flush(ri: int, mrow: list, srows: list) -> list[ACCERecord]:
        recs: list[ACCERecord] = []
        tag_raw = _gcol(mrow, 1)
        name    = _gcol(mrow, 2)
        qty_raw = _gcol(mrow, 3)
        if not tag_raw:
            return recs

        tags = (re.findall(r'[А-ЯЁ]{1,2}[а-яё]?-\d+', tag_raw) or
                re.findall(r'[A-Z]{1,4}-\d+', tag_raw))
        if not tags:
            return recs

        qty = _parse_qty(qty_raw) if qty_raw else len(tags)
        mat = _gcol(mrow, cm.get('mat', -1))

        all_rows = [mrow] + srows
        p_vals = [_num(_gcol(r, cm.get('P', -1))) for r in all_rows]
        t_vals = [_num(_gcol(r, cm.get('T', -1))) for r in all_rows]
        p_vals = [v for v in p_vals if v and v > 0]
        t_vals = [v for v in t_vals if v and 0 < v < 1500]
        pressure = max(p_vals) if p_vals else None
        temp     = max(t_vals) if t_vals else None

        volume   = _num(_gcol(mrow, cm.get('V', -1)))
        d_raw    = _gcol(mrow, cm.get('D', -1))
        d_val    = _num(d_raw)
        diameter = d_val / 1000 if d_val and d_val > 10 else d_val
        area     = _num(_gcol(mrow, cm.get('F', -1)))

        for tag in tags:
            rec = ACCERecord(
                source_file=filepath, source_sheet=sheet_label,
                source_row=ri, raw_tag=tag, user_tag=_make_tag(tag),
                description_ru=name, quantity=1, material=mat,
                pressure=pressure, pressure_unit='МПа',
                design_temperature=temp,
                volume=volume, diameter_m=diameter,
                heat_transfer_area_m2=area,
                action='NEW', is_valid=True)
            recs.append(rec)
        return recs

    for ri, row in enumerate(table):
        if ri < DATA_START:
            continue
        col0 = _c(row[0]) if row else ''
        is_main = bool(re.match(r'^\d+$', col0.strip()))
        if is_main:
            if main_row_data is not None:
                records.extend(_flush(main_row_data[0], main_row_data[1], sub_rows))
            main_row_data = (ri, row)
            sub_rows = []
        else:
            if main_row_data is not None:
                sub_rows.append(row)

    if main_row_data is not None:
        records.extend(_flush(main_row_data[0], main_row_data[1], sub_rows))

    return records


# ── PUBLIC API ────────────────────────────────────────────────────

def parse(file_path: str) -> list[ACCERecord]:
    """
    Reads a PDF equipment list and returns a list of ACCERecord objects.
    Auto-detects format: Equipment List, Process Regulations, or Tech Specifications.
    """
    if not __import__('os').path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    records: list[ACCERecord] = []

    with pdfplumber.open(file_path) as pdf:
        text0 = pdf.pages[0].extract_text() or ""
        tbs0  = pdf.pages[0].extract_tables()
        ncols = len(tbs0[0][0]) if tbs0 else 0

        is_el = (ncols >= 20 and
                 not re.search(r'5\.[123]\s+Характеристика', text0, re.I))

        all_text = ' '.join(p.extract_text() or '' for p in pdf.pages[:3])
        is_specs = bool(re.search(
            r'Технологические\s+спецификации\s+позиционного', all_text, re.I))

        if is_el:
            for page in pdf.pages:
                records.extend(_parse_el_page(page, file_path))

        elif is_specs:
            seen: set[str] = set()
            for page in pdf.pages:
                for rec in _parse_specs_page(page, file_path):
                    key = f"{rec.source_sheet}:{rec.raw_tag}"
                    if key in seen:
                        continue
                    seen.add(key)
                    records.append(rec)

        else:
            sec: Optional[str] = None
            seen = set()
            for pi, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                tbs  = page.extract_tables()
                if not tbs:
                    continue
                sec = _detect_section(text, sec)
                fn = {'5.1': _row51, '5.2': _row52, '5.3': _row53}.get(sec)
                if not fn:
                    continue
                for ri, row in enumerate(tbs[0]):
                    rec = fn(row, ri, file_path)
                    if rec is None:
                        continue
                    key = f"{sec}:{rec.raw_tag}"
                    if key in seen:
                        continue
                    seen.add(key)
                    records.append(rec)

    return records
