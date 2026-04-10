"""
parser_pdf.py — Универсальный парсер PDF с перечнями оборудования.

Поддерживает два формата:
  1. Equipment List (Пример_1): одна страница A3, 23 колонки
  2. Технологический регламент (Пример_2): 29 страниц, разделы 5.1/5.2/5.3
"""

import re
from dataclasses import dataclass, field
from typing import Optional
import pdfplumber

# ── Транслитерация ────────────────────────────────────────────────
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

def _translit(t): return ''.join(_CYR.get(c, c) for c in t)
def _make_tag(raw):
    lat = _translit(raw.strip())
    s = _FORB.sub('-', lat)
    s = re.sub(r'-{2,}','-',s).strip('-')[:20]
    return s or f"TAG-{raw[:10]}"

# ── Схема данных ──────────────────────────────────────────────────
@dataclass
class ACCERecord:
    source_file:str=""; source_sheet:str=""; source_row:int=0
    raw_tag:str=""; user_tag:str=""
    description_ru:str=""; description_en:str=""
    parent_area:str=""; quantity:int=1; material:str=""
    weight_unit_kg:Optional[float]=None
    tech_raw:str=""
    design_temperature:Optional[float]=None
    operating_temperature:Optional[float]=None
    pressure:Optional[float]=None; pressure_unit:str=""
    volume:Optional[float]=None
    diameter_m:Optional[float]=None; length_m:Optional[float]=None
    dn_mm:Optional[int]=None
    flow_rate:Optional[float]=None; flow_rate_unit:str=""
    capacity_kw:Optional[float]=None
    motor_power_kw:Optional[float]=None
    lift_capacity_t:Optional[float]=None
    heat_transfer_area_m2:Optional[float]=None
    heat_duty_gcalh:Optional[float]=None
    vfd:bool=False; explosion_proof:bool=False; operation_status:str="DUTY"
    acce_item_symbol:str=""; acce_item_type:str=""
    action:str="NEW"; is_valid:bool=False
    errors:list=field(default_factory=list)
    warnings:list=field(default_factory=list)

# ── Утилиты ───────────────────────────────────────────────────────
def _c(cell)->str:
    if cell is None: return ""
    return re.sub(r'\s+',' ',str(cell).replace('\n',' ')).strip()

def _num(s)->Optional[float]:
    if not s: return None
    s=str(s).strip()
    neg=bool(re.match(r'минус',s,re.I))
    d=re.sub(r'[^\d,.]','',s).replace(',','.')
    if not d: return None
    try: v=float(d); return -v if neg else v
    except: return None

def _split_bi(t):
    if '/' in t:
        p=t.split('/',1); return p[0].strip(),p[1].strip()
    return t.strip(),""

def _parse_qty(raw)->int:
    raw=_c(raw)
    m=re.search(r'(\d+)\s*(?:блок|бло)\S*\s+по\s+(\d+)',raw,re.I)
    if m: return int(m.group(1))*int(m.group(2))
    m=re.search(r'\d+',raw)
    return int(m.group()) if m else 1

def _maxnum(*vals):
    v=[x for x in vals if x is not None]
    return max(v) if v else None

# ── Классификация Equipment List (та же логика что xlsx-парсер) ───
def _pu(tech):
    m=re.search(r'P\s*=\s*[\d,. -]+?\s*(MPa|kPa|Pa)\b',tech,re.I)
    return m.group(1).upper() if m else ""

def _classify_el(tech,en,ru):
    desc=(en+" "+ru).lower()
    hq=bool(re.search(r'(?<![A-Za-z])Q\s*=',tech,re.I))
    hp=bool(re.search(r'(?<![A-Za-z])P\s*=',tech,re.I))
    hv=bool(re.search(r'(?<![A-Za-z])V\s*=',tech,re.I))
    hdn=bool(re.search(r'\bDN\s*\d+',tech,re.I))
    hphi=bool(re.search(r'[ΦΦ]\s*[\d.]+\s*[×x]',tech))
    ht=bool(re.search(r'\bT\s*=\s*[\d.]+\s*t\b',tech,re.I))
    hh=bool(re.search(r'\bH\s*=\s*[\d.]+\s*m\b',tech,re.I))
    qkw=bool(re.search(r'Q\s*=\s*[\d,]+\s*kW',tech,re.I))
    qm3=bool(re.search(r'Q\s*=\s*[\d,. -]+?\s*m3/h',tech,re.I))
    qth=bool(re.search(r'Q\s*=\s*[\d,. -]+?\s*t/h',tech,re.I))
    pu=_pu(tech)
    if ht and hh: return ("CE","BRIDGE CRN") if re.search(r'crane|кран мост|bridge',desc) else ("HO","HOIST")
    if hphi: return "FU","HEATER"
    if qkw:  return "STB","BOILER"
    if hv and not hq and not hp: return "VT","CYLINDER"
    if hdn and not hq and not hp: return "VT","CYLINDER"
    if qm3 and pu=="PA": return "FN","CENTRIF"
    if (qm3 or qth) and pu in("MPA","KPA"):
        if re.search(r'pump|насос',desc): return "CP","CENTRIF"
        if re.search(r'fan|blower|вентилятор',desc): return "FN","CENTRIF"
        if re.search(r'compressor|компрессор',desc): return "GC","CENTRIF"
        if re.search(r'deaerat|деаэратор',desc): return "VT","CYLINDER"
        if re.search(r'soften|умягч',desc): return "VT","CYLINDER"
        if qth: return "VT","CYLINDER"
        return "CP","CENTRIF"
    if hq and not hp:
        if re.search(r'fan|blower|вентилятор',desc): return "FN","CENTRIF"
        if re.search(r'burner|горелка',desc): return "FU","HEATER"
        if re.search(r'pump|насос',desc): return "CP","CENTRIF"
    for pat,sym,sub in [
        (r'boiler|котел|котёл',"STB","BOILER"),
        (r'incinerator|инсинератор',"FU","HEATER"),
        (r'furnace|печь',"FU","HEATER"),
        (r'pump|насос',"CP","CENTRIF"),
        (r'fan|blower|вентилятор',"FN","CENTRIF"),
        (r'compressor|компрессор',"GC","CENTRIF"),
        (r'hoist|таль|подъёмник|подъемник',"HO","HOIST"),
        (r'crane|кран',"CE","BRIDGE CRN"),
        (r'tank|vessel|резервуар|ёмкость|бак',"VT","CYLINDER"),
    ]:
        if re.search(pat,desc): return sym,sub
    return "",""

def _parse_tech_el(raw)->dict:
    r={}
    if not raw: return r
    m=re.search(r'Q\s*=\s*([\d,]+)\s*kW',raw,re.I)
    if m: r['capacity_kw']=_num(m.group(1)); return r
    m=re.search(r'Q\s*=\s*([\d,. -]+?)\s*(m3/h|t/h)',raw,re.I)
    if m: r['flow_rate']=_num(m.group(1)); r['flow_rate_unit']=m.group(2)
    m=re.search(r'P\s*=\s*([\d,. -]+?)\s*(MPa|kPa|Pa)\b',raw,re.I)
    if m: r['pressure']=_num(m.group(1)); r['pressure_unit']=m.group(2)
    m=re.search(r'V\s*=\s*([\d.,]+)\s*m3',raw,re.I)
    if m: r['volume']=_num(m.group(1))
    m=re.search(r'DN\s*(\d+)',raw,re.I)
    if m: r['dn_mm']=int(m.group(1))
    m=re.search(r'[ΦΦ]?\s*([\d.]+)\s*[×x]\s*([\d.]+)',raw)
    if m: r['diameter_m']=float(m.group(1)); r['length_m']=float(m.group(2))
    m=re.search(r'\bT\s*=\s*([\d.]+)\s*t\b',raw)
    if m: r['lift_capacity_t']=float(m.group(1))
    return r

# ── Парсинг тех. описания из регламента (русские паттерны) ────────
def _parse_tech_reg(text)->dict:
    r={}
    t=text.replace('\n',' ')
    # Диаметр — максимальный из D верх / D низ / D=
    diams=[]
    for m in re.finditer(r'D\s*(?:верх|нижн|средн)?[\s.]*(?:части\s*)?=\s*([\d\s]+)\s*мм',t,re.I):
        v=_num(m.group(1).replace(' ',''))
        if v: diams.append(v)
    if diams: r['diameter_m']=max(diams)/1000
    # Высота цил. части
    m=re.search(r'Нцил[\s.]*=\s*([\d\s]+)\s*мм',t,re.I)
    if m: r['length_m']=_num(m.group(1).replace(' ',''))/1000
    # Поверхность теплообмена
    m=re.search(r'[Пп]оверхност[ьи][\s\w]*теплообмена[\s\S]*?([\d\s]+(?:[,.][\d]+)?)\s*м[²2]',t)
    if m:
        v=_num(m.group(1).replace(' ',''))
        if v: r['heat_transfer_area_m2']=v
    # Тепловая нагрузка Мкал/ч → Гкал/ч
    m=re.search(r'[Тт]епловая нагрузка[\s\S]*?([\d\s]+(?:[,.][\d]+)?)\s*Мкал',t)
    if m:
        v=_num(m.group(1).replace(' ',''))
        if v: r['heat_duty_gcalh']=v/1000
    # Давление расчётное — max из трубного/межтрубного
    pvals=[_num(m.group(1)) for m in
           re.finditer(r'(?:трубное|межтрубное)\s+пространство\s*[–\-]\s*([\d,]+)',t,re.I)]
    pvals=[p for p in pvals if p]
    if pvals: r['pressure']=max(pvals); r['pressure_unit']="МПа"
    # Температура расчётная — max
    blk=re.search(r'[Тт]емпература расчетная.*?°?С[:\s]*(.*?)(?:Давление|\Z)',t,re.DOTALL)
    if blk:
        tvals=[_num(m.group(1)) for m in re.finditer(r'([\d,]+)',blk.group(1))]
        tvals=[v for v in tvals if v and 0<v<1000]
        if tvals: r['design_temperature']=max(tvals)
    return r

def _column_subtype(tech)->str:
    t=tech.lower()
    if re.search(r'тарел[кь]|trayed',t): return "TRAYED"
    if re.search(r'насадк|packed|кольц',t): return "PACKED"
    return "TRAYED"

def _he_subtype(tech)->str:
    m=re.search(r'\(([A-Z]{3,4})\)',tech)
    code=m.group(1) if m else ""
    return {"AES":"FLOAT HEAD","AJS":"FLOAT HEAD","AHS":"FLOAT HEAD",
            "AGS":"FLOAT HEAD","AJU":"U TUBE","AEU":"U TUBE",
            "AKT":"KETTLE REBOIL","AKS":"KETTLE REBOIL",
            "BEM":"FIXED T S","BKT":"FIXED T S"}.get(code,"FLOAT HEAD")

def _parse_pv_pair(text)->Optional[float]:
    """Из 'Верх X Низ Y' берём max."""
    vals=[_num(n) for n in re.findall(r'[\d,\.]+',text.replace(' ',''))]
    vals=[v for v in vals if v is not None]
    return max(vals) if vals else None

def _parse_cond_val(s: str) -> Optional[float]:
    """
    Парсит значение условий работы из таблицы 5.1.
    Обрабатывает: '0,14', '014', ',350', '296 - 350', '20 мм. рт.ст.'
    Для диапазона 'X - Y' берёт максимум.
    Для '014' (потеряна запятая) → 0.14 (0XX → 0.XX).
    """
    if not s: return None
    s = s.strip()
    # Пропускаем мм. рт.ст.
    if 'рт' in s.lower(): return None
    # Если строка содержит '-' — диапазон, берём максимум
    parts = re.split(r'\s*[-–]\s*', s)
    vals = []
    for part in parts:
        part = part.strip().lstrip(',').rstrip(',')
        cleaned = re.sub(r'[^\d,.]', '', part)
        if not cleaned: continue
        # '014' без запятой и с ведущим нулём перед цифрами — скорее всего 0.14
        if cleaned.startswith('0') and ',' not in cleaned and '.' not in cleaned and len(cleaned) >= 2:
            cleaned = '0.' + cleaned[1:]
        cleaned = cleaned.replace(',', '.')
        try: vals.append(float(cleaned))
        except: pass
    return max(vals) if vals else None

def _assign_pt(c16:str, c17:str, c18:str, c19:str):
    """
    Умное разделение давления и температуры из таблицы 5.1.
    В зависимости от страницы PDF порядок колонок:
      - (давление верха, давление куба, температура верха, температура куба)
      - (температура верха, температура куба, давление верха, давление куба)
    Определяем по величине: давление < 20 МПа, температура > 40 °С.
    Вакуумные колонны (мм. рт.ст.): возвращаем только температуру.

    Возвращает (pressure_МПа, temp_°С).
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
        # Температуры первыми, давления вторыми
        t = _maxnum(v16, v17)
        p = _maxnum(v18, v19)
    else:
        # Стандарт: давления первыми
        p = _maxnum(v16, v17)
        t = _maxnum(v18, v19)
    return p, t

# ── Формат 1: Equipment List ──────────────────────────────────────
def _parse_el_page(page,filepath)->tuple:
    records,warns=[],[]
    tables=page.extract_tables()
    if not tables: return records,warns
    for ri,row in enumerate(tables[0]):
        if len(row)<14: continue
        tag=_c(row[1])
        if not tag or not re.search(r'[\w]+-[\w]+',tag): continue
        name=_c(row[2]); ru,en=_split_bi(name)
        tech=_c(row[3])
        qty=_parse_qty(_c(row[4]))
        wt=_num(_c(row[5]))
        oper="SPAR" if re.search(r'РЕЗЕРВ|STANDBY|SPAR',_c(row[9]).upper()) else "DUTY"
        pwr_raw=_c(row[10])
        pwr=_num(pwr_raw) if pwr_raw and pwr_raw!='--' else None
        vfd=bool(re.search(r'ДА|YES',_c(row[11]).upper()))
        expl=bool(re.search(r'ДА|YES',_c(row[13]).upper()))
        td=_parse_tech_el(tech)
        sym,sub=_classify_el(tech,en,ru)
        w=[]
        if not sym: w.append(f"Не определён тип Icarus: '{en or ru}'")
        rec=ACCERecord(
            source_file=filepath,source_sheet="Equipment List",source_row=ri,
            raw_tag=tag,user_tag=_make_tag(tag),
            description_ru=ru,description_en=en,
            quantity=qty,weight_unit_kg=wt,tech_raw=tech,
            flow_rate=td.get('flow_rate'),flow_rate_unit=td.get('flow_rate_unit',''),
            pressure=td.get('pressure'),pressure_unit=td.get('pressure_unit',''),
            volume=td.get('volume'),diameter_m=td.get('diameter_m'),
            length_m=td.get('length_m'),dn_mm=td.get('dn_mm'),
            capacity_kw=td.get('capacity_kw'),motor_power_kw=pwr,
            lift_capacity_t=td.get('lift_capacity_t'),
            vfd=vfd,explosion_proof=expl,operation_status=oper,
            acce_item_symbol=sym,acce_item_type=sub,
            action="NEW",is_valid=len(w)==0,warnings=w)
        records.append(rec)
        if w: warns.append({'row':ri,'tag':tag,'reason':'; '.join(w)})
    return records,warns

# ── Формат 2: Регламент ───────────────────────────────────────────
def _detect_section(text,cur):
    if re.search(r'5\.1.{0,10}Характеристика.{0,30}ректификацион',text,re.I): return "5.1"
    if re.search(r'5\.2.{0,10}Характеристика.{0,30}колонного',text,re.I): return "5.2"
    if re.search(r'5\.3.{0,10}Характеристика.{0,30}теплообменного',text,re.I): return "5.3"
    return cur

def _row51(row,ri,fp):
    if len(row)<22: return None
    tag=_c(row[2])
    if not tag or not re.search(r'[А-ЯЁA-Z]-\d{2,3}',tag): return None
    name=_c(row[3]); qty=_parse_qty(_c(row[4]))
    pressure, temp = _assign_pt(_c(row[16]),_c(row[17]),_c(row[18]),_c(row[19]))
    tech=_c(row[20]); mat=_c(row[21])
    td=_parse_tech_reg(tech)
    return ACCERecord(
        source_file=fp,source_sheet="5.1",source_row=ri,
        raw_tag=tag,user_tag=_make_tag(tag),description_ru=name,
        quantity=qty,material=mat,tech_raw=tech,
        pressure=pressure,pressure_unit="МПа",
        design_temperature=temp,
        diameter_m=td.get('diameter_m'),length_m=td.get('length_m'),
        acce_item_symbol="TW",acce_item_type=_column_subtype(tech),
        action="NEW",is_valid=True)

def _row52(row,ri,fp):
    if len(row)<17: return None
    tag=_c(row[2])
    if not tag or not re.search(r'[А-ЯЁA-Z]-\d{2,3}',tag): return None
    name=_c(row[3]); qty=_parse_qty(_c(row[4]))
    tech=_c(row[14]); mat=_c(row[16])
    td=_parse_tech_reg(tech)
    return ACCERecord(
        source_file=fp,source_sheet="5.2",source_row=ri,
        raw_tag=tag,user_tag=_make_tag(tag),description_ru=name,
        quantity=qty,material=mat,tech_raw=tech,
        pressure=_parse_pv_pair(_c(row[11])),pressure_unit="МПа",
        design_temperature=_parse_pv_pair(_c(row[8])),
        diameter_m=td.get('diameter_m'),length_m=td.get('length_m'),
        acce_item_symbol="TW",acce_item_type=_column_subtype(tech),
        action="NEW",is_valid=True)

def _row53(row,ri,fp):
    if len(row)<16: return None
    tag=_c(row[2])
    if not tag or not re.search(r'[А-ЯЁ]-\d{3}',tag): return None
    name=_c(row[3]); qty=_parse_qty(_c(row[4]))
    tech=_c(row[15]); mat=_c(row[16]) if len(row)>16 else ""
    td=_parse_tech_reg(tech)
    bm=re.search(r'[А-ЯЁ]-\d{3}',tag); btag=bm.group() if bm else tag
    w=[]
    if not td.get('heat_transfer_area_m2'): w.append("Нет площади теплообмена в тех. описании")
    return ACCERecord(
        source_file=fp,source_sheet="5.3",source_row=ri,
        raw_tag=btag,user_tag=_make_tag(btag),description_ru=name,
        quantity=qty,material=mat,tech_raw=tech,
        pressure=td.get('pressure'),pressure_unit="МПа",
        design_temperature=td.get('design_temperature'),
        heat_transfer_area_m2=td.get('heat_transfer_area_m2'),
        heat_duty_gcalh=td.get('heat_duty_gcalh'),
        acce_item_symbol="HE",acce_item_type=_he_subtype(tech),
        action="NEW",is_valid=len(w)==0,warnings=w)

# ── Главная функция ───────────────────────────────────────────────
def parse_pdf(filepath:str)->tuple:
    """
    Универсальный PDF-парсер.
    Автоматически определяет формат:
      - Equipment List (Пример_1): 1 страница A3, 23 колонки
      - Технологический регламент (Пример_2): разделы 5.1/5.2/5.3

    Возвращает (list[ACCERecord], list[dict]).
    """
    records,warns=[],[]
    with pdfplumber.open(filepath) as pdf:
        text0=pdf.pages[0].extract_text() or ""
        tbs0=pdf.pages[0].extract_tables()
        ncols=len(tbs0[0][0]) if tbs0 else 0
        is_el=(len(pdf.pages)<=2 and ncols>=20 and
               not re.search(r'5\.[123]\s+Характеристика',text0,re.I))
        if is_el:
            for page in pdf.pages:
                r,w=_parse_el_page(page,filepath)
                records.extend(r); warns.extend(w)
        else:
            sec=None; seen=set()
            for pi,page in enumerate(pdf.pages):
                text=page.extract_text() or ""
                tbs=page.extract_tables()
                if not tbs: continue
                sec=_detect_section(text,sec)
                fn={'5.1':_row51,'5.2':_row52,'5.3':_row53}.get(sec)
                if not fn: continue
                for ri,row in enumerate(tbs[0]):
                    rec=fn(row,ri,filepath)
                    if rec is None: continue
                    key=f"{sec}:{rec.raw_tag}"
                    if key in seen: continue
                    seen.add(key); records.append(rec)
                    if rec.warnings:
                        warns.append({'page':pi+1,'section':sec,
                                      'tag':rec.raw_tag,'reason':'; '.join(rec.warnings)})
    return records,warns