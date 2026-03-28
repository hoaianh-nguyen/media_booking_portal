#!/usr/bin/env python3
"""
Run this once to extract Excel data into JSON files.
    python3 prepare_data.py
"""

import pandas as pd
import json, re, os

EXCEL_FILE = "In-app Automatic Flow.xlsx"
if not os.path.exists(EXCEL_FILE):
    EXCEL_FILE = "In-app_Automatic_Flow.xlsx"

# ── helpers ───────────────────────────────────────────────────────────────
def clean(v):
    return '' if (v is None or (isinstance(v, float) and str(v) == 'nan')) else str(v).strip()

PROG_MAP = {
    '[PLF] BAU':'PLF_BAU','[PLF]BAU':'PLF_BAU','[PLF CP]_BAU':'PLF_BAU',
    '[PLF CP]':'PLF_CP','[PLF] WC':'PLF_WC','[PLF] CP':'PLF_CP',
    '[PLF] Mass Ads':'PLF_MASSADS','[Mass Ads]':'PLF_MASSADS',
    '[PLF] Mart':'PLF_MART','[Mart]':'PLF_MART',
    '[PLF] New User':'PLF_NEWUSER','[SJBP]':'SJBP','[JBP]':'JBP',
    '[SuperVIP]':'SUPERVIP','[VIP]':'VIP',
    '[Win 1]':'WIN1','[Win 2]':'WIN2','[Win 3]':'WIN3','[Win 4]':'WIN4',
    '[Alacarte]':'ALACARTE','[Alacarte CB]':'ALACARECB','[PNS]':'PNS',
}

def prog(s):
    for k, v in PROG_MAP.items():
        if k in str(s): return v
    m = re.search(r'\[([^\]]+)\]', str(s))
    return m.group(1).upper().replace(' ','_') if m else 'OTHER'

def cities(text):
    out = [c for c in ['HCM','HN','DN','HP','CT','OTH']
           if re.search(r'\b'+c+r'\b', str(text))]
    return out or ['HCM']

def parse_date(tl, month):
    yr_map = {'Jan':'2026','Feb':'2026','Mar':'2026','Apr':'2026',
              'May':'2026','Jun':'2026','Jul':'2026','Aug':'2026',
              'Sep':'2026','Oct':'2026','Nov':'2026','Dec':'2025'}
    yr = yr_map.get(str(month).strip().capitalize()[:3], '2026')
    m = re.search(r'(\d{1,2})/(\d{1,2})\s*[-\u2013]\s*(\d{1,2})/(\d{1,2})', str(tl))
    if m: return f'{yr}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}', \
                 f'{yr}-{m.group(4).zfill(2)}-{m.group(3).zfill(2)}'
    m2 = re.search(r'(\d{1,2})/(\d{1,2})', str(tl))
    if m2: d = f'{yr}-{m2.group(2).zfill(2)}-{m2.group(1).zfill(2)}'; return d, d
    return '', ''

def banner(a, b=''):
    t = (str(a)+' '+str(b)).lower()
    for kw, bn in [('htc','HTC'),('scroll','HTC'),('hometop','Hometop'),
                   ('static','Static'),('bottom','Bottom'),('flash','Flash'),
                   ('float','Floating'),('pop','Popup'),
                   ('ongoing order','OngoingOrder'),('ongoing','Ongoing'),
                   ('search','Search'),('collection list','CollList'),
                   ('flash sale','Flash'),('recommendation','Rec'),
                   ('collection','Collection')]:
        if kw in t: return bn
    # MKT time-range slots (e.g. "0-5H", "5-10H") → Search Placeholder
    import re as _re
    if _re.search(r'^\d', t.strip()):
        return 'Search'
    return 'Hometop'



def normalise_timeslot(text, source=''):
    """Map any timeslot text to a standard 5-bucket time label."""
    t = str(text or '').lower().strip()
    
    # Vietnamese keywords (MKT and BD descriptions)
    if any(x in t for x in ['sáng','sang','bữa sáng','ăn sáng','bua sang']):
        return '0-10h'
    if any(x in t for x in ['trưa','trua','bữa trưa','ăn trưa','bua trua']):
        return '10-13h'
    if any(x in t for x in ['xế','bữa xế','ăn xế','bua xe']):
        return '13-16h'
    if any(x in t for x in ['tối','toi','bữa tối','ăn tối','buổi tối','bua toi']):
        return '16-21h'
    if any(x in t for x in ['đêm','dem','ăn đêm','khuya','ban đêm']):
        return '21-24h'
    
    # Extract hour range from patterns: (16H-19H), (14h-17h), 0-8h, 8-16h, 16-24h
    m = re.search(r'(\d{1,2})h?\s*[-–]\s*(\d{1,2})h', t)
    if m:
        start_h = int(m.group(1))
        # Map start hour to bucket
        if 0 <= start_h < 10:   return '0-10h'
        if 10 <= start_h < 13:  return '10-13h'
        if 13 <= start_h < 16:  return '13-16h'
        if 16 <= start_h < 21:  return '16-21h'
        if start_h >= 21:        return '21-24h'
    
    # HTC slot keys: 0-8h, 8-16h, 16-24h
    if '0-8h' in t or '0h-8h' in t:     return '0-10h'
    if '8-16h' in t or '8h-16h' in t:    return '10-13h'  # approximate
    if '16-24h' in t or '16h-24h' in t:  return '16-21h'
    
    return ''

def mkt_timeslot(text):
    """Detect time slot from MKT booking name."""
    return normalise_timeslot(text, 'MKT')

def guess_month(s):
    for mo in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']:
        if mo.lower() in str(s).lower(): return mo
    return 'Jan'

# ── parse visibility cell into booking array ──────────────────────────────
def parse_booking(raw_val):
    """Convert a visibility cell value into a sorted list of booking dicts."""
    bookings = [b.strip() for b in raw_val.split('\n') if b.strip()]
    result = []
    for bk in bookings:
        tm = re.match(r'^([0-9]{1,2}h[0-9h +&\-]{0,30}h)[: ]*', bk, re.I)
        raw_time = tm.group(1).strip() if tm else ''
        rest = bk[tm.end():].strip() if tm else bk
        brand = re.sub(r'\[[^\]]+\]_*', '', rest).replace('_', ' ').strip()
        l2 = get_l2(brand + ' ' + rest)
        times = raw_time.split('+') if raw_time else ['']
        for t in times:
            result.append({'t': t.strip(), 'p': rest, 'b': brand, 'l': l2})
    def _hr(t):
        m = re.match(r'^(\d+)', t or '')
        return int(m.group(1)) if m else 99
    result.sort(key=lambda x: _hr(x.get('t','')))
    return result

# ── extract bookings ──────────────────────────────────────────────────────
def extract_bookings(xl):
    rows = []
    uid = 1

    # BD Programs — parse media columns (HCM=col17, HN=col18, DN=col19) per line
    df = pd.read_excel(xl, sheet_name='BD Programs', header=None)
    CITY_COLS = [(16,'HCM'),(17,'HN'),(18,'DN')]  # 0-indexed
    for i in range(4, len(df)):
        r = df.iloc[i]
        prog_raw = clean(r[3]); brand = clean(r[4])
        if not prog_raw and not brand: continue
        if prog_raw in ('Programs','Chon data validation'): continue
        month = clean(r[0]) or 'Jan'
        status = (clean(r[20]) if len(r) > 20 else '')[:20]
        pic = clean(r[2])[:20]
        p = prog(prog_raw)
        tl_global = clean(r[8])
        start_g, end_g = parse_date(tl_global, month)
        
        # Try to extract per-banner bookings from media columns
        added = False
        for col_idx, city_name in CITY_COLS:
            if col_idx >= len(r): continue
            cell = clean(r[col_idx])
            if not cell or cell in ('Media Slot', 'nan'): continue
            # Each line in the cell describes a different banner slot
            for line in cell.split('\n'):
                line = line.strip()
                if not line or len(line) < 4: continue
                bn = banner(line)
                start_l, end_l = parse_date(line, month)
                rows.append({
                    'id': uid, 'source': 'BD', 'month': month,
                    'prog': p, 'prog_raw': prog_raw[:40],
                    'brand': brand[:80], 'scheme': clean(r[5])[:60],
                    'start': start_l or start_g, 'end': end_l or end_g,
                    'cities': [city_name], 'city': city_name,
                    'status': status, 'pic': pic,
                    'banner': bn, 'timeslot': line[:60], 'cat': '',
                })
                uid += 1
                added = True
        
        # Fallback: if no media columns had data, create one booking from timeline
        if not added:
            c = cities(clean(r[7]) + ' ' + tl_global)
            rows.append({
                'id': uid, 'source': 'BD', 'month': month,
                'prog': p, 'prog_raw': prog_raw[:40],
                'brand': brand[:80], 'scheme': clean(r[5])[:60],
                'start': start_g, 'end': end_g,
                'cities': c, 'city': c[0] if c else 'HCM',
                'status': status, 'pic': pic,
                'banner': banner(prog_raw, clean(r[6])), 'timeslot': 'all', 'cat': '',
            })
            uid += 1

    # MKT|Campaign
    df = pd.read_excel(xl, sheet_name='MKT|Campaign', header=None)
    for i in range(1, len(df)):
        r = df.iloc[i]
        deal = clean(r[2])
        if len(deal) < 5: continue
        if '[PLF' not in deal and '[Mart]' not in deal: continue
        end_raw = r[5]; end = ''
        if hasattr(end_raw, 'strftime'): end = end_raw.strftime('%Y-%m-%d')
        elif isinstance(end_raw, str):
            m = re.search(r'(\d{4}-\d{2}-\d{2})', end_raw)
            if m: end = m.group(1)
        slot = clean(r[1])
        mkt_ts = normalise_timeslot(deal, 'MKT')
        rows.append({
            'id': uid, 'source': 'MKT', 'month': guess_month(deal),
            'prog': prog(deal), 'prog_raw': '',
            'brand': deal[:80], 'scheme': clean(r[3])[:60],
            'start': '2026-01-01', 'end': end or '2026-12-31',
            'cities': ['HCM','HN','DN'], 'city': 'HCM',
            'status': clean(r[0])[:20], 'pic': '',
            'banner': banner(slot), 'timeslot': mkt_ts, 'cat': '',
        })
        uid += 1

    return [b for b in rows if b['brand'] and b['brand'] != 'nan']

# ── extract visibility ────────────────────────────────────────────────────
def extract_visibility(xl):
    CITY_MAP = {'HCM':'HCM','Ho Chi Minh':'HCM','Ha Noi':'HN',
                'HN':'HN','Da Nang':'DN','DN':'DN','OTH':'OTH'}
    # Complete SLOT_MAP — col B label → slot key
    # Floating & Popup are handled by section context (col A), not here
    SLOT_MAP = [
        # Search Placeholder (col A: "Search Place Holder")
        ('01:00 - 05:00','Search_01-05h'), ('05:00 - 10:00','Search_05-10h'),
        ('10:00 - 12:00','Search_10-12h'), ('12:00 - 14:00','Search_12-14h'),
        ('14:00 - 17:00','Search_14-17h'), ('17:00 - 19:00','Search_17-19h'),
        ('19:00 - 22:00','Search_19-22h'), ('22:00 - 01:00','Search_22-01h'),
        # Home Circle / HTC (col A: "Home Circle")
        ('1st Scroll 0-8h','HTC_1st_0-8h'), ('1st Scroll 8-16h','HTC_1st_8-16h'),
        ('1st Scroll 16-24h','HTC_1st_16-24h'),
        ('2nd Scroll 0-8h','HTC_2nd_0-8h'), ('2nd Scroll 8-16h','HTC_2nd_8-16h'),
        ('2nd Scroll 16-24h','HTC_2nd_16-24h'),
        ('2st Scroll 0-8h','HTC_2nd_0-8h'), ('2st Scroll 8-16h','HTC_2nd_8-16h'),  # typo variant
        ('2st Scroll 16-24h','HTC_2nd_16-24h'),
        # Recommendation
        ('Recommendation 1','Rec_1'), ('Recommendation 2','Rec_2'),
        ('Recommendation 3','Rec_3'), ('Recommendation 4','Rec_4'),
        ('Recommendation 5','Rec_5'),
        # Hometop Banner
        ('Hometop Banner 0','Hometop_0'),
        ('Hometop Banner 1','Hometop_1'), ('Hometop Banner 2','Hometop_2'),
        ('Hometop Banner 3','Hometop_3'), ('Hometop Banner 4','Hometop_4'),
        ('Hometop Banner 5','Hometop_5'), ('Hometop Banner 6','Hometop_6'),
        ('Hometop Banner 7','Hometop_7'), ('Hometop Banner 8','Hometop_8'),
        ('Hometop Banner 9','Hometop_9'), ('Hometop Banner 10','Hometop_10'),
        ('Hometop Banner 11','Hometop_11'), ('Hometop Banner 12','Hometop_12'),
        ('Hometop Banner 13','Hometop_13'), ('Hometop Banner 14','Hometop_14'),
        # Static Banner
        ('Static Banner 1','Static_1'), ('Static Banner 2','Static_2'),
        ('Static Banner 3','Static_3'), ('Static Banner 4','Static_4'),
        ('Static Banner 5','Static_5'), ('Static Banner 6','Static_6'),
        # Bottom Banner
        ('Bottom Banner 1','Bottom_1'), ('Bottom Banner 2','Bottom_2'),
        ('Bottom Banner 3','Bottom_3'), ('Bottom Banner 4','Bottom_4'),
        ('Bottom Banner 5','Bottom_5'), ('Bottom Banner 6','Bottom_6'),
        # Collection List Banner (different from Collection)
        ('Collection List Banner 1','CollList_1'), ('Collection List Banner 2','CollList_2'),
        ('Collection List Banner 3','CollList_3'),
        # Search Banner
        ('Search Banner 1','Search_Banner'), ('Search Banner 2','Search_Banner2'),
        ('Search Banner','Search_Banner'),
        # Flash Sale
        ('Flash Sale 1','Flash_1'), ('Flash Sale 2','Flash_2'),
        ('Flash Sale 3','Flash_3'), ('Flash Sale 4','Flash_4'),
        # On-Going / Ongoing
        ('On-Going','Ongoing'), ('OnGoing','Ongoing'), ('Ongoing Order','OngoingOrder'),
        # Collection (rows 99-114)
        ('Collection 1','Coll_1'), ('Collection 2','Coll_2'), ('Collection 3','Coll_3'),
        ('Collection 4','Coll_4'), ('Collection 5','Coll_5'), ('Collection 6','Coll_6'),
        ('Collection 7','Coll_7'), ('Collection 8','Coll_8'), ('Collection 9','Coll_9'),
        ('Collection 10','Coll_10'), ('Collection 11','Coll_11'), ('Collection 12','Coll_12'),
        ('Collection 13','Coll_13'), ('Collection 14','Coll_14'), ('Collection 15','Coll_15'),
        ('Collection 16','Coll_16'),
    ]

    def parse_sheet(df):
        # Build date→city→col index (row 2 = dates, row 4 = cities)
        date_city_col = {}
        for col in range(3, df.shape[1]):
            cell = df.iloc[1, col]
            d = None
            if hasattr(cell, 'strftime'): d = cell.strftime('%Y-%m-%d')
            elif isinstance(cell, str):
                mx = re.search(r'(\d{4}-\d{2}-\d{2})', cell)
                if mx: d = mx.group(1)
            if not d or d < '2026-01': continue
            city = CITY_MAP.get(str(df.iloc[3, col]).strip())
            if city: date_city_col[(d, city)] = col

        # Build slot→row index with section-aware parsing
        slot_rows = {}
        current_section = ''
        for i in range(4, min(135, len(df))):
            col_a = str(df.iloc[i, 0] or '').strip()
            col_b = str(df.iloc[i, 1] or '').strip()
            # Update section when col A has a value
            if col_a and col_a.lower() not in ('none', ''):
                current_section = col_a.split('\n')[0].strip().lower()
            lb = (col_b.split('\n')[0].strip()) or (col_a.split('\n')[0].strip())
            if not lb or lb.lower() in ('none', ''): continue
            lb_lower = lb.lower()

            # Floating Banner — section-aware (time slots shared with Popup)
            if 'floating' in current_section:
                for t, k in [('00:00 - 06:00','Float_00-06'),('06:00 - 10:00','Float_06-10'),
                             ('10:00 - 13:00','Float_10-13'),('13:00 - 16:00','Float_13-16'),
                             ('16:00 - 19:00','Float_16-19'),('19:00 - 22:00','Float_19-22'),
                             ('22:00 - 24:00','Float_22-24')]:
                    if t in lb:
                        layer = 'L2' if 'layer 2' in lb_lower else 'L1'
                        slot_rows[i] = f'{k}_{layer}'
                        break
                continue

            # Popup Banner — section-aware
            if 'popup' in current_section:
                for t, k in [('00:00 - 06:00','Popup_00-06'),('06:00 - 10:00','Popup_06-10'),
                             ('10:00 - 13:00','Popup_10-13'),('13:00 - 16:00','Popup_13-16'),
                             ('16:00 - 19:00','Popup_16-19'),('19:00 - 22:00','Popup_19-22'),
                             ('22:00 - 24:00','Popup_22-24')]:
                    if t in lb:
                        layer = 'L2' if 'layer 2' in lb_lower else 'L1'
                        slot_rows[i] = f'{k}_{layer}'
                        break
                continue

            # Foody App / History — skip
            if 'foody' in current_section or 'history' in current_section:
                continue

            # Standard SLOT_MAP matching
            for pattern, key in SLOT_MAP:
                if pattern in lb:
                    if key in ('Hometop_1','Hometop_2') and 'Mart' in lb: continue
                    slot_rows[i] = key
                    break

        vis = {}
        for (date, city), col in date_city_col.items():
            if col >= df.shape[1]: continue
            for row, slot in slot_rows.items():
                val = df.iloc[row, col]
                if val and str(val) != 'nan' and str(val).strip():
                    # Keep full value with \n as multi-booking separator
                    raw_val = str(val).strip().replace('\r', '')
                    vis.setdefault(date, {}).setdefault(city, {})[slot] = parse_booking(raw_val)
        return vis

    try:
        df3b  = pd.read_excel(xl, sheet_name='Visibility 3B',  header=None)
        dfoth = pd.read_excel(xl, sheet_name='Visibility OTH', header=None)
        merged = parse_sheet(df3b)
        for d, cities in parse_sheet(dfoth).items():
            merged.setdefault(d, {}).update(cities)
        return merged
    except Exception as e:
        print(f"  Warning: {e}")
        return {}

# ── L2 category mapping ──────────────────────────────────────────────────
L2_BRANDS = {
    'highlands coffee':'Cafe','starbucks':'Cafe','the coffee house':'Cafe',
    'phuc long':'Cafe','phúc long':'Cafe','phe la':'Cafe','phê la':'Cafe',
    'cheese coffee':'Cafe','katinat':'Cafe','cà phê muối':'Cafe',
    'cà phê':'Cafe','ca phe':'Cafe','cafe':'Cafe',
    'tra sua':'Trà Sữa','trà sữa':'Trà Sữa','tocotoco':'Trà Sữa',
    'sunday basic':'Trà Sữa','sekai':'Trà Sữa','mixue':'Trà Sữa',
    'rau má mix':'Trà Sữa','may cha':'Trà Sữa','gong cha':'Trà Sữa',
    'kfc':'Gà Rán','popeyes':'Gà Rán','gà rán':'Gà Rán','ga ran':'Gà Rán',
    'texas chicken':'Gà Rán','jollibee':'Gà Rán',
    "domino's":'Pizza','dominos':'Pizza','pizza hut':'Pizza','pizza':'Pizza',
    "mcdonald's":'Burger','mcdonalds':'Burger','lotteria':'Burger','burger':'Burger',
    'cơm tấm':'Cơm Tấm','com tam':'Cơm Tấm',
    'bún đậu':'Bún Đậu','bun dau':'Bún Đậu',
    'phở':'Phở','pho':'Phở','hủ tiếu':'Hủ Tiếu',
    'hanuri':'Món Hàn','korean':'Món Hàn','hàn quốc':'Món Hàn',
    'bánh mì':'Bánh Mì','banh mi':'Bánh Mì',
    'chè':'Chè','che ':'Chè',
    'cơm':'Cơm',
}

def get_l2(text):
    t = text.lower()
    for kw, cat in L2_BRANDS.items():
        if kw in t: return cat
    return ''

# ── caps data ─────────────────────────────────────────────────────────────
CAPS = {
    "daily": {
        "Hometop": {"PLF":3,"SJBP":2,"JBP":2,"VIP":1,"IND":6,"PNS":1,"TOTAL":14},
        "HTC":     {"PLF":3,"SJBP":2,"JBP":2,"VIP":1,"IND":3,"TOTAL":3},
        "Static":  {"VIP":2,"IND":4,"TOTAL":6},
        "Bottom":  {"VIP":2,"IND":4,"TOTAL":6},
        "Flash":   {"SJBP":4,"IND":4,"TOTAL":4},
        "Floating":{"PLF":2,"SJBP":1,"JBP":1,"VIP":1,"IND":4,"TOTAL":7},
        "Popup":   {"PLF":2,"SJBP":2,"JBP":2,"VIP":1,"IND":3,"TOTAL":7},
        "Ongoing": {"PNS":1,"TOTAL":1},
        "Search":  {"PNS":1,"TOTAL":1},
    },
    "monthly": {
        "HTC":     {"HCM":90,"HN":90,"DN":90,"OTH":90},
        "Hometop": {"HCM":60,"HN":60,"DN":60,"OTH":60},
        "Static":  {"HCM":90,"HN":90,"DN":90,"OTH":90},
        "Bottom":  {"HCM":90,"HN":90,"DN":90,"OTH":90},
        "Flash":   {"HCM":120,"HN":120,"DN":120,"OTH":120},
        "Floating":{"HCM":210,"HN":210,"DN":210,"OTH":210},
        "Popup":   {"HCM":180,"HN":180,"DN":180,"OTH":180},
        "Ongoing": {"HCM":30,"HN":30,"DN":30,"OTH":30},
        "Search":  {"HCM":30,"HN":30,"DN":30,"OTH":30},
    }
}

# ── main ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: Cannot find '{EXCEL_FILE}'")
        print("Place your Excel file in the same folder as this script.")
        exit(1)

    print(f"Reading {EXCEL_FILE}...")
    xl = pd.ExcelFile(EXCEL_FILE)

    print("  Extracting bookings...")
    bookings = extract_bookings(xl)
    print(f"  → {len(bookings)} bookings")

    print("  Extracting visibility...")
    vis = extract_visibility(xl)
    print(f"  → {len(vis)} dates")

    os.makedirs('data', exist_ok=True)

    # Slim bookings to reduce wire size
    def slim_booking(b):
        return {
            'id':  b['id'],   'src': b['source'],   'mo':  b['month'],
            'p':   b['prog'], 'pr':  b['prog_raw'],  'b':   b['brand'],
            'st':  b.get('start',''), 'en': b.get('end',''),
            'cs':  b.get('cities',[]), 'city': b.get('city',''),
            'ss':  b.get('status',''), 'pic': b.get('pic',''),
            'bn':  b['banner'],  'ts': normalise_timeslot(b.get('timeslot',''), b.get('source','')),
        }
    slimmed = [slim_booking(b) for b in bookings]
    with open('data/bookings.json',   'w', encoding='utf-8') as f:
        json.dump(slimmed, f, ensure_ascii=False, separators=(',',':'))
    with open('data/visibility.json', 'w', encoding='utf-8') as f:
        json.dump(vis, f, ensure_ascii=False, separators=(',',':'))
    with open('data/caps.json',       'w', encoding='utf-8') as f:
        json.dump(CAPS, f, ensure_ascii=False, separators=(',',':'))

    print(f"\nDone! Files saved to data/")
    bk_kb  = os.path.getsize('data/bookings.json')   // 1024
    vis_kb = os.path.getsize('data/visibility.json') // 1024
    print(f"  bookings.json   : {bk_kb} KB")
    print(f"  visibility.json : {vis_kb} KB")
    print(f"  caps.json       : {os.path.getsize('data/caps.json')} bytes")
    print(f"\nNow run: python3 app.py")
