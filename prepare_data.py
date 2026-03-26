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
                   ('search','Search')]:
        if kw in t: return bn
    return 'Hometop'

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

    # BD Programs
    df = pd.read_excel(xl, sheet_name='BD Programs', header=None)
    for i in range(4, len(df)):
        r = df.iloc[i]
        prog_raw = clean(r[3]); brand = clean(r[4])
        if not prog_raw and not brand: continue
        if prog_raw in ('Programs','Chon data validation'): continue
        month = clean(r[0]) or 'Jan'
        detail = clean(r[7]); tl = clean(r[8])
        start, end = parse_date(tl, month)
        c = cities(detail + ' ' + tl)
        rows.append({
            'id': uid, 'source': 'BD', 'month': month,
            'prog': prog(prog_raw), 'prog_raw': prog_raw[:40],
            'brand': brand[:80], 'scheme': clean(r[5])[:60],
            'start': start, 'end': end,
            'cities': c, 'city': c[0] if c else 'HCM',
            'status': (clean(r[20]) if len(r) > 20 else '')[:20],
            'pic': clean(r[2])[:20],
            'banner': banner(prog_raw, clean(r[6])),
            'timeslot': 'all', 'cat': '',
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
        rows.append({
            'id': uid, 'source': 'MKT', 'month': guess_month(deal),
            'prog': prog(deal), 'prog_raw': '',
            'brand': deal[:80], 'scheme': clean(r[3])[:60],
            'start': '2026-01-01', 'end': end or '2026-12-31',
            'cities': ['HCM','HN','DN'], 'city': 'HCM',
            'status': clean(r[0])[:20], 'pic': '',
            'banner': banner(slot), 'timeslot': slot[:40], 'cat': '',
        })
        uid += 1

    return [b for b in rows if b['brand'] and b['brand'] != 'nan']

# ── extract visibility ────────────────────────────────────────────────────
def extract_visibility(xl):
    CITY_MAP = {'HCM':'HCM','Ho Chi Minh':'HCM','Ha Noi':'HN',
                'HN':'HN','Da Nang':'DN','DN':'DN','OTH':'OTH'}
    SLOT_MAP = [
        ('01:00 - 05:00','Search_01-05h'), ('05:00 - 10:00','Search_05-10h'),
        ('10:00 - 12:00','Search_10-12h'), ('12:00 - 14:00','Search_12-14h'),
        ('14:00 - 17:00','Search_14-17h'), ('17:00 - 19:00','Search_17-19h'),
        ('19:00 - 22:00','Search_19-22h'),
        ('22:00 - 01:00','Search_22-01h'),
        ('1st Scroll 0-8h','HTC_1st_0-8h'),('1st Scroll 8-16h','HTC_1st_8-16h'),
        ('1st Scroll 16-24h','HTC_1st_16-24h'),('2nd Scroll 0-8h','HTC_2nd_0-8h'),
        ('2nd Scroll 8-16h','HTC_2nd_8-16h'),('2nd Scroll 16-24h','HTC_2nd_16-24h'),
        ('Hometop Banner 0','Hometop_0'),
        ('Hometop Banner 1','Hometop_1'),('Hometop Banner 2','Hometop_2'),
        ('Hometop Banner 3','Hometop_3'),('Hometop Banner 4','Hometop_4'),
        ('Hometop Banner 5','Hometop_5'),('Hometop Banner 6','Hometop_6'),
        ('Hometop Banner 7','Hometop_7'),('Hometop Banner 8','Hometop_8'),
        ('Hometop Banner 9','Hometop_9'),('Hometop Banner 10','Hometop_10'),
        ('Hometop Banner 11','Hometop_11'),('Hometop Banner 12','Hometop_12'),
        ('Hometop Banner 13','Hometop_13'),
        ('Static Banner 1','Static_1'),('Static Banner 2','Static_2'),
        ('Static Banner 3','Static_3'),('Static Banner 4','Static_4'),
        ('Static Banner 5','Static_5'),('Static Banner 6','Static_6'),
        ('Bottom Banner 1','Bottom_1'),('Bottom Banner 2','Bottom_2'),
        ('Bottom Banner 3','Bottom_3'),('Bottom Banner 4','Bottom_4'),
        ('Bottom Banner 5','Bottom_5'),('Bottom Banner 6','Bottom_6'),
        ('Flash Sale 1','Flash_1'),('Flash Sale 2','Flash_2'),
        ('Flash Sale 3','Flash_3'),('Flash Sale 4','Flash_4'),
        ('On-Going','Ongoing'),('Search Banner','Search_Banner'),
    ]

    def parse_sheet(df):
        # Build date→city→col index
        date_city_col = {}
        for col in range(3, df.shape[1]):
            cell = df.iloc[1, col]
            d = None
            if hasattr(cell, 'strftime'): d = cell.strftime('%Y-%m-%d')
            elif isinstance(cell, str):
                mx = re.search(r'(\d{4}-\d{2}-\d{2})', cell)
                if mx: d = mx.group(1)
            if not d or d < '2026-01': continue   # only 2026+
            city = CITY_MAP.get(str(df.iloc[3, col]).strip())
            if city: date_city_col[(d, city)] = col

        # Build slot→row index
        slot_rows = {}
        for i in range(4, min(125, len(df))):
            lb = (clean(df.iloc[i,1]) or clean(df.iloc[i,0])).strip()
            for pattern, key in SLOT_MAP:
                if pattern in lb:
                    # Hometop_1 must not match "Hometop Banner 13 Mart"
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

    with open('data/bookings.json',   'w', encoding='utf-8') as f:
        json.dump(bookings, f, ensure_ascii=False, separators=(',',':'))

    with open('data/visibility.json', 'w', encoding='utf-8') as f:
        json.dump(vis,      f, ensure_ascii=False, separators=(',',':'))

    with open('data/caps.json',       'w', encoding='utf-8') as f:
        json.dump(CAPS,     f, ensure_ascii=False, separators=(',',':'))

    print(f"\nDone! Files saved to data/")
    bk_kb  = os.path.getsize('data/bookings.json')   // 1024
    vis_kb = os.path.getsize('data/visibility.json') // 1024
    print(f"  bookings.json   : {bk_kb} KB")
    print(f"  visibility.json : {vis_kb} KB")
    print(f"  caps.json       : {os.path.getsize('data/caps.json')} bytes")
    print(f"\nNow run: python3 app.py")
