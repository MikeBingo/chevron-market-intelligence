"""
generate_dashboard_data.py
──────────────────────────────────────────────────────────────────────────────
Reads the latest TIMB FCV daily summary file and injects live regional and
overview data into the Market Intelligence Dashboard 2026.html.

Updates two data blocks in the HTML:
  • /* RA_DATA_START */ … /* RA_DATA_END */   — Regional Competitiveness (P1)
  • /* OV_DATA_START */ … /* OV_DATA_END */   — Overview chart data (P0)

Run this script after each new daily FCV file arrives (or via run_2026_market_data.py).
"""

import os, sys, re, json, glob
from datetime import datetime, date, timedelta
from pathlib import Path

_WIN_ROOT = r"C:\Users\MikeBingo\Chevron Leaf Tobacco\Power BI Uploads - 2026 TIMB FCV Daily Files"
_LNX_ROOT = "/sessions/vibrant-tender-cray/mnt/Power BI Uploads - 2026 TIMB FCV Daily Files"
ROOT      = _WIN_ROOT if os.path.exists(_WIN_ROOT) else _LNX_ROOT
DAILY_DIR = ROOT   # daily FCV files live directly in root folder
DASHBOARD = os.path.join(ROOT, "Market Intelligence Dashboard 2026.html")

# ── Region configuration (static per season) ──────────────────────────────────
REGIONS = {
    'KAROI':   {'col': '#2D6A4F', 'crop25': 54001067, 'target': 1800000, 'sheet': 'KAROI'},
    'RUSAPE':  {'col': '#E67E22', 'crop25': 16091025, 'target': 800000,  'sheet': 'RUSAPE '},
    'BINDURA': {'col': '#3498DB', 'crop25': 7020698,  'target': 900000,  'sheet': 'BINDURA '},
    'HARARE':  {'col': '#9B59B6', 'crop25': 83000000, 'target': 300000,  'sheet': 'HARARE'},
}

STRIDE = 10   # columns per day block in regional sheets
CHEVRON_FRAG = 'Chevron Tobacco'


# ── Helpers ────────────────────────────────────────────────────────────────────
def day_num_from_path(p):
    m = re.search(r'day (\d+)', p, re.IGNORECASE)
    return int(m.group(1)) if m else 0

def find_latest_file():
    files = glob.glob(os.path.join(DAILY_DIR, "2026 daily fcv summary day *.xlsx"))
    if not files:
        sys.exit("No daily FCV files found in root folder.")
    # Sort descending by day number; skip files that can't be opened (e.g. OneDrive cloud-only stubs)
    for f in sorted(files, key=day_num_from_path, reverse=True):
        try:
            import zipfile
            with open(f, 'rb') as fh:
                fh.read(4)   # just check we can read bytes
            return f, day_num_from_path(f)
        except OSError:
            continue
    sys.exit("No readable daily FCV files found.")

def load_wb(path):
    import openpyxl
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        sys.exit(f"Cannot read {path}: {e}")

def n(v, default=0.0):
    """Safe numeric conversion."""
    try:
        return float(v) if v is not None else default
    except (TypeError, ValueError):
        return default

def wgt_price(mass, value):
    return value / mass if mass > 0 else 0.0


# ── Step 1: Build day → date mapping (scan all daily files) ───────────────────
def build_day_dates():
    files = glob.glob(os.path.join(DAILY_DIR, "2026 daily fcv summary day *.xlsx"))
    day_dates = {}
    for f in sorted(files, key=day_num_from_path):
        d = day_num_from_path(f)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            ws = wb['Auction and Contract']
            for row in ws.iter_rows(values_only=True, min_row=1, max_row=6):
                for v in row:
                    if v and 'Date:' in str(v):
                        dt_str = str(v).replace('Date:', '').strip()
                        try:
                            day_dates[d] = datetime.strptime(dt_str, '%d/%m/%Y').date()
                        except ValueError:
                            pass
            wb.close()
        except Exception:
            pass
    return day_dates


# ── Step 2: Read one regional sheet ───────────────────────────────────────────
def read_regional_sheet(ws, max_day):
    """
    Returns:
      chev_daily: list of (mass, value) per day [0-indexed = day-1]
      mkt_daily:  list of (mass, value) per day
      peers:      dict company_name → (total_mass, total_value)
    """
    rows = list(ws.iter_rows(values_only=True))

    # Find Chevron row and TOTAL row
    chev_row  = None
    total_row = None
    company_rows = []

    for r in rows:
        if not r[0]:
            continue
        name = str(r[0]).strip()
        if CHEVRON_FRAG in name:
            chev_row = r
        if name == 'TOTAL':
            total_row = r
        elif name not in ('Main Menu', 'Company') and not name.startswith('Previous') and name != 'TOTAL':
            company_rows.append(r)

    if chev_row is None or total_row is None:
        return [], [], {}

    chev_daily = []
    mkt_daily  = []

    for d in range(1, max_day + 1):
        c = 1 + (d - 1) * STRIDE
        cm = n(chev_row[c])
        cv = n(chev_row[c + 1])
        mm = n(total_row[c])
        mv = n(total_row[c + 1])
        chev_daily.append((cm, cv))
        mkt_daily.append((mm, mv))

    # Peers (seasonal total per company)
    peers = {}
    for r in company_rows:
        name = str(r[0]).strip()
        pm = sum(n(r[1 + (d - 1) * STRIDE]) for d in range(1, max_day + 1))
        pv = sum(n(r[2 + (d - 1) * STRIDE]) for d in range(1, max_day + 1))
        if pm > 0:
            peers[name] = (pm, pv)

    return chev_daily, mkt_daily, peers


# ── Step 3: Compute regional D object ─────────────────────────────────────────
def compute_region(region_key, cfg, chev_daily, mkt_daily, peers_raw,
                   day_dates, max_day):
    """Compute the full D[region] structure."""
    DAY = max_day   # latest day number

    # ── Determine period boundaries ──────────────────────────────────────────
    latest_date = day_dates.get(DAY)

    # WTD: from Monday of latest day's week
    if latest_date:
        monday = latest_date - timedelta(days=latest_date.weekday())
        wtd_start = next((d for d in range(1, DAY + 1) if day_dates.get(d, date(2000,1,1)) >= monday), 1)
    else:
        wtd_start = max(1, DAY - 4)

    # MTD: from 1st of latest day's month
    if latest_date:
        month_start = date(latest_date.year, latest_date.month, 1)
        mtd_start = next((d for d in range(1, DAY + 1) if day_dates.get(d, date(2000,1,1)) >= month_start), 1)
    else:
        mtd_start = max(1, DAY - 19)

    # March / April period boundaries for weekly breakdown
    # March = D1 to last day before April 1; April = from April 1 to DAY
    apr_start = None
    if day_dates:
        apr_start = next((d for d in range(1, DAY + 1)
                          if day_dates.get(d, date(2000,1,1)).month == 4), None)
    mar_end = (apr_start - 1) if apr_start and apr_start > 1 else DAY

    def sum_range(daily, start, end):
        """Sum (mass, value) for day-range [start, end] (1-indexed)."""
        tm = tv = 0.0
        for d in range(start, end + 1):
            m, v = daily[d - 1]
            tm += m; tv += v
        return tm, tv

    # ── Checkpoint arrays [Day, WTD, MTD, STD] ──────────────────────────────
    def checkpoints(daily):
        day_m, day_v = daily[DAY - 1]
        wtd_m, wtd_v = sum_range(daily, wtd_start, DAY)
        mtd_m, mtd_v = sum_range(daily, mtd_start, DAY)
        std_m, std_v = sum_range(daily, 1, DAY)
        return (
            [day_m, wtd_m, mtd_m, std_m],
            [day_v, wtd_v, mtd_v, std_v],
        )

    chev_mass, chev_val = checkpoints(chev_daily)
    mkt_mass,  mkt_val  = checkpoints(mkt_daily)

    def safe_prices(masses, values):
        return [round(wgt_price(m, v), 4) for m, v in zip(masses, values)]

    chev_price = safe_prices(chev_mass, chev_val)
    mkt_price  = safe_prices(mkt_mass,  mkt_val)
    price_gap  = [round(c - m, 4) if c and m else None
                  for c, m in zip(chev_price, mkt_price)]

    # ── STD summary ──────────────────────────────────────────────────────────
    std_chev_m, std_chev_v = chev_mass[3], chev_val[3]
    std_mkt_m,  std_mkt_v  = mkt_mass[3],  mkt_val[3]
    std_chev_p = wgt_price(std_chev_m, std_chev_v)
    std_mkt_p  = wgt_price(std_mkt_m,  std_mkt_v)
    price_gap_std = round(std_chev_p - std_mkt_p, 4)

    # ── Pace curves ──────────────────────────────────────────────────────────
    crop25 = cfg['crop25']
    target = cfg['target']

    # Cumulative masses per day
    chev_cum = []; mkt_cum = []
    cc = mc = 0.0
    for i in range(DAY):
        cc += chev_daily[i][0]
        mc += mkt_daily[i][0]
        chev_cum.append(cc)
        mkt_cum.append(mc)

    # Normalise: mkt vs crop25, chev vs target
    mkt_pace  = [round(mc / crop25 * 100, 3) for mc in mkt_cum]
    # Chevron pace: where chev had 0 that day show 0 not null
    chev_pace = [round(cc / target * 100, 3) if cc > 0 else 0.0 for cc in chev_cum]

    # Make final values exactly 100 if it reaches crop25/target
    day_labels = [f'D{d}' for d in range(1, DAY + 1)]

    # ── Pace gap ─────────────────────────────────────────────────────────────
    chev_pace_std = chev_cum[-1] / target * 100 if target else 0
    mkt_pace_std  = mkt_cum[-1]  / crop25 * 100 if crop25 else 0
    pace_gap_pp   = round(chev_pace_std - mkt_pace_std, 2)

    # ── Weekly breakdown ─────────────────────────────────────────────────────
    weekly = []
    if latest_date:
        # March period
        mar_cm, mar_cv = sum_range(chev_daily, 1, mar_end)
        mar_mm, mar_mv = sum_range(mkt_daily,  1, mar_end)
        mar_chev_p = wgt_price(mar_cm, mar_cv) or None
        mar_mkt_p  = wgt_price(mar_mm, mar_mv) or None
        mar_diff   = round(mar_chev_p - mar_mkt_p, 4) if mar_chev_p and mar_mkt_p else None
        mar_loss   = round(mar_cm * mar_diff, 0) if mar_diff and mar_cm else None
        weekly.append({'wk': '2026 Mar (D1-19)',
                       'mkt': mar_mkt_p, 'chev': mar_chev_p,
                       'mass': int(mar_cm), 'diff': mar_diff, 'loss': mar_loss})

        # April period
        if apr_start:
            apr_cm, apr_cv = sum_range(chev_daily, apr_start, DAY)
            apr_mm, apr_mv = sum_range(mkt_daily,  apr_start, DAY)
            apr_chev_p = wgt_price(apr_cm, apr_cv) or None
            apr_mkt_p  = wgt_price(apr_mm, apr_mv) or None
            apr_diff   = round(apr_chev_p - apr_mkt_p, 4) if apr_chev_p and apr_mkt_p else None
            apr_loss   = round(apr_cm * apr_diff, 0) if apr_diff and apr_cm else None
            weekly.append({'wk': f'2026 Apr (D{apr_start}-{DAY})',
                           'mkt': apr_mkt_p, 'chev': apr_chev_p,
                           'mass': int(apr_cm), 'diff': apr_diff, 'loss': apr_loss})

    season_loss = round(sum(w['loss'] for w in weekly if w['loss'] is not None), 0)

    # ── Risk tags ─────────────────────────────────────────────────────────────
    supply = '✓ ON TRACK' if pace_gap_pp >= 0 else '⚠ LAGGING'
    price_risk = '⚠ UNDERPAYING' if price_gap_std < -0.05 else '✓ COMPETITIVE'
    is_lagging = pace_gap_pp < 0
    is_under   = price_gap_std < -0.05
    if is_lagging and is_under:
        twin = '🔴 TWIN RISK'
    elif is_lagging or is_under:
        twin = '⚠ WATCH'
    else:
        twin = '✓ COMPETITIVE'

    # ── Peers list ────────────────────────────────────────────────────────────
    peers = []
    for name, (pm, pv) in sorted(peers_raw.items(), key=lambda x: -wgt_price(x[1][0], x[1][1])):
        peers.append({
            'n': name,
            'std': int(round(pm)),
            'price': round(wgt_price(pm, pv), 4),
            'me': CHEVRON_FRAG in name,
        })

    return {
        'col':          cfg['col'],
        'crop25':       crop25,
        'target':       target,
        'twin':         twin,
        'supply':       supply,
        'price_risk':   price_risk,
        'pace_gap_pp':  pace_gap_pp,
        'price_gap_std': price_gap_std,
        'mkt_mass':     [int(round(v)) for v in mkt_mass],
        'chev_mass':    [int(round(v)) for v in chev_mass],
        'mkt_price':    mkt_price,
        'chev_price':   chev_price,
        'price_gap':    price_gap,
        'mkt_pace_by_day':  mkt_pace,
        'chev_pace_by_day': chev_pace,
        'day_labels':   day_labels,
        'weekly':       weekly,
        'season_loss':  int(season_loss),
        'mkt_std':      int(round(std_mkt_m)),
        'mkt_price_std': round(std_mkt_p, 4),
        'peers':        peers,
    }


# ── Step 4: Build RA_DATA JS block ────────────────────────────────────────────
def build_ra_data(D_data, cr_data):
    """Produce JS: const D = {...}; const cr = [...]"""
    d_json = json.dumps(D_data, separators=(',', ':'))
    cr_json = json.dumps(cr_data, separators=(',', ':'))
    return f"const D={d_json};\nconst cr={cr_json};"


# ── Step 5: Build OV_DATA JS block ────────────────────────────────────────────
def build_ov_data(chev_daily_all, day_dates, max_day):
    """
    chev_daily_all: dict region→list of (mass, value) per day
    Produces: const cd=[...]; const DD={...};
    """
    # cd: daily total Chevron mass across all regions (only days with any volume)
    cd = []
    for d in range(1, max_day + 1):
        total = sum(chev_daily_all[r][d - 1][0] for r in chev_daily_all)
        if total > 0:
            dt = day_dates.get(d)
            cd.append({'d': d, 'm': int(round(total))})

    # DD: day → short date label (e.g. '29Apr')
    DD = {}
    for d, dt in sorted(day_dates.items()):
        if d <= max_day:
            DD[d] = dt.strftime('%-d%b')   # Linux: no zero-pad

    cd_json = json.dumps(cd, separators=(',', ':'))
    dd_json = json.dumps(DD, separators=(',', ':'))
    return f"const cd={cd_json};\nconst DD={dd_json};"



# ── Step 6: Extract P0 Season Overview data from xlsx ────────────────────────
# 2025 per-floor auction hardcoded constants (closed season, will not change)
_TSF25_M  = 3970000;  _TSF25_V  = 13748.11 * 1000  # 3.97M kg, $3.463/kg
_PTSF25_M = 2520000;  _PTSF25_V = 8489.88 * 1000   # 2.52M kg, $3.369/kg
# ETF had no 2025 equivalent

def extract_p0_overview(wb):
    """Read season-level overview data from the workbook.

    Returns a dict with:
      nat26_m, nat26_p, nat25_m, nat25_p  (national crop)
      con26_m, con26_p, con25_m, con25_p  (contract)
      auc26_m, auc26_p, auc25_m, auc25_p  (auction)
      tsf26_m, tsf26_v, tsf26_p           (TSF floor 2026)
      ptf26_m, ptf26_v, ptf26_p           (PTSF floor 2026)
      etf26_m, etf26_v, etf26_p           (ETF floor 2026)
      regions: list of {name, co, m, v, p} 2026 only
    """
    def n(v):
        try: return float(v) if v is not None else 0.0
        except: return 0.0

    def safe_price(m, v):
        return v / m if m > 0 else 0.0

    out = {}

    # ── Auction and Contract sheet (national + per-floor 2026) ────────────────
    ws = wb['Auction and Contract']
    rows = list(ws.iter_rows(values_only=True, min_row=5, max_row=10))
    # row[0] = header (SEASONAL, TSF, ESF, PTF, TOTAL AUCTION, TOTAL CONTRACT, TOTAL 2026, TOTAL 2025, ...)
    # row[1] = mass row
    # row[2] = value row
    # row[3] = avg price row
    mass_row  = rows[1]  # index 1 = row 6
    val_row   = rows[2]  # index 2 = row 7
    # Columns: 0=label, 1=TSF, 2=ESF(ETF), 3=PTF(PTSF), 4=TOTAL_AUCTION, 5=TOTAL_CONTRACT, 6=TOTAL_2026, 7=TOTAL_2025
    tsf26_m = n(mass_row[1]);  tsf26_v = n(val_row[1])
    etf26_m = n(mass_row[2]);  etf26_v = n(val_row[2])
    ptf26_m = n(mass_row[3]);  ptf26_v = n(val_row[3])
    auc26_m = n(mass_row[4]);  auc26_v = n(val_row[4])
    con26_m = n(mass_row[5]);  con26_v = n(val_row[5])
    nat26_m = n(mass_row[6]);  nat26_v = n(val_row[6])
    nat25_m = n(mass_row[7]);  nat25_v = n(val_row[7])

    out['tsf26_m'] = tsf26_m;  out['tsf26_v'] = tsf26_v;  out['tsf26_p'] = safe_price(tsf26_m, tsf26_v)
    out['etf26_m'] = etf26_m;  out['etf26_v'] = etf26_v;  out['etf26_p'] = safe_price(etf26_m, etf26_v)
    out['ptf26_m'] = ptf26_m;  out['ptf26_v'] = ptf26_v;  out['ptf26_p'] = safe_price(ptf26_m, ptf26_v)

    # ── Seasonal Auction summary (2026 + 2025) ────────────────────────────────
    ws = wb['Seasonal Auction summary']
    rows = list(ws.iter_rows(values_only=True, min_row=5, max_row=9))
    auc_mass = rows[1];  auc_val = rows[2]
    auc26_m2 = n(auc_mass[1]);  auc26_v2 = n(auc_val[1])
    auc25_m  = n(auc_mass[2]);  auc25_v  = n(auc_val[2])
    # Prefer direct total from Auction and Contract sheet; use Seasonal for 2025
    out['auc26_m'] = auc26_m if auc26_m > 0 else auc26_m2
    out['auc26_p'] = safe_price(out['auc26_m'], auc26_v)
    out['auc25_m'] = auc25_m
    out['auc25_p'] = safe_price(auc25_m, auc25_v)

    # ── Seasonal Contract Sales (2026 + 2025) ─────────────────────────────────
    ws = wb['Seasonal Contract Sales']
    rows = list(ws.iter_rows(values_only=True, min_row=5, max_row=9))
    con_mass = rows[1];  con_val = rows[2]
    con25_m = n(con_mass[2]);  con25_v = n(con_val[2])
    out['con26_m'] = con26_m
    out['con26_p'] = safe_price(con26_m, con26_v)
    out['con25_m'] = con25_m
    out['con25_p'] = safe_price(con25_m, con25_v)

    # ── National totals (combine contract + auction 2026; use Auction and Contract for 2025) ──
    out['nat26_m'] = nat26_m
    out['nat26_p'] = safe_price(nat26_m, nat26_v)
    out['nat25_m'] = nat25_m
    out['nat25_p'] = safe_price(nat25_m, nat25_v)

    # ── SELLING POINT SUMMARY (regional 2026 only) ────────────────────────────
    ws = wb['SELLING POINT SUMMARY']
    rows = list(ws.iter_rows(values_only=True, min_row=5, max_row=20))
    # header row[0] = (SELLING POINT, No. OF CONTRACTORS, MASS (KG), VALUE (US$), AVE (US$/KG))
    regions = []
    EXCLUDE = {'GRAND TOTAL', 'Previous', 'Main Menu', 'Next', 'SELLING POINT', 'MUTOKO'}
    for r in rows[1:]:  # skip header
        if r[0] is None: continue
        nm = str(r[0]).strip()
        if nm in EXCLUDE or nm == '': continue
        co = int(n(r[1])); m = n(r[2]); v = n(r[3]); p = n(r[4])
        if m > 0:
            regions.append({'name': nm, 'co': co, 'm': m, 'v': v, 'p': p})
    out['regions'] = regions

    return out


# ── Step 7: Build P0 header HTML ─────────────────────────────────────────────
def build_p0_hdr_html(day, end_date):
    """end_date: datetime.date object for last selling day"""
    import datetime
    month_abbr = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
                  7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
    date_str = str(end_date.day).zfill(2) + ' ' + month_abbr[end_date.month]
    return (
        '      <p>Chevron Tobacco &nbsp;&middot;&nbsp; Season Day 1&ndash;' + str(day) +
        ' &nbsp;&middot;&nbsp; 04 Mar &ndash; ' + date_str + ' 2026</p>\n'
        '    </div>\n'
        '  </div>\n'
        '  <div style="display:flex;gap:9px;align-items:center">\n'
        '    <span class="badge">DAY ' + str(day) + '</span>\n'
        '    <span style="font-size:11px;color:var(--mut)">2026 vs 2025 Day 1&ndash;' + str(day) + ' equivalent</span>\n'
        '  </div>\n'
    )


# ── Step 8: Build P0 overview body HTML ──────────────────────────────────────
def build_p0_ov_html(ov, day, end_date):
    """Build the full P0_OV injection block (topbar + overview table + floors + regions)."""
    import datetime
    month_abbr = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
                  7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
    date_str = str(end_date.day).zfill(2) + ' ' + month_abbr[end_date.month]

    def fmt_m(kg):
        """e.g. 211.2M"""
        return '{:.1f}M'.format(kg / 1e6)

    def fmt_mn(kg):
        """e.g. 13.8M with 1dp"""
        return '{:.2f}M'.format(kg / 1e6)

    def fmt_p(p):
        """e.g. $2.631"""
        return '${:.3f}'.format(p)

    def fmt_v(usd):
        """e.g. $545.7M"""
        return '${:.1f}M'.format(usd / 1e6)

    def pct(new_v, old_v):
        if old_v == 0: return ('', 'up')
        c = (new_v - old_v) / old_v * 100
        if c >= 0:
            return ('&#9650; +{:.1f}%'.format(c), 'up')
        else:
            return ('&#9660; &minus;{:.1f}%'.format(abs(c)), 'dn')

    nat26m = ov['nat26_m']; nat25m = ov['nat25_m']
    nat26p = ov['nat26_p']; nat25p = ov['nat25_p']
    con26m = ov['con26_m']; con25m = ov['con25_m']
    con26p = ov['con26_p']; con25p = ov['con25_p']
    auc26m = ov['auc26_m']; auc25m = ov['auc25_m']
    auc26p = ov['auc26_p']; auc25p = ov['auc25_p']

    tsf26m = ov['tsf26_m']; tsf26v = ov['tsf26_v']; tsf26p = ov['tsf26_p']
    ptf26m = ov['ptf26_m']; ptf26v = ov['ptf26_v']; ptf26p = ov['ptf26_p']
    etf26m = ov['etf26_m']; etf26v = ov['etf26_v']; etf26p = ov['etf26_p']

    tsf25m = _TSF25_M;  tsf25v = _TSF25_V;  tsf25p = tsf25v / tsf25m
    ptf25m = _PTSF25_M; ptf25v = _PTSF25_V; ptf25p = ptf25v / ptf25m

    # Change tags
    nat_mc, nat_mc_cls = pct(nat26m, nat25m)
    nat_pc, nat_pc_cls = pct(nat26p, nat25p)
    con_mc, con_mc_cls = pct(con26m, con25m)
    con_pc, con_pc_cls = pct(con26p, con25p)
    auc_mc, auc_mc_cls = pct(auc26m, auc25m)
    auc_pc, auc_pc_cls = pct(auc26p, auc25p)
    tsf_mc, tsf_mc_cls = pct(tsf26m, tsf25m)
    tsf_pc, tsf_pc_cls = pct(tsf26p, tsf25p)
    ptf_mc, ptf_mc_cls = pct(ptf26m, ptf25m)
    ptf_pc, ptf_pc_cls = pct(ptf26p, ptf25p)

    ROW_STYLE = 'display:grid;grid-template-columns:52px 1fr 1fr 1fr'
    CELL_MUT  = 'padding:9px 0 9px 13px;font-size:11px;color:var(--mut);font-weight:600;display:flex;align-items:center'
    CELL_25   = 'padding:9px 10px;text-align:right;font-size:12px;color:var(--mut)'
    CELL_CHG  = 'padding:9px 10px;text-align:right;font-size:12px;font-weight:700'

    h = ''
    # ── p0-topbar ─────────────────────────────────────────────────────────────
    h += '<div class="p0-topbar">\n'
    h += '  <div class="p0-topbar-title">TIMB FCV &mdash; 2026 SEASON OVERVIEW</div>\n'
    h += '  <div class="p0-topbar-sub">Season Day ' + str(day) + ' &nbsp;&middot;&nbsp; Full Season &nbsp;&middot;&nbsp; 04 Mar &ndash; ' + date_str + ' 2026</div>\n'
    h += '</div>\n'
    h += '<div class="p0-body">\n\n'

    # ── Overview label ────────────────────────────────────────────────────────
    h += '  <!-- ====== OVERVIEW SUMMARY TABLE ====== -->\n'
    h += '  <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#888;margin-bottom:10px">'
    h += 'Overview &mdash; 2026 vs 2025 (Day 1&ndash;' + str(day) + ' equivalent)</div>\n'
    h += '  <div class="ovw-grid">\n\n'

    def ovw_card(cls, icon_cls, icon, title, m25, p25, m26, p26, color, mc, mc_cls, pc, pc_cls):
        s  = '    <div class="ovw-card ' + cls + '">\n'
        s += '      <div class="ovw-head ' + cls + '"><div class="ovw-icon ' + icon_cls + '">' + icon + '</div>' + title + '</div>\n'
        s += '      <div class="ovw-body">\n'
        s += '        <div style="' + ROW_STYLE + ';border-bottom:1px solid var(--brd)">\n'
        s += '          <div></div>\n'
        s += '          <div class="col-hdr" style="padding:5px 10px;font-size:10px;color:var(--mut);font-weight:600;text-align:right;display:flex;align-items:center;justify-content:flex-end">2025</div>\n'
        s += '          <div class="col-hdr" style="padding:5px 10px;font-size:10px;color:var(--mut);font-weight:600;text-align:right;display:flex;align-items:center;justify-content:flex-end">2026</div>\n'
        s += '          <div class="col-hdr" style="padding:5px 10px;font-size:10px;color:var(--mut);font-weight:600;text-align:right;display:flex;align-items:center;justify-content:flex-end">Change</div>\n'
        s += '        </div>\n'
        s += '        <div style="' + ROW_STYLE + ';border-bottom:1px solid var(--brd)">\n'
        s += '          <div class="ovw-row" style="display:contents"><div style="' + CELL_MUT + '">Mass</div>'
        s += '<div style="' + CELL_25 + '">' + fmt_m(m25) + '</div>'
        s += '<div style="padding:9px 10px;text-align:right;font-size:13px;font-weight:700;color:' + color + '">' + fmt_m(m26) + '</div>'
        s += '<div style="' + CELL_CHG + '" class="' + mc_cls + '">' + mc + '</div></div>\n'
        s += '        </div>\n'
        s += '        <div style="' + ROW_STYLE + '">\n'
        s += '          <div style="display:contents"><div style="' + CELL_MUT + '">Price</div>'
        s += '<div style="' + CELL_25 + '">' + fmt_p(p25) + '</div>'
        s += '<div style="padding:9px 10px;text-align:right;font-size:13px;font-weight:700;color:' + color + '">' + fmt_p(p26) + '</div>'
        s += '<div style="' + CELL_CHG + '" class="' + pc_cls + '">' + pc + '</div></div>\n'
        s += '        </div>\n'
        s += '      </div>\n'
        s += '    </div>\n\n'
        return s

    h += ovw_card('nat', 'ni', '&Sigma;', 'National Crop',
                  nat25m, nat25p, nat26m, nat26p, '#d4c060',
                  nat_mc, nat_mc_cls, nat_pc, nat_pc_cls)
    h += ovw_card('con', 'ci', 'C', 'Contract',
                  con25m, con25p, con26m, con26p, 'var(--tel)',
                  con_mc, con_mc_cls, con_pc, con_pc_cls)
    h += ovw_card('auc', 'ai', 'A', 'Auction',
                  auc25m, auc25p, auc26m, auc26p, 'var(--blu)',
                  auc_mc, auc_mc_cls, auc_pc, auc_pc_cls)

    h += '  </div>\n\n'  # /ovw-grid

    # ── Auction section ────────────────────────────────────────────────────────
    auc_mvol, auc_mvol_cls = pct(auc26m, auc25m)
    auc_mprc, auc_mprc_cls = pct(auc26p, auc25p)
    h += '  <!-- ====== AUCTION SECTION ====== -->\n'
    h += '  <div class="sec-hdr">\n'
    h += '    <div class="sec-icon si-a">A</div>\n'
    h += '    <h2>AUCTION</h2>\n'
    h += '    <div class="sec-totals">\n'
    h += '      <span>2026: <strong>' + fmt_mn(auc26m) + ' kg</strong></span>\n'
    h += '      <span>2025: <strong>' + fmt_mn(auc25m) + ' kg</strong></span>\n'
    h += '      <span class="' + auc_mvol_cls + '" style="font-weight:700">' + auc_mvol + ' vol</span>\n'
    h += '      <span class="' + auc_mprc_cls + '" style="font-weight:700">' + auc_mprc + ' price</span>\n'
    h += '    </div>\n'
    h += '  </div>\n\n'

    def floor_card(name, note, m26, v26, p26, m25, v25, p25, is_new=False):
        mc, mc_cls = pct(m26, m25) if not is_new else ('', 'up')
        pc, pc_cls = pct(p26, p25) if not is_new else ('', 'dn')
        s  = '    <div class="fc fl">\n'
        s += '      <div class="fn">' + name
        if note:
            s += ' <span style="font-weight:400;font-size:9px;opacity:.55">' + note + '</span>'
        s += '</div>\n'
        s += '      <div class="cmp">\n'
        s += '        <div><div class="yr">2026</div>'
        s += '<div class="fv y6b">' + fmt_mn(m26) + ' <span style="font-size:11px">kg</span></div>'
        s += '<div class="fp">' + fmt_v(v26) + ' &nbsp;&middot;&nbsp; ' + fmt_p(p26) + '/kg</div></div>\n'
        if is_new:
            s += '        <div><div class="yr">2025</div><div class="fv y5">&mdash;</div>'
            s += '<div class="fp" style="font-size:10px">No 2025 equivalent</div></div>\n'
        else:
            s += '        <div><div class="yr">2025</div>'
            s += '<div class="fv y5">' + fmt_mn(m25) + ' <span style="font-size:11px">kg</span></div>'
            s += '<div class="fp">' + fmt_v(v25) + ' &nbsp;&middot;&nbsp; ' + fmt_p(p25) + '/kg</div></div>\n'
        s += '      </div>\n'
        s += '      <div class="tags">'
        if is_new:
            s += '<span class="tag twn">New floor in 2026</span>'
        else:
            s += '<span class="tag t' + mc_cls + '">' + mc + ' vol</span>'
            s += '<span class="tag t' + pc_cls + '">' + pc + ' price</span>'
        s += '</div>\n'
        s += '    </div>\n'
        return s

    h += '  <div class="fc-grid">\n'
    h += floor_card('TSF Floor', None, tsf26m, tsf26v, tsf26p, tsf25m, tsf25v, tsf25p)
    h += floor_card('PTSF Floor', '= Premier 2025', ptf26m, ptf26v, ptf26p, ptf25m, ptf25v, ptf25p)
    h += floor_card('ETF Floor', 'New 2026', etf26m, etf26v, etf26p, 0, 0, 0, is_new=True)
    h += '  </div>\n\n'

    # ── Contract section ───────────────────────────────────────────────────────
    con_mvol, con_mvol_cls = pct(con26m, con25m)
    con_mprc, con_mprc_cls = pct(con26p, con25p)
    h += '  <!-- ====== CONTRACT SECTION ====== -->\n'
    h += '  <div class="sec-hdr">\n'
    h += '    <div class="sec-icon si-c">C</div>\n'
    h += '    <h2>CONTRACT</h2>\n'
    h += '    <div class="sec-totals">\n'
    h += '      <span>2026: <strong>' + fmt_m(con26m) + ' kg</strong></span>\n'
    h += '      <span>2025: <strong>' + fmt_m(con25m) + ' kg</strong></span>\n'
    h += '      <span class="' + con_mvol_cls + '" style="font-weight:700">' + con_mvol + ' vol</span>\n'
    h += '      <span class="' + con_mprc_cls + '" style="font-weight:700">' + con_mprc + ' price</span>\n'
    h += '    </div>\n'
    h += '  </div>\n\n'

    h += '  <div class="fc-grid" style="grid-template-columns:repeat(3,1fr)">\n'
    for reg in ov['regions']:
        nm  = reg['name']; co = reg['co']; m = reg['m']; v = reg['v']; p = reg['p']
        s  = '    <div class="fc rg">\n'
        s += '      <div class="fn">' + nm + '</div>\n'
        s += '      <div class="cmp">\n'
        s += '        <div><div class="yr">2026</div>'
        s += '<div class="fv y6g">' + fmt_m(m) + ' <span style="font-size:11px">kg</span></div>'
        s += '<div class="fp">' + fmt_v(v) + ' &nbsp;&middot;&nbsp; ' + fmt_p(p) + '/kg &nbsp;&middot;&nbsp; ' + str(co) + ' co</div></div>\n'
        s += '      </div>\n'
        s += '      <div class="tags"><span class="tag twn">2026 Season-to-Date</span></div>\n'
        s += '    </div>\n'
        h += s
    h += '  </div>\n'

    return h


# ── Step 9: Inject all blocks into HTML ───────────────────────────────────────
def _inject_html_block(html, start_marker, end_marker, content):
    s = html.find(start_marker)
    e = html.find(end_marker)
    if s == -1 or e == -1:
        print('  WARNING: marker not found: ' + start_marker)
        return html, False
    return html[:s + len(start_marker)] + '\n' + content + '\n' + html[e:], True


def inject_into_dashboard(ra_js, ov_js, p0_hdr_html, p0_ov_html):
    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()

    changed = False

    # JS data blocks (script section)
    for ms, me, new_js in [
        ('/* RA_DATA_START */', '/* RA_DATA_END */', ra_js),
        ('/* OV_DATA_START */', '/* OV_DATA_END */', ov_js),
    ]:
        if ms not in html:
            print('  WARNING: JS marker not found: ' + ms)
            continue
        start = html.index(ms) + len(ms)
        end   = html.index(me)
        html  = html[:start] + '\n' + new_js + '\n' + html[end:]
        changed = True

    # P0 HTML blocks
    html, ok = _inject_html_block(html, '<!-- P0_HDR_START -->', '<!-- P0_HDR_END -->', p0_hdr_html)
    if ok: changed = True
    html, ok = _inject_html_block(html, '<!-- P0_OV_START -->', '<!-- P0_OV_END -->', p0_ov_html)
    if ok: changed = True

    if changed:
        with open(DASHBOARD, 'w', encoding='utf-8') as f:
            f.write(html)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("generate_dashboard_data.py -- Regional & Overview Live Data")
    print("=" * 60)

    # 1. Find latest file
    latest_path, max_day = find_latest_file()
    print("\n  Latest file: Day " + str(max_day) + " -- " + os.path.basename(latest_path))

    # 2. Build day -> date mapping
    print("  Building day->date mapping from all daily files...")
    day_dates = build_day_dates()
    print("  Found dates for " + str(len(day_dates)) + " days")

    # 3. Load latest workbook
    print("  Loading latest daily FCV file...")
    wb = load_wb(latest_path)

    # 4. Process each region
    print("  Processing regional sheets...")
    D_data = {}
    cr_data = []
    chev_daily_all = {}

    for region_key, cfg in REGIONS.items():
        sheet_name = cfg['sheet']
        if sheet_name not in wb.sheetnames:
            print("  WARNING: Sheet '" + sheet_name + "' not found -- skipping " + region_key)
            continue

        ws = wb[sheet_name]
        chev_daily, mkt_daily, peers_raw = read_regional_sheet(ws, max_day)

        chev_daily_all[region_key] = chev_daily

        region_data = compute_region(
            region_key, cfg, chev_daily, mkt_daily, peers_raw, day_dates, max_day
        )
        D_data[region_key] = region_data

        chev_std = region_data['chev_mass'][3]
        mkt_std  = region_data['mkt_std']
        cr_data.append({'r': region_key, 'm': chev_std})

        print("    {:10s}: Chev={:>8,} kg  Mkt={:>12,} kg  Peers={}  PaceGap={:+.1f}pp".format(
            region_key, chev_std, mkt_std, len(region_data['peers']), region_data['pace_gap_pp']))

    # 5. Build JS blocks
    print("\n  Building JS data blocks...")
    ra_js = build_ra_data(D_data, cr_data)
    ov_js = build_ov_data(chev_daily_all, day_dates, max_day)

    # 6. Extract P0 overview data
    print("  Extracting P0 season overview data...")
    ov_data = extract_p0_overview(wb)
    regions_read = [r['name'] for r in ov_data.get('regions', [])]
    print("  Regions from SELLING POINT SUMMARY: " + ', '.join(regions_read))
    print("  National 2026: {:.1f}M kg @ ${:.3f}/kg".format(
        ov_data['nat26_m']/1e6, ov_data['nat26_p']))
    print("  National 2025: {:.1f}M kg @ ${:.3f}/kg".format(
        ov_data['nat25_m']/1e6, ov_data['nat25_p']))

    # 7. Build P0 HTML blocks
    end_date = day_dates.get(max_day)
    if end_date is None:
        from datetime import date
        end_date = date.today()
    p0_hdr_html = build_p0_hdr_html(max_day, end_date)
    p0_ov_html  = build_p0_ov_html(ov_data, max_day, end_date)

    # 8. Inject into dashboard
    print("  Injecting into dashboard HTML...")
    inject_into_dashboard(ra_js, ov_js, p0_hdr_html, p0_ov_html)
    print("\n  ✓ Dashboard updated successfully.")
    print("    Dashboard: " + DASHBOARD)
    print("=" * 60)


if __name__ == '__main__':
    main()
