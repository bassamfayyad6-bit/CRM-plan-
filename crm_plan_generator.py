"""
CRM Plan Generator
==================
Input : Cold_Rolling_Schedule_*.xlsx  (Sheet 1 = master plan, Sheet 2 = CRM queue)
Output: CRM_Plan_<date>.xlsx

Structure:
  Rows 1–100   : priority-sorted passes (fewest remaining → first)
  Row  101     : ⚠️  CHANGE WORK ROLL  note  (yellow)
  Rows 102–104 : 3 warm-up coils  (thick P1/P2/P3, passes_left ≥ 3)
                 with note: "warm up the mill — take surface sample after 1st pass"
  Row  105     : ⚠️  UPDATE PLAN FROM HERE  note  (orange)

Priority logic for main 100:
  1. Fewest CRM passes remaining (from master plan)
  2. Tie-break: earliest Delivery date
  Coils NOT found in master → appended at very end of the 100

Warm-up coil criteria:
  - TH [mm] > 1.5 mm
  - PASS in P1 / P2 / P3
  - passes_left >= 3  (not near the end)
  - NOT already selected in the main 100
  - Sort by TH [mm] descending (thickest first = best for heating rolls)
"""

import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── palette ───────────────────────────────────────────────────────────────────
HEADER_BG  = '1F3864'
HEADER_FG  = 'FFFFFF'
LEGEND_BG  = '2E75B6'
WR_BG      = 'FFD700'   # gold  – Work Roll change
WR_FG      = '7B0000'   # dark red text
UPD_BG     = 'F4B183'   # orange – Update plan
UPD_FG     = '7B0000'
WARM_BG    = 'FFF2CC'   # light yellow – warm-up rows
ROW_1PASS  = 'FFE699'   # amber   – 1 pass left
ROW_2PASS  = 'E2EFDA'   # green   – 2 passes
ROW_UNKN   = 'F2DCDB'   # red     – not in master
ROW_WHITE  = 'FFFFFF'

COLS = [
    ('NO',            8),
    ('COIL Man #',   18),
    ('A',             7),
    ('T.T',           7),
    ('TH [mm]',       9),
    ('Width',         8),
    ('T.W',           8),
    ('Int+Trim',      9),
    ('Targeted Th.', 12),
    ('Steel spool',  11),
    ('PASS',          7),
    ('Previous',     14),
    ('Process',      10),
    ('NEXT',         10),
    ('Passes Left',  11),
    ('Final Dest.',  12),
    ('Delivery date',14),
    ('Notes',        28),
    ('Customer',     28),
]
NCOLS = len(COLS)

# ── helpers ───────────────────────────────────────────────────────────────────
def thin():
    s = Side(style='thin', color='000000')
    return Border(left=s, right=s, top=s, bottom=s)

def med_border(color='7B0000'):
    s = Side(style='medium', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def cs(ws, row, col, value, bold=False, bg=None, fg='000000',
       align='center', size=10, wrap=False, italic=False, border=True,
       num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name='Calibri', bold=bold, color=fg, size=size, italic=italic)
    if bg:
        c.fill = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    if border:
        c.border = thin()
    if num_fmt:
        c.number_format = num_fmt
    return c

def fmt_date(val):
    try:
        return pd.Timestamp(val).strftime('%Y-%m-%d')
    except Exception:
        return str(val) if pd.notna(val) else ''

def safe_str(val):
    return str(val).strip() if pd.notna(val) else ''

def safe_float(val, decimals=3):
    try:
        return round(float(val), decimals)
    except Exception:
        return val

# ── load data ─────────────────────────────────────────────────────────────────
def load_data(filepath):
    master = pd.read_excel(filepath, sheet_name='Rolling Production Plan,', header=3)
    master.columns = [str(c).strip() for c in master.columns]

    crm = pd.read_excel(filepath, sheet_name='CRM ', header=2)
    crm.columns = [str(c).strip() for c in crm.columns]
    crm = crm[crm['COIL Man #'].notna() & crm['PASS'].notna()].copy()
    crm = crm[crm['PASS'].astype(str).str.startswith('P')].copy()
    crm = crm.reset_index(drop=True)
    return master, crm

# ── remaining passes ──────────────────────────────────────────────────────────
def remaining_passes(master, coil_id, cur_pass):
    journey = master[master['COIL Man #'].astype(str).str.strip() ==
                     str(coil_id).strip()].copy()
    if journey.empty:
        return 999, 'UNKNOWN'
    journey = journey.sort_values('NO')
    cm = journey[journey['PASS'].notna() &
                 journey['PASS'].astype(str).str.startswith('P')]
    if cm.empty:
        return 999, 'UNKNOWN'
    cur = cm[cm['PASS'].astype(str).str.strip() == str(cur_pass).strip()]
    if cur.empty:
        return 999, 'UNKNOWN'
    rem = cm[cm['NO'] >= cur.iloc[0]['NO']]
    return len(rem), safe_str(rem.iloc[-1]['NEXT'])

# ── prioritize ────────────────────────────────────────────────────────────────
def prioritize(master, crm):
    rows = []
    for _, r in crm.iterrows():
        pl, fd = remaining_passes(master, r['COIL Man #'], r['PASS'])
        rows.append({**r.to_dict(), '_passes_left': pl, '_final_dest': fd})
    df = pd.DataFrame(rows)

    def del_sort(v):
        try:
            return pd.Timestamp(v).timestamp()
        except Exception:
            return 9e18

    df['_del_sort'] = df['Delivery date'].apply(del_sort)
    df = df.sort_values(['_passes_left', '_del_sort'],
                        ascending=[True, True]).reset_index(drop=True)
    return df

# ── select warm-up coils ──────────────────────────────────────────────────────
def select_warmup(master, crm, used_coils, n=3):
    """
    Pick n warm-up coils: thick (>1.5mm), P1/P2/P3, passes_left >= 3,
    not already in used_coils. Sort by TH descending.
    """
    candidates = []
    for _, r in crm.iterrows():
        coil_id = safe_str(r['COIL Man #'])
        if coil_id in used_coils:
            continue
        th = r.get('TH [mm]', 0)
        cur_pass = safe_str(r['PASS'])
        try:
            th_f = float(th)
        except Exception:
            continue
        if th_f <= 1.5:
            continue
        if cur_pass not in ('P1', 'P2', 'P3'):
            continue
        pl, fd = remaining_passes(master, coil_id, cur_pass)
        if pl < 3 or pl >= 999:
            continue
        candidates.append({**r.to_dict(),
                            '_passes_left': pl,
                            '_final_dest': fd,
                            '_th_f': th_f})

    if not candidates:
        return pd.DataFrame()

    df = pd.DataFrame(candidates).sort_values('_th_f', ascending=False)
    return df.head(n).reset_index(drop=True)

# ── row bg colour ─────────────────────────────────────────────────────────────
def row_bg(passes_left, warmup=False):
    if warmup:
        return WARM_BG
    if passes_left == 1:
        return ROW_1PASS
    if passes_left == 2:
        return ROW_2PASS
    if passes_left >= 999:
        return ROW_UNKN
    return ROW_WHITE

# ── write one data row ────────────────────────────────────────────────────────
def write_data_row(ws, excel_row, no_val, row, warmup=False):
    bg = row_bg(row['_passes_left'], warmup)

    def w(col, val, num_fmt=None):
        cs(ws, excel_row, col, val, bg=bg, align='center', num_fmt=num_fmt)

    ws.row_dimensions[excel_row].height = 17
    w(1,  no_val)
    w(2,  safe_str(row.get('COIL Man #', '')))
    w(3,  row.get('A', ''))
    w(4,  row.get('T.T', ''))
    w(5,  safe_float(row.get('TH [mm]', '')), '0.000')
    w(6,  row.get('Width', ''))
    w(7,  row.get('T.W', ''))
    w(8,  row.get('Int + Final Trim', ''))
    w(9,  safe_float(row.get('Targeted Th.', '')), '0.000')
    w(10, row.get('Steel spool', ''))
    w(11, safe_str(row.get('PASS', '')))
    w(12, safe_str(row.get('Previous', '')) if pd.notna(row.get('Previous')) else '')
    w(13, safe_str(row.get('Process', '')))
    w(14, safe_str(row.get('NEXT', '')) if pd.notna(row.get('NEXT')) else '')

    pl = row['_passes_left']
    pl_cell = ws.cell(row=excel_row, column=15,
                      value=pl if pl < 999 else 'N/A')
    pl_cell.font = Font(name='Calibri', bold=(pl >= 999),
                        color=('C00000' if pl >= 999 else '000000'), size=10)
    pl_cell.fill = PatternFill('solid', start_color=bg)
    pl_cell.alignment = Alignment(horizontal='center', vertical='center')
    pl_cell.border = thin()

    w(16, safe_str(row['_final_dest']))
    w(17, fmt_date(row.get('Delivery date')))

    notes_val = safe_str(row.get('Notes', ''))
    if warmup and not notes_val:
        notes_val = 'Warm-up coil — take surface sample after 1st pass on new WR'
    nc = ws.cell(row=excel_row, column=18, value=notes_val)
    nc.font = Font(name='Calibri', size=9,
                   color=('C00000' if 'ANN' in notes_val.upper() else
                          ('7B5200' if warmup else '000000')))
    nc.fill = PatternFill('solid', start_color=bg)
    nc.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    nc.border = thin()

    cust = safe_str(row.get('Customer', ''))
    cc = ws.cell(row=excel_row, column=19, value=cust)
    cc.font = Font(name='Calibri', size=9)
    cc.fill = PatternFill('solid', start_color=bg)
    cc.alignment = Alignment(horizontal='left', vertical='center')
    cc.border = thin()

# ── banner row (merged, coloured) ─────────────────────────────────────────────
def write_banner(ws, excel_row, text, bg, fg, font_size=12, bdr_color='7B0000'):
    ws.row_dimensions[excel_row].height = 28
    ws.merge_cells(start_row=excel_row, start_column=1,
                   end_row=excel_row, end_column=NCOLS)
    c = ws.cell(row=excel_row, column=1, value=text)
    c.font  = Font(name='Calibri', bold=True, color=fg, size=font_size)
    c.fill  = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = med_border(bdr_color)

# ── build Excel ───────────────────────────────────────────────────────────────
def build_excel(df_100, df_warmup, plan_date):
    wb = Workbook()
    ws = wb.active
    ws.title = 'CRM Plan'

    # title
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    tc = ws.cell(row=1, column=1,
                 value=f'CRM Production Plan  —  {plan_date}  (100 passes)')
    tc.font = Font(name='Calibri', bold=True, color=HEADER_FG, size=14)
    tc.fill = PatternFill('solid', start_color=HEADER_BG)
    tc.alignment = Alignment(horizontal='center', vertical='center')

    # legend
    ws.row_dimensions[2].height = 16
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOLS)
    lc = ws.cell(row=2, column=1,
                 value=('🟡 1 pass left – send to next stage immediately   '
                        '🟢 2 passes left   ⬜ 3+ passes   '
                        '🔴 Not in master plan   🟨 Warm-up coils'))
    lc.font = Font(name='Calibri', italic=True, color=HEADER_FG, size=9)
    lc.fill = PatternFill('solid', start_color=LEGEND_BG)
    lc.alignment = Alignment(horizontal='center', vertical='center')

    # header row
    HDR = 3
    ws.row_dimensions[HDR].height = 22
    for ci, (name, width) in enumerate(COLS, 1):
        cs(ws, HDR, ci, name, bold=True, bg=HEADER_BG, fg=HEADER_FG)
        ws.column_dimensions[get_column_letter(ci)].width = width

    ws.freeze_panes = f'A{HDR + 1}'

    # ── main 100 rows ─────────────────────────────────────────────────────
    DATA_START = HDR + 1
    for idx, row in df_100.iterrows():
        write_data_row(ws, DATA_START + idx, idx + 1, row)

    # ── row 101: CHANGE WORK ROLL ─────────────────────────────────────────
    WR_ROW = DATA_START + 100
    write_banner(ws, WR_ROW,
                 '⚠️   STOP  —  CHANGE WORK ROLL   |   '
                 'وقّف الماكينة وغيّر الـ Work Roll   ⚠️',
                 WR_BG, WR_FG, font_size=13)

    # ── rows 102–104: warm-up coils ───────────────────────────────────────
    WARM_START = WR_ROW + 1

    # warm-up sub-header
    ws.row_dimensions[WARM_START].height = 20
    ws.merge_cells(start_row=WARM_START, start_column=1,
                   end_row=WARM_START, end_column=NCOLS)
    wh = ws.cell(row=WARM_START, column=1,
                 value=('🔥  Warm-up coils  —  roll back to back on new Work Roll  |  '
                        'كويلات التسخين — شغّلهم متتاليين على الـ Work Roll الجديد  '
                        '|  خذ surface sample بعد أول باص'))
    wh.font  = Font(name='Calibri', bold=True, color='7B5200', size=10)
    wh.fill  = PatternFill('solid', start_color='FFE0B2')
    wh.alignment = Alignment(horizontal='center', vertical='center')
    wh.border = med_border('BF8600')

    if not df_warmup.empty:
        for wi, (_, wrow) in enumerate(df_warmup.iterrows()):
            write_data_row(ws, WARM_START + 1 + wi, f'W{wi+1}', wrow, warmup=True)
        next_row_after_warmup = WARM_START + 1 + len(df_warmup)
    else:
        # fallback if no candidates found
        ws.row_dimensions[WARM_START + 1].height = 17
        ws.merge_cells(start_row=WARM_START+1, start_column=1,
                       end_row=WARM_START+1, end_column=NCOLS)
        fb = ws.cell(row=WARM_START+1, column=1,
                     value='No warm-up candidates found — select manually')
        fb.font  = Font(name='Calibri', italic=True, color='C00000', size=10)
        fb.fill  = PatternFill('solid', start_color=WARM_BG)
        fb.alignment = Alignment(horizontal='center', vertical='center')
        fb.border = thin()
        next_row_after_warmup = WARM_START + 2

    # ── final banner: UPDATE PLAN ─────────────────────────────────────────
    write_banner(ws, next_row_after_warmup,
                 '📋   حدّث الخطة من هنا  —  UPDATE THE PLAN FROM HERE   📋',
                 UPD_BG, UPD_FG, font_size=12, bdr_color='BF6000')

    # auto-filter on header
    ws.auto_filter.ref = (f'A{HDR}:'
                          f'{get_column_letter(NCOLS)}{DATA_START + 99}')
    return wb

# ── main ──────────────────────────────────────────────────────────────────────
def main(filepath, output_path=None):
    print(f'Loading: {filepath}')
    master, crm = load_data(filepath)
    print(f'  Master rows: {len(master)} | CRM queue: {len(crm)}')

    print('Computing priorities...')
    df_sorted = prioritize(master, crm)
    df_100    = df_sorted.head(100).reset_index(drop=True)

    used_coils = set(df_100['COIL Man #'].astype(str).str.strip().tolist())

    print('Selecting warm-up coils...')
    df_warmup = select_warmup(master, crm, used_coils, n=3)
    if df_warmup.empty:
        print('  ⚠ No warm-up candidates found')
    else:
        for _, r in df_warmup.iterrows():
            print(f"  Warmup: {r['COIL Man #']}  {r['PASS']}  "
                  f"TH={r['TH [mm]']}mm  passes_left={r['_passes_left']}")

    plan_date = datetime.today().strftime('%Y-%m-%d')
    if output_path is None:
        output_path = f'CRM_Plan_{plan_date}.xlsx'

    print(f'Building Excel → {output_path}')
    wb = build_excel(df_100, df_warmup, plan_date)
    wb.save(output_path)
    print('Done ✓')
    return output_path

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python crm_plan_generator.py <schedule.xlsx> [output.xlsx]')
        sys.exit(1)
    out = sys.argv[2] if len(sys.argv) > 2 else None
    main(sys.argv[1], out)
