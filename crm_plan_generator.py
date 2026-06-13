import streamlit as st
import pandas as pd
import tempfile, os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── colours ───────────────────────────────────────────────────────────────────
HEADER_BG  = '1F3864'; HEADER_FG  = 'FFFFFF'
LEGEND_BG  = '2E75B6'
SEC_FINAL  = 'E2EFDA'   # green   – final pass (1 left)
SEC_PUSH   = 'FFF2CC'   # yellow  – push coils (2 left)
SEC_INTANN = 'DDEEFF'   # blue    – INT Ann
SEC_INTTRIM= 'EDE7F6'   # purple  – INT Trim
SEC_NEW    = 'F5F5F5'   # grey    – new P1/P2
SEC_WARM   = 'FFE0B2'   # orange  – warm-up
NOTE_WR    = 'FFD700'   # gold    – WR change banner
NOTE_SEC   = 'D6E4F0'   # light   – section headers
NOTE_END   = 'F4B183'   # orange  – end / update plan
ROW_UNKN   = 'F2DCDB'   # red     – not in master

COLS = [
    ('NO',8),('COIL Man #',18),('A',7),('T.T',7),('TH [mm]',9),
    ('Width',8),('T.W',8),('Int+Trim',9),('Targeted Th.',12),
    ('Steel spool',11),('PASS',7),('Previous',14),('Process',10),
    ('NEXT',10),('Passes Left',11),('Final Dest.',12),
    ('Delivery date',14),('Notes',30),('Customer',28),
]
NCOLS = len(COLS)

# ── openpyxl helpers ──────────────────────────────────────────────────────────
def thin():
    s = Side(style='thin', color='000000')
    return Border(left=s, right=s, top=s, bottom=s)

def med_border(color='1F3864'):
    s = Side(style='medium', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def cs(ws, row, col, value, bold=False, bg=None, fg='000000',
       align='center', size=10, wrap=False, num_fmt=None, italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name='Calibri', bold=bold, color=fg, size=size, italic=italic)
    if bg:
        c.fill = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    c.border = thin()
    if num_fmt:
        c.number_format = num_fmt
    return c

def banner(ws, r, text, bg, fg='000000', size=11, bdr='1F3864'):
    ws.row_dimensions[r].height = 24
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOLS)
    c = ws.cell(row=r, column=1, value=text)
    c.font  = Font(name='Calibri', bold=True, color=fg, size=size)
    c.fill  = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = med_border(bdr)

def fmt_date(val):
    try:    return pd.Timestamp(val).strftime('%Y-%m-%d')
    except: return str(val) if pd.notna(val) else ''

def safe_str(val):
    return str(val).strip() if pd.notna(val) else ''

def safe_float(val):
    try:    return round(float(val), 3)
    except: return val

# ── data loading ──────────────────────────────────────────────────────────────
def load_data(filepath):
    master = pd.read_excel(filepath, sheet_name='Rolling Production Plan,', header=3)
    master.columns = [str(c).strip() for c in master.columns]
    crm = pd.read_excel(filepath, sheet_name='CRM ', header=2)
    crm.columns = [str(c).strip() for c in crm.columns]
    crm = crm[crm['COIL Man #'].notna() & crm['PASS'].notna()].copy()
    crm = crm[crm['PASS'].astype(str).str.startswith('P')].copy()
    return master, crm.reset_index(drop=True)

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

# ── classify coils into sections ──────────────────────────────────────────────
def classify(master, crm):
    rows = []
    for _, r in crm.iterrows():
        pl, fd = remaining_passes(master, r['COIL Man #'], r['PASS'])
        nxt    = safe_str(r.get('NEXT', ''))
        notes  = safe_str(r.get('Notes', '')).upper()

        # Section assignment
        if pl == 1:
            sec = 'FINAL'
        elif pl == 2:
            sec = 'PUSH'
        elif 'INT' in nxt.upper() and 'ANN' in nxt.upper():
            sec = 'INT_ANN'
        elif 'INT' in nxt.upper() and 'TRIM' in nxt.upper():
            sec = 'INT_TRIM'
        elif pl >= 999:
            sec = 'UNKNOWN'
        else:
            sec = 'NEW'

        def del_ts(v):
            try:    return pd.Timestamp(v).timestamp()
            except: return 9e18

        rows.append({**r.to_dict(),
                     '_passes_left': pl,
                     '_final_dest':  fd,
                     '_section':     sec,
                     '_del_sort':    del_ts(r.get('Delivery date'))})

    df = pd.DataFrame(rows)
    return df

def select_warmup(master, crm, used_coils, n=3):
    candidates = []
    for _, r in crm.iterrows():
        coil_id  = safe_str(r['COIL Man #'])
        if coil_id in used_coils:
            continue
        try:
            th_f = float(r.get('TH [mm]', 0))
        except:
            continue
        cur_pass = safe_str(r['PASS'])
        if th_f <= 1.5 or cur_pass not in ('P1', 'P2', 'P3'):
            continue
        pl, fd = remaining_passes(master, coil_id, cur_pass)
        if pl < 3 or pl >= 999:
            continue
        candidates.append({**r.to_dict(),
                            '_passes_left': pl,
                            '_final_dest':  fd,
                            '_th_f':        th_f})
    if not candidates:
        return pd.DataFrame()
    return (pd.DataFrame(candidates)
              .sort_values('_th_f', ascending=False)
              .head(n).reset_index(drop=True))

# ── build ordered plan ────────────────────────────────────────────────────────
SECTION_ORDER = ['FINAL', 'INT_ANN', 'PUSH', 'INT_TRIM', 'NEW', 'UNKNOWN']

SECTION_META = {
    'FINAL':    ('FINAL PASS COILS — 1 pass remaining, send to next stage immediately',
                 SEC_FINAL,  '1B5E20'),
    'PUSH':     ('PUSH COILS — 2 passes remaining to become final',
                 SEC_PUSH,   '7B5200'),
    'INT_ANN':  ('INT ANNEALING COILS — next step: Intermediate Annealing',
                 SEC_INTANN, '1A3A5C'),
    'INT_TRIM': ('INT TRIM COILS — next step: Intermediate Trimming',
                 SEC_INTTRIM,'4A235A'),
    'NEW':      ('NEW COILS — P1 / P2 first & second pass',
                 SEC_NEW,    '3E3E3E'),
    'UNKNOWN':  ('COILS NOT FOUND IN MASTER PLAN — verify manually',
                 ROW_UNKN,   'C00000'),
}

SEC_ROW_BG = {
    'FINAL':    SEC_FINAL,
    'PUSH':     SEC_PUSH,
    'INT_ANN':  SEC_INTANN,
    'INT_TRIM': SEC_INTTRIM,
    'NEW':      SEC_NEW,
    'UNKNOWN':  ROW_UNKN,
}

def build_plan(master, crm):
    df = classify(master, crm)
    warmup_used = set()

    # warm-up: thick P1/P2/P3 not in any section bucket yet
    df_warmup = select_warmup(master, crm, warmup_used, n=3)
    if not df_warmup.empty:
        warmup_used = set(df_warmup['COIL Man #'].astype(str).str.strip())

    # remove warm-up coils from main pool
    df = df[~df['COIL Man #'].astype(str).str.strip().isin(warmup_used)].copy()

    sections = {}
    for sec in SECTION_ORDER:
        sub = df[df['_section'] == sec].sort_values('_del_sort').reset_index(drop=True)
        sections[sec] = sub

    return df_warmup, sections

# ── write one data row ────────────────────────────────────────────────────────
def write_row(ws, excel_row, no_val, row, bg):
    ws.row_dimensions[excel_row].height = 17

    def w(col, val, num_fmt=None):
        cs(ws, excel_row, col, val, bg=bg, num_fmt=num_fmt)

    w(1,  no_val)
    w(2,  safe_str(row.get('COIL Man #', '')))
    w(3,  row.get('A', ''))
    w(4,  row.get('T.T', ''))
    w(5,  safe_float(row.get('TH [mm]', '')),   '0.000')
    w(6,  row.get('Width', ''))
    w(7,  row.get('T.W', ''))
    w(8,  row.get('Int + Final Trim', ''))
    w(9,  safe_float(row.get('Targeted Th.', '')), '0.000')
    w(10, row.get('Steel spool', ''))
    w(11, safe_str(row.get('PASS', '')))
    w(12, safe_str(row.get('Previous', '')) if pd.notna(row.get('Previous')) else '')
    w(13, safe_str(row.get('Process', '')))
    w(14, safe_str(row.get('NEXT', '')) if pd.notna(row.get('NEXT')) else '')

    pl   = row['_passes_left']
    pl_c = ws.cell(row=excel_row, column=15, value=pl if pl < 999 else 'N/A')
    pl_c.font      = Font(name='Calibri', bold=(pl >= 999),
                          color=('C00000' if pl >= 999 else '000000'), size=10)
    pl_c.fill      = PatternFill('solid', start_color=bg)
    pl_c.alignment = Alignment(horizontal='center', vertical='center')
    pl_c.border    = thin()

    w(16, safe_str(row['_final_dest']))
    w(17, fmt_date(row.get('Delivery date')))

    notes_val = safe_str(row.get('Notes', ''))
    nc = ws.cell(row=excel_row, column=18, value=notes_val)
    nc.font      = Font(name='Calibri', size=9,
                        color='C00000' if 'ANN' in notes_val.upper() else '000000')
    nc.fill      = PatternFill('solid', start_color=bg)
    nc.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    nc.border    = thin()

    cc = ws.cell(row=excel_row, column=19, value=safe_str(row.get('Customer', '')))
    cc.font      = Font(name='Calibri', size=9)
    cc.fill      = PatternFill('solid', start_color=bg)
    cc.alignment = Alignment(horizontal='left', vertical='center')
    cc.border    = thin()

# ── build Excel ───────────────────────────────────────────────────────────────
def build_excel(df_warmup, sections, plan_date):
    wb = Workbook()
    ws = wb.active
    ws.title = 'CRM Plan'

    # title
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    tc = ws.cell(row=1, column=1,
                 value=f'CRM Production Plan  —  {plan_date}  (100 passes)')
    tc.font      = Font(name='Calibri', bold=True, color=HEADER_FG, size=14)
    tc.fill      = PatternFill('solid', start_color=HEADER_BG)
    tc.alignment = Alignment(horizontal='center', vertical='center')

    # legend
    ws.row_dimensions[2].height = 15
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOLS)
    lc = ws.cell(row=2, column=1,
                 value=('🟢 Final Pass (1 left)   '
                        '🟡 Push Coils (2 left)   '
                        '🔵 INT Ann   '
                        '🟣 INT Trim   '
                        '⬜ New Coils   '
                        '🔴 Not in master   '
                        '🟠 Warm-up'))
    lc.font      = Font(name='Calibri', italic=True, color=HEADER_FG, size=9)
    lc.fill      = PatternFill('solid', start_color=LEGEND_BG)
    lc.alignment = Alignment(horizontal='center', vertical='center')

    # column headers
    HDR = 3
    ws.row_dimensions[HDR].height = 22
    for ci, (name, width) in enumerate(COLS, 1):
        cs(ws, HDR, ci, name, bold=True, bg=HEADER_BG, fg=HEADER_FG)
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.freeze_panes = f'A{HDR + 1}'

    cur_row  = HDR + 1
    pass_cnt = 0
    MAX      = 100

    # ── WR change + warm-up ───────────────────────────────────────────────
    banner(ws, cur_row,
           '⚠️   CHANGE WORK ROLL   —   Warm up the mill with the coils below   ⚠️',
           NOTE_WR, '7B0000', size=12, bdr='BF8600')
    cur_row += 1

    if not df_warmup.empty:
        for _, wrow in df_warmup.iterrows():
            if pass_cnt >= MAX:
                break
            write_row(ws, cur_row, f'W{pass_cnt+1}', wrow, SEC_WARM)
            cur_row  += 1
            pass_cnt += 1
    else:
        banner(ws, cur_row,
               'No warm-up candidates found — select manually',
               SEC_WARM, 'C00000', size=10)
        cur_row += 1

    # ── main sections ─────────────────────────────────────────────────────
    for sec in SECTION_ORDER:
        if pass_cnt >= MAX:
            break
        sub = sections.get(sec, pd.DataFrame())
        if sub.empty:
            continue

        label, sec_bg, sec_fg = SECTION_META[sec]
        row_bg_col = SEC_ROW_BG[sec]

        # section separator
        banner(ws, cur_row, f'▶   {label}', NOTE_SEC, '1F3864', size=10, bdr='2E75B6')
        cur_row += 1

        for _, row in sub.iterrows():
            if pass_cnt >= MAX:
                break
            write_row(ws, cur_row, pass_cnt + 1, row, row_bg_col)
            cur_row  += 1
            pass_cnt += 1

    # ── end note ──────────────────────────────────────────────────────────
    if pass_cnt >= MAX:
        banner(ws, cur_row,
               '📋   100 PASSES REACHED  —  UPDATE THE PLAN BEFORE CONTINUING   📋',
               NOTE_END, '7B0000', size=12, bdr='BF6000')
    else:
        banner(ws, cur_row,
               f'📋   PLAN COMPLETE  —  {pass_cnt} passes scheduled  —  UPDATE PLAN AS NEEDED   📋',
               NOTE_END, '7B0000', size=12, bdr='BF6000')

    ws.auto_filter.ref = (f'A{HDR}:{get_column_letter(NCOLS)}{HDR + 1 + MAX}')
    return wb

# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.set_page_config(page_title='CRM Plan Generator', page_icon='🏭', layout='centered')
st.title('🏭 CRM Plan Generator')

uploaded = st.file_uploader('Upload Full Schedule (Cold Rolling Schedule)', type=['xlsx'])

if uploaded:
    with st.spinner('Processing...'):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded.read())
                tmp_path = tmp.name

            master, crm  = load_data(tmp_path)
            df_warmup, sections = build_plan(master, crm)

            total = sum(len(v) for v in sections.values())
            plan_date = datetime.today().strftime('%Y-%m-%d')
            wb        = build_excel(df_warmup, sections, plan_date)
            out_path  = os.path.join(tempfile.gettempdir(), f'CRM_Plan_{plan_date}.xlsx')
            wb.save(out_path)

            with open(out_path, 'rb') as f:
                excel_bytes = f.read()

            os.unlink(tmp_path)

            st.success(f'Done — {total} coils processed')
            st.download_button(
                label='⬇️  Download CRM Plan',
                data=excel_bytes,
                file_name=f'CRM_Plan_{plan_date}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
                type='primary'
            )
        except Exception as e:
            st.error(f'Error: {e}')
            import traceback
            st.code(traceback.format_exc())
