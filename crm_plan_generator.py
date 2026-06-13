import streamlit as st
import pandas as pd
import tempfile
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_BG = '1F3864'; HEADER_FG = 'FFFFFF'; LEGEND_BG = '2E75B6'
WR_BG = 'FFD700'; WR_FG = '7B0000'; UPD_BG = 'F4B183'; UPD_FG = '7B0000'
WARM_BG = 'FFF2CC'; ROW_1PASS = 'FFE699'; ROW_2PASS = 'E2EFDA'
ROW_UNKN = 'F2DCDB'; ROW_WHITE = 'FFFFFF'

COLS = [
    ('NO',8),('COIL Man #',18),('A',7),('T.T',7),('TH [mm]',9),
    ('Width',8),('T.W',8),('Int+Trim',9),('Targeted Th.',12),
    ('Steel spool',11),('PASS',7),('Previous',14),('Process',10),
    ('NEXT',10),('Passes Left',11),('Final Dest.',12),
    ('Delivery date',14),('Notes',28),('Customer',28),
]
NCOLS = len(COLS)

def thin():
    s = Side(style='thin', color='000000')
    return Border(left=s, right=s, top=s, bottom=s)

def med_border(color='7B0000'):
    s = Side(style='medium', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def cs(ws, row, col, value, bold=False, bg=None, fg='000000',
       align='center', size=10, wrap=False, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name='Calibri', bold=bold, color=fg, size=size)
    if bg:
        c.fill = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    c.border = thin()
    if num_fmt:
        c.number_format = num_fmt
    return c

def fmt_date(val):
    try:
        return pd.Timestamp(val).strftime('%Y-%m-%d')
    except:
        return str(val) if pd.notna(val) else ''

def safe_str(val):
    return str(val).strip() if pd.notna(val) else ''

def safe_float(val):
    try:
        return round(float(val), 3)
    except:
        return val

def load_data(filepath):
    master = pd.read_excel(filepath, sheet_name='Rolling Production Plan,', header=3)
    master.columns = [str(c).strip() for c in master.columns]
    crm = pd.read_excel(filepath, sheet_name='CRM ', header=2)
    crm.columns = [str(c).strip() for c in crm.columns]
    crm = crm[crm['COIL Man #'].notna() & crm['PASS'].notna()].copy()
    crm = crm[crm['PASS'].astype(str).str.startswith('P')].copy()
    return master, crm.reset_index(drop=True)

def remaining_passes(master, coil_id, cur_pass):
    journey = master[master['COIL Man #'].astype(str).str.strip() == str(coil_id).strip()].copy()
    if journey.empty:
        return 999, 'UNKNOWN'
    journey = journey.sort_values('NO')
    cm = journey[journey['PASS'].notna() & journey['PASS'].astype(str).str.startswith('P')]
    if cm.empty:
        return 999, 'UNKNOWN'
    cur = cm[cm['PASS'].astype(str).str.strip() == str(cur_pass).strip()]
    if cur.empty:
        return 999, 'UNKNOWN'
    rem = cm[cm['NO'] >= cur.iloc[0]['NO']]
    return len(rem), safe_str(rem.iloc[-1]['NEXT'])

def prioritize(master, crm):
    rows = []
    for _, r in crm.iterrows():
        pl, fd = remaining_passes(master, r['COIL Man #'], r['PASS'])
        rows.append({**r.to_dict(), '_passes_left': pl, '_final_dest': fd})
    df = pd.DataFrame(rows)
    def del_sort(v):
        try:
            return pd.Timestamp(v).timestamp()
        except:
            return 9e18
    df['_del_sort'] = df['Delivery date'].apply(del_sort)
    return df.sort_values(['_passes_left','_del_sort'], ascending=[True,True]).reset_index(drop=True)

def select_warmup(master, crm, used_coils, n=3):
    candidates = []
    for _, r in crm.iterrows():
        coil_id = safe_str(r['COIL Man #'])
        if coil_id in used_coils:
            continue
        try:
            th_f = float(r.get('TH [mm]', 0))
        except:
            continue
        cur_pass = safe_str(r['PASS'])
        if th_f <= 1.5 or cur_pass not in ('P1','P2','P3'):
            continue
        pl, fd = remaining_passes(master, coil_id, cur_pass)
        if pl < 3 or pl >= 999:
            continue
        candidates.append({**r.to_dict(), '_passes_left': pl, '_final_dest': fd, '_th_f': th_f})
    if not candidates:
        return pd.DataFrame()
    return pd.DataFrame(candidates).sort_values('_th_f', ascending=False).head(n).reset_index(drop=True)

def row_bg(passes_left, warmup=False):
    if warmup:           return WARM_BG
    if passes_left == 1: return ROW_1PASS
    if passes_left == 2: return ROW_2PASS
    if passes_left >= 999: return ROW_UNKN
    return ROW_WHITE

def write_data_row(ws, excel_row, no_val, row, warmup=False):
    bg = row_bg(row['_passes_left'], warmup)
    ws.row_dimensions[excel_row].height = 17
    def w(col, val, num_fmt=None):
        cs(ws, excel_row, col, val, bg=bg, num_fmt=num_fmt)
    w(1, no_val); w(2, safe_str(row.get('COIL Man #','')))
    w(3, row.get('A','')); w(4, row.get('T.T',''))
    w(5, safe_float(row.get('TH [mm]','')), '0.000')
    w(6, row.get('Width','')); w(7, row.get('T.W',''))
    w(8, row.get('Int + Final Trim','')); w(9, safe_float(row.get('Targeted Th.','')), '0.000')
    w(10, row.get('Steel spool','')); w(11, safe_str(row.get('PASS','')))
    w(12, safe_str(row.get('Previous','')) if pd.notna(row.get('Previous')) else '')
    w(13, safe_str(row.get('Process',''))); w(14, safe_str(row.get('NEXT','')) if pd.notna(row.get('NEXT')) else '')
    pl = row['_passes_left']
    pl_c = ws.cell(row=excel_row, column=15, value=pl if pl < 999 else 'N/A')
    pl_c.font = Font(name='Calibri', bold=(pl>=999), color=('C00000' if pl>=999 else '000000'), size=10)
    pl_c.fill = PatternFill('solid', start_color=bg)
    pl_c.alignment = Alignment(horizontal='center', vertical='center')
    pl_c.border = thin()
    w(16, safe_str(row['_final_dest'])); w(17, fmt_date(row.get('Delivery date')))
    notes_val = safe_str(row.get('Notes',''))
    if warmup and not notes_val:
        notes_val = 'Warm-up coil — take surface sample after 1st pass on new WR'
    nc = ws.cell(row=excel_row, column=18, value=notes_val)
    nc.font = Font(name='Calibri', size=9, color=('C00000' if 'ANN' in notes_val.upper() else ('7B5200' if warmup else '000000')))
    nc.fill = PatternFill('solid', start_color=bg)
    nc.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    nc.border = thin()
    cc = ws.cell(row=excel_row, column=19, value=safe_str(row.get('Customer','')))
    cc.font = Font(name='Calibri', size=9)
    cc.fill = PatternFill('solid', start_color=bg)
    cc.alignment = Alignment(horizontal='left', vertical='center')
    cc.border = thin()

def write_banner(ws, excel_row, text, bg, fg, font_size=12, bdr_color='7B0000'):
    ws.row_dimensions[excel_row].height = 28
    ws.merge_cells(start_row=excel_row, start_column=1, end_row=excel_row, end_column=NCOLS)
    c = ws.cell(row=excel_row, column=1, value=text)
    c.font = Font(name='Calibri', bold=True, color=fg, size=font_size)
    c.fill = PatternFill('solid', start_color=bg)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = med_border(bdr_color)

def build_excel(df_100, df_warmup, plan_date):
    wb = Workbook(); ws = wb.active; ws.title = 'CRM Plan'
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    tc = ws.cell(row=1, column=1, value=f'CRM Production Plan  —  {plan_date}  (100 passes)')
    tc.font = Font(name='Calibri', bold=True, color=HEADER_FG, size=14)
    tc.fill = PatternFill('solid', start_color=HEADER_BG)
    tc.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 16
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOLS)
    lc = ws.cell(row=2, column=1, value='🟡 1 pass left   🟢 2 passes left   ⬜ 3+ passes   🔴 Not in master plan   🟨 Warm-up coils')
    lc.font = Font(name='Calibri', italic=True, color=HEADER_FG, size=9)
    lc.fill = PatternFill('solid', start_color=LEGEND_BG)
    lc.alignment = Alignment(horizontal='center', vertical='center')
    HDR = 3; ws.row_dimensions[HDR].height = 22
    for ci, (name, width) in enumerate(COLS, 1):
        cs(ws, HDR, ci, name, bold=True, bg=HEADER_BG, fg=HEADER_FG)
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.freeze_panes = f'A{HDR+1}'
    DATA_START = HDR + 1
    for idx, row in df_100.iterrows():
        write_data_row(ws, DATA_START + idx, idx + 1, row)
    WR_ROW = DATA_START + 100
    write_banner(ws, WR_ROW, '⚠️   STOP  —  CHANGE WORK ROLL   |   وقّف الماكينة وغيّر الـ Work Roll   ⚠️', WR_BG, WR_FG, 13)
    WARM_START = WR_ROW + 1
    ws.row_dimensions[WARM_START].height = 20
    ws.merge_cells(start_row=WARM_START, start_column=1, end_row=WARM_START, end_column=NCOLS)
    wh = ws.cell(row=WARM_START, column=1, value='🔥  Warm-up coils — roll back to back on new WR  |  كويلات التسخين — شغّلهم متتاليين  |  خذ surface sample بعد أول باص')
    wh.font = Font(name='Calibri', bold=True, color='7B5200', size=10)
    wh.fill = PatternFill('solid', start_color='FFE0B2')
    wh.alignment = Alignment(horizontal='center', vertical='center')
    wh.border = med_border('BF8600')
    if not df_warmup.empty:
        for wi, (_, wrow) in enumerate(df_warmup.iterrows()):
            write_data_row(ws, WARM_START + 1 + wi, f'W{wi+1}', wrow, warmup=True)
        next_row = WARM_START + 1 + len(df_warmup)
    else:
        ws.row_dimensions[WARM_START+1].height = 17
        ws.merge_cells(start_row=WARM_START+1, start_column=1, end_row=WARM_START+1, end_column=NCOLS)
        fb = ws.cell(row=WARM_START+1, column=1, value='No warm-up candidates found — select manually')
        fb.font = Font(name='Calibri', italic=True, color='C00000', size=10)
        fb.fill = PatternFill('solid', start_color=WARM_BG)
        fb.alignment = Alignment(horizontal='center', vertical='center')
        fb.border = thin()
        next_row = WARM_START + 2
    write_banner(ws, next_row, '📋   حدّث الخطة من هنا  —  UPDATE THE PLAN FROM HERE   📋', UPD_BG, UPD_FG, 12, 'BF6000')
    ws.auto_filter.ref = f'A{HDR}:{get_column_letter(NCOLS)}{DATA_START+99}'
    return wb

# ── UI ────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title='CRM Plan Generator', page_icon='🏭', layout='centered')
st.title('🏭 CRM Plan Generator')

uploaded = st.file_uploader('ارفع الخطة الكاملة (Cold Rolling Schedule)', type=['xlsx'])

if uploaded:
    with st.spinner('جاري المعالجة...'):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded.read())
                tmp_path = tmp.name

            master, crm      = load_data(tmp_path)
            df_sorted        = prioritize(master, crm)
            df_100           = df_sorted.head(100).reset_index(drop=True)
            used_coils       = set(df_100['COIL Man #'].astype(str).str.strip())
            df_warmup        = select_warmup(master, crm, used_coils, n=3)
            plan_date        = datetime.today().strftime('%Y-%m-%d')
            wb               = build_excel(df_100, df_warmup, plan_date)
            out_path         = os.path.join(tempfile.gettempdir(), f'CRM_Plan_{plan_date}.xlsx')
            wb.save(out_path)

            with open(out_path, 'rb') as f:
                excel_bytes = f.read()

            os.unlink(tmp_path)

            st.success(f'✅ تم — {len(crm)} كويل في القائمة')
            st.download_button(
                label='⬇️  تحميل خطة CRM',
                data=excel_bytes,
                file_name=f'CRM_Plan_{plan_date}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
                type='primary'
            )
        except Exception as e:
            st.error(f'خطأ: {e}')
