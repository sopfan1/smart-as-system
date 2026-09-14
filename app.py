import streamlit as st
import pandas as pd
import sqlite3
import io
import os
import math
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage, ImageOps

# -------------------------------------------------------------
# [1. 기본 설정 및 경로 지정]
# -------------------------------------------------------------
st.set_page_config(layout="wide", page_title="Smart AS ERP - 현업 고정 최적화 버전")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "attached_images")
os.makedirs(IMG_DIR, exist_ok=True)

TEMPLATE_FILE = os.path.join(BASE_DIR, "양750-03 수리 Report_260715_2.xlsx")
if not os.path.exists(TEMPLATE_FILE):
    TEMPLATE_FILE = os.path.join(BASE_DIR, "양750-03 수리 Report_260715.xlsx")

conn = sqlite3.connect("smart_as.db", check_same_thread=False)

EXCEL_FIELDS = [
    'NO.', '접수일', '발생일', '제조처', '접수처', '프로젝트', '제품명', '제품 S/N', '위치', '접수횟수', 
    '유/무상', '접수내역', '확인내역', 
    '1차_육안', '1차_육안_일자',
    '1차_특성', '1차_특성_일자',
    '1차_조합', '1차_조합_일자',
    '1차_AGING', '1차_AGING_일자',
    '1차_FULL부하', '1차_FULL부하_일자',
    '재검_육안', '재검_육안_일자',
    '재검_특성', '재검_특성_일자',
    '재검_조합', '재검_조합_일자',
    '재검_AGING', '재검_AGING_일자',
    '재검_FULL부하', '재검_FULL부하_일자',
    '불량원인', '수리내역', 'F/W', 'BASE PBA S/N', 'SMPS S/N', 
    'MAC', '투입부품1', '부품1 수량', '투입부품2', '부품2 수량', '투입부품3', '부품3 수량', 
    '투입부품4', '부품4 수량', '교체 (전) S/N', '교체(후) S/N', 
    '처리결과', '담당자', '완료일자', '인계일자', '비고',
    '확인내역_사진', '수리내역_사진'
]

DATE_FIELDS = [
    '접수일', '발생일', '완료일자', '인계일자',
    '1차_육안_일자', '1차_특성_일자', '1차_조합_일자', '1차_AGING_일자', '1차_FULL부하_일자',
    '재검_육안_일자', '재검_특성_일자', '재검_조합_일자', '재검_AGING_일자', '재검_FULL부하_일자'
]

# -------------------------------------------------------------
# [2. 데이터베이스 초기화 및 초고속 로드]
# -------------------------------------------------------------
def init_db():
    cursor = conn.cursor()
    columns_def = ", ".join([f'"{col}" TEXT' for col in EXCEL_FIELDS])
    cursor.execute(f'''
        CREATE TABLE IF NOT EXISTS as_data (
            {columns_def}
        )
    ''')
    conn.commit()

init_db()

def load_db_data(order_mode="최신순 (마지막 NO.부터)"):
    try:
        df = pd.read_sql("SELECT rowid as rowid_val, * FROM as_data", conn)
    except:
        init_db()
        df = pd.read_sql("SELECT rowid as rowid_val, * FROM as_data", conn)
        
    for col in EXCEL_FIELDS:
        if col not in df.columns: df[col] = ""
    ed_df = df[['rowid_val'] + EXCEL_FIELDS].copy()
    
    for col in DATE_FIELDS:
        if col in ed_df.columns:
            ed_df[col] = ed_df[col].astype(str).str.strip().replace(['nan', 'None', 'NAT', 'NaT'], '')
            
    ed_df['temp_no_sort'] = ed_df['NO.'].apply(lambda x: safe_int_no(x) if safe_int_no(x) is not None else 999999)
    ed_df.sort_values(by='temp_no_sort', ascending=True, inplace=True)
    ed_df.drop(columns=['temp_no_sort'], inplace=True)

    if order_mode.startswith("최신순"):
        ed_df = ed_df.iloc[::-1].reset_index(drop=True)
    else:
        ed_df = ed_df.reset_index(drop=True)
    return ed_df

# -------------------------------------------------------------
# [3. 보조 유틸 함수]
# -------------------------------------------------------------
def clean_date_str(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['nan', 'none', 'nat', 'null']: return ""
    if " " in val_str and "\n" not in val_str and "~" not in val_str: val_str = val_str.split(" ")[0]
    if "T" in val_str and "\n" not in val_str and "~" not in val_str: val_str = val_str.split("T")[0]
    return val_str

def calc_aging_48h(start_date_str, multiline=False):
    if not start_date_str: return ""
    raw = str(start_date_str).strip()
    if "~" in raw: return raw
    try:
        dt = datetime.strptime(raw[:10], "%Y-%m-%d")
        dt_end = dt + timedelta(days=2)
        s_str = dt.strftime("%Y-%m-%d")
        e_str = dt_end.strftime("%Y-%m-%d")
        if multiline: return f"{s_str}\n~\n{e_str}"
        return f"{s_str} ~ {e_str}"
    except:
        return raw

def safe_int_no(val):
    if pd.isna(val): return None
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['nan', 'none', 'null', '']: return None
    if "(" in val_str: val_str = val_str.split("(")[0].strip()
    try:
        float_val = float(val_str)
        if math.isnan(float_val): return None
        return int(float_val)
    except:
        return None

def calculate_reception_counts(df_to_calc):
    sn_counter = {}
    counts = []
    seen_no_map = {}
    
    for _, row in df_to_calc.iterrows():
        sn_clean = str(row.get('제품 S/N', '')).strip()
        no_val = str(row.get('NO.', '')).strip()
        
        if not sn_clean or sn_clean.lower() in ['nan', 'none', '-', '']:
            counts.append("1")
        else:
            if no_val and no_val in seen_no_map:
                counts.append(seen_no_map[no_val])
            else:
                sn_counter[sn_clean] = sn_counter.get(sn_clean, 0) + 1
                assigned_count = str(sn_counter[sn_clean])
                if no_val: seen_no_map[no_val] = assigned_count
                counts.append(assigned_count)
                
    df_to_calc['접수횟수'] = counts
    return df_to_calc

def make_fast_ledger_excel(dataframe):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'AS관리대장'

    excel_headers = [
        'NO.', '접수일', '발생일', '제조처', '접수처', '프로젝트', '제품명', '제품 S/N', '위치', '접수횟수', 
        '유/무상', '접수내역', '확인내역', 
        '검사_육안', '검사_육안_일자', '검사_전자소자', '검사_전자소자_일자', 
        '검사_기능조합', '검사_기능조합_일자', '검사_AGING', '검사_AGING_일자', 
        '검사_FULL부하', '검사_FULL부하_일자',
        '불량원인', '수리내역', 'F/W', 'BASE PBA S/N', 'SMPS S/N', 
        'MAC', '투입부품1', '부품1 수량', '투입부품2', '부품2 수량', '투입부품3', '부품3 수량', 
        '투입부품4', '부품4 수량', '교체 (전) S/N', '교체(후) S/N', 
        '처리결과', '담당자', '완료일자', '인계일자', '비고'
    ]
    ws.append(excel_headers)
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

    for col_num in range(1, len(excel_headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = openpyxl.styles.Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    curr_row = 2
    for _, r in dataframe.iterrows():
        r_dict = r.to_dict()
        has_retest = any(str(r_dict.get(k, '')).strip() != '' for k in list(INSPECT_2ND_PAIRS.keys()) + list(INSPECT_2ND_PAIRS.values()))

        row1_vals = [
            r_dict.get('NO.', ''), r_dict.get('접수일', ''), r_dict.get('발생일', ''), r_dict.get('제조처', ''),
            r_dict.get('접수처', ''), r_dict.get('프로젝트', ''), r_dict.get('제품명', ''), r_dict.get('제품 S/N', ''),
            r_dict.get('위치', ''), r_dict.get('접수횟수', ''), r_dict.get('유/무상', ''), r_dict.get('접수내역', ''),
            r_dict.get('확인내역', ''),
            r_dict.get('1차_육안', ''), r_dict.get('1차_육안_일자', ''),
            r_dict.get('1차_특성', ''), r_dict.get('1차_특성_일자', ''),
            r_dict.get('1차_조합', ''), r_dict.get('1차_조합_일자', ''),
            r_dict.get('1차_AGING', ''), calc_aging_48h(r_dict.get('1차_AGING_일자', '')),
            r_dict.get('1차_FULL부하', ''), r_dict.get('1차_FULL부하_일자', ''),
            r_dict.get('불량원인', ''), r_dict.get('수리내역', ''), r_dict.get('F/W', ''),
            r_dict.get('BASE PBA S/N', ''), r_dict.get('SMPS S/N', ''), r_dict.get('MAC', ''),
            r_dict.get('투입부품1', ''), r_dict.get('부품1 수량', ''), r_dict.get('투입부품2', ''), r_dict.get('부품2 수량', ''),
            r_dict.get('투입부품3', ''), r_dict.get('부품3 수량', ''), r_dict.get('투입부품4', ''), r_dict.get('부품4 수량', ''),
            r_dict.get('교체 (전) S/N', ''), r_dict.get('교체(후) S/N', ''),
            r_dict.get('처리결과', ''), r_dict.get('담당자', ''), r_dict.get('완료일자', ''), r_dict.get('인계일자', ''),
            r_dict.get('비고', '')
        ]

        if has_retest:
            row2_vals = list(row1_vals)
            row2_vals[13] = r_dict.get('재검_육안', '')
            row2_vals[14] = r_dict.get('재검_육안_일자', '')
            row2_vals[15] = r_dict.get('재검_특성', '')
            row2_vals[16] = r_dict.get('재검_특성_일자', '')
            row2_vals[17] = r_dict.get('재검_조합', '')
            row2_vals[18] = r_dict.get('재검_조합_일자', '')
            row2_vals[19] = r_dict.get('재검_AGING', '')
            row2_vals[20] = calc_aging_48h(r_dict.get('재검_AGING_일자', ''))
            row2_vals[21] = r_dict.get('재검_FULL부하', '')
            row2_vals[22] = r_dict.get('재검_FULL부하_일자', '')

            for col_i, v in enumerate(row1_vals, 1):
                c = ws.cell(row=curr_row, column=col_i, value=str(v))
                c.border = thin_border
            for col_i, v in enumerate(row2_vals, 1):
                c = ws.cell(row=curr_row + 1, column=col_i, value=str(v))
                c.border = thin_border
                if 14 <= col_i <= 23: c.fill = yellow_fill

            for col_i in list(range(1, 14)) + list(range(24, len(excel_headers) + 1)):
                ws.merge_cells(start_row=curr_row, end_row=curr_row + 1, start_column=col_i, end_column=col_i)
            curr_row += 2
        else:
            for col_i, v in enumerate(row1_vals, 1):
                c = ws.cell(row=curr_row, column=col_i, value=str(v))
                c.border = thin_border
            curr_row += 1

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()

def prepare_excel_image(filepath):
    if not filepath or pd.isna(filepath): return None
    fp = str(filepath).strip()
    if not fp or fp.lower() in ['nan', 'none', '']: return None

    resolved_path = None
    if os.path.exists(fp): resolved_path = fp
    else:
        base_name = os.path.basename(fp)
        p1 = os.path.join(IMG_DIR, base_name)
        p2 = os.path.join(BASE_DIR, base_name)
        if os.path.exists(p1): resolved_path = p1
        elif os.path.exists(p2): resolved_path = p2

    if not resolved_path or not os.path.exists(resolved_path): return None
    try:
        target_w = round((4.5 / 2.54) * 96)
        target_h = round((6.0 / 2.54) * 96)
        with PILImage.open(resolved_path) as raw_img:
            img = ImageOps.exif_transpose(raw_img).resize((target_w, target_h), PILImage.Resampling.BILINEAR)
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format='JPEG', quality=80)
            buf.seek(0)
            xl_img = OpenpyxlImage(buf)
            xl_img.width = target_w
            xl_img.height = target_h
            xl_img._image_buffer = buf
            return xl_img
    except: return None

@st.cache_data(show_spinner=False)
def generate_repair_report(selected_rows_tuple, template_filename=TEMPLATE_FILE):
    selected_rows_data = [dict(r) for r in selected_rows_tuple]
    if not os.path.exists(template_filename):
        wb = openpyxl.Workbook()
        wb.active.title = "Report"
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    wb = openpyxl.load_workbook(template_filename, data_only=False)
    template_ws = wb.active

    for idx, row_data in enumerate(selected_rows_data):
        no_val = str(row_data.get('NO.', idx + 1)).replace('/', '_').strip()
        sheet_title = f"Report_NO_{no_val}"[:30]
        ws = template_ws if idx == 0 else wb.copy_worksheet(template_ws)
        ws.title = sheet_title

        def merge_lines(v1, v2):
            s1, s2 = str(v1).strip(), str(v2).strip()
            if s1 and s2: return f"{s1}\n{s2}"
            return s1 or s2 or ""

        inspect_merged = {
            '육안_결과': merge_lines(row_data.get('1차_육안', ''), row_data.get('재검_육안', '')),
            '육안_일자': merge_lines(row_data.get('1차_육안_일자', ''), row_data.get('재검_육안_일자', '')),
            '특성_결과': merge_lines(row_data.get('1차_특성', ''), row_data.get('재검_특성', '')),
            '특성_일자': merge_lines(row_data.get('1차_특성_일자', ''), row_data.get('재검_특성_일자', '')),
            '조합_결과': merge_lines(row_data.get('1차_조합', ''), row_data.get('재검_조합', '')),
            '조합_일자': merge_lines(row_data.get('1차_조합_일자', ''), row_data.get('재검_조합_일자', '')),
            'AGING_결과': merge_lines(row_data.get('1차_AGING', ''), row_data.get('재검_AGING', '')),
            'AGING_일자': merge_lines(calc_aging_48h(row_data.get('1차_AGING_일자', ''), True), calc_aging_48h(row_data.get('재검_AGING_일자', ''), True)),
            'FULL부하_결과': merge_lines(row_data.get('1차_FULL부하', ''), row_data.get('재검_FULL부하', '')),
            'FULL부하_일자': merge_lines(row_data.get('1차_FULL부하_일자', ''), row_data.get('재검_FULL부하_일자', ''))
        }

        dc, ra = str(row_data.get('불량원인', '')).strip(), str(row_data.get('수리내역', '')).strip()
        repair_desc = f"[불량원인] {dc}\n[수리내역] {ra}" if (dc and ra) else (f"[불량원인] {dc}" if dc else ra)

        for r in range(1, min(ws.max_row + 1, 50)):
            for c in range(1, min(ws.max_column + 1, 15)):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if not val: continue
                v_str = str(val).strip()

                if v_str == "접수일": ws.cell(row=r, column=c+1).value = clean_date_str(row_data.get('접수일', ''))
                elif v_str == "프로젝트": ws.cell(row=r, column=c+1).value = row_data.get('프로젝트', '')
                elif v_str == "제품명": ws.cell(row=r, column=c+1).value = row_data.get('제품명', '')
                elif "S/N" in v_str.upper(): ws.cell(row=r, column=c+1).value = row_data.get('제품 S/N', '')
                elif v_str == "접수내역": ws.cell(row=r, column=c+1).value = row_data.get('접수내역', '')
                elif "불량증상" in v_str.replace(" ", ""): ws.cell(row=r, column=10).value = row_data.get('확인내역', '')
                elif "수리내역" in v_str.replace(" ", ""): ws.cell(row=r, column=10).value = repair_desc
                elif v_str == "비고": ws.cell(row=r, column=c+1).value = row_data.get('비고', '')
                elif "확인내역_사진" in v_str.replace(" ", ""):
                    ws.cell(row=r, column=c).value = ""
                    img1 = prepare_excel_image(row_data.get('확인내역_사진'))
                    if img1: add_image_with_nudge(ws, img1, c, r, 4, 4)
                elif "수리내역_사진" in v_str.replace(" ", ""):
                    ws.cell(row=r, column=c).value = ""
                    img2 = prepare_excel_image(row_data.get('수리내역_사진'))
                    if img2: add_image_with_nudge(ws, img2, c, r, 4, 4)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def add_image_with_nudge(ws, xl_img, col, row, nx, ny):
    marker = AnchorMarker(col=col-1, colOff=nx*12700, row=row-1, rowOff=ny*12700)
    size = XDRPositiveSize2D(cx=pixels_to_EMU(xl_img.width), cy=pixels_to_EMU(xl_img.height))
    xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(xl_img)

# -------------------------------------------------------------
# [4. 세션 기반 인증 및 UI 실행부]
# -------------------------------------------------------------
st.sidebar.title("🔐 사용자 인증 및 권한")

if "auth_status" not in st.session_state:
    st.session_state["auth_status"] = False

auth_mode = st.sidebar.radio("접속 모드 선택", ["일반 사용자 (조회/REPORT 출력)", "마스터 관리자 (등록/수정 권한)"])

if auth_mode == "마스터 관리자 (등록/수정 권한)":
    input_id = st.sidebar.text_input("관리자 ID 입력", value="jwko")
    input_pw = st.sidebar.text_input("관리자 비밀번호 입력", type="password", value="qcteam12!")
    
    if input_id == "jwko" and input_pw == "qcteam12!":
        st.session_state["auth_status"] = True
        st.sidebar.success("✅ 마스터 관리자 인증 완료")
    else:
        st.session_state["auth_status"] = False
        if input_id or input_pw: st.sidebar.error("❌ 비밀번호 불일치")
else:
    st.session_state["auth_status"] = False
    st.sidebar.info("👁️ 일반 사용자 모드 (조회 전용)")

is_master = st.session_state["auth_status"]

st.title("🏢 Smart AS Management & KPI System")

if "last_order" not in st.session_state: st.session_state["last_order"] = "최신순 (마지막 NO.부터)"
current_order = st.radio("대장 표시 순서", ["최신순 (마지막 NO.부터)", "과거순 (NO. 1부터)"], horizontal=True)

if current_order != st.session_state["last_order"]:
    st.session_state["last_order"] = current_order
    st.session_state["display_df"] = load_db_data(current_order)
    if '선택' not in st.session_state["display_df"].columns: st.session_state["display_df"].insert(0, '선택', False)

if "display_df" not in st.session_state or st.session_state["display_df"] is None:
    st.session_state["display_df"] = load_db_data(current_order)
    if '선택' not in st.session_state["display_df"].columns: st.session_state["display_df"].insert(0, '선택', False)

if is_master:
    with st.expander("📥 [관리자 전용] 엑셀 파일 업로드 및 DB 동기화", expanded=False):
        uploaded_file = st.file_uploader("AS관리대장 엑셀 (.xlsx)", type=["xlsx"])
        if uploaded_file and st.button("🚀 DB 반영"):
            try:
                ex_df = pd.read_excel(uploaded_file)
                ex_df.columns = [str(c).strip() for c in ex_df.columns]
                for c in EXCEL_FIELDS:
                    if c not in ex_df.columns: ex_df[c] = ""
                ex_df['NO.'] = [str(safe_int_no(v)) if safe_int_no(v) is not None else str(i+1) for i, v in enumerate(ex_df['NO.'])]
                ex_df = calculate_reception_counts(ex_df)
                ex_df[EXCEL_FIELDS].fillna("").astype(str).to_sql("as_data", conn, if_exists="replace", index=False)
                st.session_state["display_df"] = load_db_data(current_order)
                if '선택' not in st.session_state["display_df"].columns: st.session_state["display_df"].insert(0, '선택', False)
                st.success("반영 완료")
                st.rerun()
            except Exception as e: st.error(f"오류: {e}")

try:
    tab1, tab2 = st.tabs(["📝 AS 관리대장 & 수리 REPORT", "📊 KPI 분석"])

    with tab1:
        st.markdown("#### 🔎 통합 검색 및 필터")
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1: search_query = st.text_input("통합 검색", value="")
        with sc2: sel_proj = st.selectbox("프로젝트", ["전체 프로젝트"] + sorted([str(p).strip() for p in st.session_state["display_df"]['프로젝트'].unique() if str(p).strip()]))
        with sc3: sel_cost = st.selectbox("유/무상", ["전체", "무상", "유상"])
        with sc4: sel_res = st.selectbox("처리결과", ["전체"] + sorted([str(r).strip() for r in st.session_state["display_df"]['처리결과'].unique() if str(r).strip()]))

        v_df = st.session_state["display_df"].copy()
        if search_query.strip():
            q = search_query.strip().lower()
            v_df = v_df[v_df.astype(str).apply(lambda row: row.str.lower().str.contains(q).any(), axis=1)]
        if sel_proj != "전체 프로젝트": v_df = v_df[v_df['프로젝트'].astype(str).str.strip() == sel_proj]
        if sel_cost != "전체": v_df = v_df[v_df['유/무상'].astype(str).str.strip().str.contains(sel_cost)]
        if sel_res != "전체": v_df = v_df[v_df['처리결과'].astype(str).str.strip() == sel_res]

        pass_fail_options = ["", "PASS", "FAIL"]
        column_config = {
            "선택": st.column_config.CheckboxColumn("선택", default=False),
            "NO.": st.column_config.TextColumn("NO.", disabled=not is_master),
            "접수횟수": st.column_config.TextColumn("접수횟수", disabled=True),
            "1차_육안": st.column_config.SelectboxColumn("1차_육안", options=pass_fail_options, disabled=not is_master),
            "1차_특성": st.column_config.SelectboxColumn("1차_특성", options=pass_fail_options, disabled=not is_master),
            "1차_조합": st.column_config.SelectboxColumn("1차_조합", options=pass_fail_options, disabled=not is_master),
            "1차_AGING": st.column_config.SelectboxColumn("1차_AGING", options=pass_fail_options, disabled=not is_master),
            "1차_FULL부하": st.column_config.SelectboxColumn("1차_FULL부하", options=pass_fail_options, disabled=not is_master),
            "재검_육안": st.column_config.SelectboxColumn("재검_육안", options=pass_fail_options, disabled=not is_master),
            "재검_특성": st.column_config.SelectboxColumn("재검_특성", options=pass_fail_options, disabled=not is_master),
            "재검_조합": st.column_config.SelectboxColumn("재검_조합", options=pass_fail_options, disabled=not is_master),
            "재검_AGING": st.column_config.SelectboxColumn("재검_AGING", options=pass_fail_options, disabled=not is_master),
            "재검_FULL부하": st.column_config.SelectboxColumn("재검_FULL부하", options=pass_fail_options, disabled=not is_master),
        }

        edited_df = st.data_editor(v_df[['선택'] + EXCEL_FIELDS], column_config=column_config, hide_index=True, height=750, num_rows="dynamic" if is_master else "fixed", key="main_editor_v10")

        for idx in edited_df.index:
            st.session_state["display_df"].loc[idx, ['선택'] + EXCEL_FIELDS] = edited_df.loc[idx, ['선택'] + EXCEL_FIELDS].values

        b1, b2, b3 = st.columns([3, 3, 4])
        
        with b1:
            save_btn = st.button("💾 표에서 수정한 내용 DB에 영구 저장", type="primary" if is_master else "secondary", disabled=not is_master)
            if save_btn and is_master:
                try:
                    raw_df = st.session_state["display_df"][EXCEL_FIELDS].copy()
                    if current_order.startswith("최신순"): raw_df = raw_df.iloc[::-1].reset_index(drop=True)
                    
                    valid_rows = []
                    today_str = datetime.now().strftime('%Y-%m-%d')
                    for _, row in raw_df.iterrows():
                        r_dict = row.to_dict()
                        if any(str(r_dict.get(c, '')).strip() for c in ['접수일', '프로젝트', '제품명', '제품 S/N', '접수내역', '확인내역']):
                            for tk, dk in list(INSPECT_1ST_PAIRS.items()) + list(INSPECT_2ND_PAIRS.items()):
                                if str(r_dict.get(tk, '')).strip().upper() in ['PASS', 'FAIL'] and not str(r_dict.get(dk, '')).strip():
                                    r_dict[dk] = today_str
                            valid_rows.append(r_dict)
                    
                    if valid_rows:
                        p_df = pd.DataFrame(valid_rows, columns=EXCEL_FIELDS)
                        p_df = calculate_reception_counts(p_df)
                        p_df[EXCEL_FIELDS].fillna("").astype(str).to_sql("as_data", conn, if_exists="replace", index=False)
                        st.session_state["display_df"] = load_db_data(current_order)
                        if '선택' not in st.session_state["display_df"].columns: st.session_state["display_df"].insert(0, '선택', False)
                        st.toast("✅ DB 영구 저장 완료", icon="💾")
                        st.rerun()
                except Exception as e: st.error(f"저장 오류: {e}")

        with b2:
            del_btn = st.button("🗑️ 선택한 행 DB에서 완전 삭제", type="primary" if is_master else "secondary", disabled=not is_master)
            if del_btn and is_master:
                targets = edited_df[edited_df['선택'] == True]
                if not targets.empty:
                    rem_df = st.session_state["display_df"].drop(index=targets.index.tolist())[EXCEL_FIELDS]
                    rem_df.fillna("").astype(str).to_sql("as_data", conn, if_exists="replace", index=False)
                    st.session_state["display_df"] = load_db_data(current_order)
                    if '선택' not in st.session_state["display_df"].columns: st.session_state["display_df"].insert(0, '선택', False)
                    st.toast("🗑️ 삭제 완료", icon="✅")
                    st.rerun()

        with b3:
            ledger_bytes = make_fast_ledger_excel(load_db_data(current_order)[EXCEL_FIELDS])
            st.download_button(label="📥 AS관리대장 엑셀 다운로드", data=ledger_bytes, file_name=f"AS관리대장_{datetime.now().strftime('%y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        sel_rows = edited_df[edited_df['선택'] == True]
        if not sel_rows.empty:
            st.markdown("---")
            st.subheader(f"📑 수리 REPORT 발행 (총 {len(sel_rows)}건 선택됨)")
            try:
                r_tuple = tuple(tuple(sorted(r.items())) for r in sel_rows.to_dict(orient="records"))
                report_bytes = generate_repair_report(r_tuple, TEMPLATE_FILE)
                st.download_button(label=f"📥 선택한 {len(sel_rows)}건 수리 REPORT 엑셀 다운로드", data=report_bytes, file_name=f"수리Report_{datetime.now().strftime('%y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            except Exception as e: st.error(f"REPORT 생성 오류: {e}")

    with tab2:
        st.subheader("📊 2026년 품질 및 AS 운영 KPI 대시보드")
        kpi_df = load_db_data(current_order)
        st.metric("총 접수 건수", f"{len(kpi_df):,} 건")

except Exception as e: st.error(f"시스템 오류: {e}")