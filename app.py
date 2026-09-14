import streamlit as st
import pandas as pd
import sqlite3
import io
import os
import math
from datetime import datetime, timedelta
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage, ImageOps

# -------------------------------------------------------------
# [1. 기본 설정 및 경로 지정]
# -------------------------------------------------------------
st.set_page_config(layout="wide", page_title="Smart AS ERP - 최적화 버전")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "attached_images")
os.makedirs(IMG_DIR, exist_ok=True)

TEMPLATE_FILE = os.path.join(BASE_DIR, "양750-03 수리 Report_260715_2.xlsx")
if not os.path.exists(TEMPLATE_FILE):
    TEMPLATE_FILE = os.path.join(BASE_DIR, "양750-03 수리 Report_260715.xlsx")


# ✅ 개선①: 연결을 캐시로 만들어 재실행마다 새 커넥션이 생기지 않도록 함
@st.cache_resource
def get_connection():
    return sqlite3.connect("smart_as.db", check_same_thread=False)

conn = get_connection()

INSPECT_1ST_PAIRS = {
    '1차_육안': '1차_육안_일자',
    '1차_특성': '1차_특성_일자',
    '1차_조합': '1차_조합_일자',
    '1차_AGING': '1차_AGING_일자',
    '1차_FULL부하': '1차_FULL부하_일자'
}

INSPECT_2ND_PAIRS = {
    '재검_육안': '재검_육안_일자',
    '재검_특성': '재검_특성_일자',
    '재검_조합': '재검_조합_일자',
    '재검_AGING': '재검_AGING_일자',
    '재검_FULL부하': '재검_FULL부하_일자'
}

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

# 검사 항목 라벨(D열) → DB 필드 매핑
# (라벨 매칭키, 1차_결과, 1차_일자, 재검_결과, 재검_일자, AGING여부)
# ⚠ 매칭키는 다른 항목명에 포함되지 않도록 고유하게 지정.
#    특히 'AGING 기능시험'에 '기능'이 들어있어, 조합시험 키는 '조합'으로 둔다.
INSPECT_ROW_MAP = [
    ('육안', '1차_육안', '1차_육안_일자', '재검_육안', '재검_육안_일자', False),
    ('전자소자', '1차_특성', '1차_특성_일자', '재검_특성', '재검_특성_일자', False),
    ('조합', '1차_조합', '1차_조합_일자', '재검_조합', '재검_조합_일자', False),
    ('AGING', '1차_AGING', '1차_AGING_일자', '재검_AGING', '재검_AGING_일자', True),
    ('FULL부하', '1차_FULL부하', '1차_FULL부하_일자', '재검_FULL부하', '재검_FULL부하_일자', False),
]

DATE_FIELDS = [
    '접수일', '발생일', '완료일자', '인계일자',
    '1차_육안_일자', '1차_특성_일자', '1차_조합_일자', '1차_AGING_일자', '1차_FULL부하_일자',
    '재검_육안_일자', '재검_특성_일자', '재검_조합_일자', '재검_AGING_일자', '재검_FULL부하_일자'
]

# 검색용 문자열에 포함할 컬럼(사진 경로는 제외)
SEARCH_COLS = tuple(c for c in EXCEL_FIELDS if c not in ('확인내역_사진', '수리내역_사진'))

# ✅ 개선②: 스타일 상수 정의 (원본 코드의 yellow_fill NameError 버그 수정)
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
HEADER_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
)
BOLD_FONT = Font(bold=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")


# -------------------------------------------------------------
# [2. 데이터베이스 초기화 및 초고속 로드]
# -------------------------------------------------------------
def init_db():
    cursor = conn.cursor()
    columns_def = ", ".join([f'"{col}" TEXT' for col in EXCEL_FIELDS])
    cursor.execute(f'CREATE TABLE IF NOT EXISTS as_data ({columns_def})')
    conn.commit()

init_db()


def safe_int_no(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ('nan', 'none', 'null', ''):
        return None
    if "(" in val_str:
        val_str = val_str.split("(")[0].strip()
    try:
        float_val = float(val_str)
        if math.isnan(float_val):
            return None
        return int(float_val)
    except (ValueError, TypeError):
        return None


@st.cache_data(show_spinner=False)
def load_db_data(order_mode="최신순 (마지막 NO.부터)"):
    try:
        df = pd.read_sql("SELECT rowid as rowid_val, * FROM as_data", conn)
    except Exception:
        init_db()
        df = pd.read_sql("SELECT rowid as rowid_val, * FROM as_data", conn)

    # ✅ 개선③: 누락 컬럼 일괄 처리
    for col in [c for c in EXCEL_FIELDS if c not in df.columns]:
        df[col] = ""
    ed_df = df[['rowid_val'] + EXCEL_FIELDS].copy()

    # ✅ 개선④: 날짜 컬럼 정리를 벡터 연산으로
    for col in DATE_FIELDS:
        if col in ed_df.columns:
            ed_df[col] = (
                ed_df[col].astype(str).str.strip()
                .replace(['nan', 'None', 'NAT', 'NaT'], '')
            )

    # 정렬용 정수 키를 벡터로 생성
    ed_df['temp_no_sort'] = ed_df['NO.'].map(
        lambda x: safe_int_no(x) if safe_int_no(x) is not None else 999999
    )
    ed_df.sort_values(by='temp_no_sort', ascending=True, inplace=True)
    ed_df.drop(columns=['temp_no_sort'], inplace=True)

    if order_mode.startswith("최신순"):
        ed_df = ed_df.iloc[::-1].reset_index(drop=True)
    else:
        ed_df = ed_df.reset_index(drop=True)
    return ed_df


def clear_data_cache():
    st.cache_data.clear()


# -------------------------------------------------------------
# [3. 보조 유틸 함수]
# -------------------------------------------------------------
def clean_date_str(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ('nan', 'none', 'nat', 'null'):
        return ""
    if " " in val_str and "\n" not in val_str and "~" not in val_str:
        val_str = val_str.split(" ")[0]
    if "T" in val_str and "\n" not in val_str and "~" not in val_str:
        val_str = val_str.split("T")[0]
    return val_str


def calc_aging_48h(start_date_str, multiline=False):
    if not start_date_str:
        return ""
    raw = str(start_date_str).strip()
    if not raw or raw.lower() in ('nan', 'none', 'nat', '-'):
        return ""
    # 이미 '~'가 포함돼 있어도 맨 앞 날짜(시작일)만 뽑아 항상 +2일로 재계산
    head = raw.split("~")[0].strip()
    head = head.replace("\n", " ").strip()[:10]
    try:
        dt = datetime.strptime(head, "%Y-%m-%d")
        dt_end = dt + timedelta(days=2)
        s_str = dt.strftime("%Y-%m-%d")
        e_str = dt_end.strftime("%Y-%m-%d")
        if multiline:
            return f"{s_str}\n~\n{e_str}"
        return f"{s_str} ~ {e_str}"
    except Exception:
        return raw


def calculate_reception_counts(df_to_calc):
    # ✅ 개선⑤: iterrows 대신 컬럼 배열을 직접 순회
    sn_arr = df_to_calc['제품 S/N'].astype(str).str.strip().tolist()
    no_arr = df_to_calc['NO.'].astype(str).str.strip().tolist()

    sn_counter = {}
    seen_no_map = {}
    counts = []
    invalid = {'nan', 'none', '-', ''}

    for sn_clean, no_val in zip(sn_arr, no_arr):
        if not sn_clean or sn_clean.lower() in invalid:
            counts.append("1")
        elif no_val and no_val in seen_no_map:
            counts.append(seen_no_map[no_val])
        else:
            sn_counter[sn_clean] = sn_counter.get(sn_clean, 0) + 1
            assigned = str(sn_counter[sn_clean])
            if no_val:
                seen_no_map[no_val] = assigned
            counts.append(assigned)

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

    for col_num in range(1, len(excel_headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = HEADER_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN

    retest_keys = list(INSPECT_2ND_PAIRS.keys()) + list(INSPECT_2ND_PAIRS.values())
    curr_row = 2

    # ✅ 개선⑥: iterrows → records 변환
    for r_dict in dataframe.to_dict(orient="records"):
        has_retest = any(str(r_dict.get(k, '')).strip() for k in retest_keys)

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
                c.border = THIN_BORDER
            for col_i, v in enumerate(row2_vals, 1):
                c = ws.cell(row=curr_row + 1, column=col_i, value=str(v))
                c.border = THIN_BORDER
                if 14 <= col_i <= 23:
                    c.fill = YELLOW_FILL

            for col_i in list(range(1, 14)) + list(range(24, len(excel_headers) + 1)):
                ws.merge_cells(start_row=curr_row, end_row=curr_row + 1,
                               start_column=col_i, end_column=col_i)
            curr_row += 2
        else:
            for col_i, v in enumerate(row1_vals, 1):
                c = ws.cell(row=curr_row, column=col_i, value=str(v))
                c.border = THIN_BORDER
            curr_row += 1

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()


# ✅ 개선⑦: 이미지 전처리 결과를 경로+수정시각 기준으로 캐시
@st.cache_data(show_spinner=False)
def _prepare_image_bytes(resolved_path, mtime):
    target_w = round((4.5 / 2.54) * 96)
    target_h = round((6.0 / 2.54) * 96)
    with PILImage.open(resolved_path) as raw_img:
        img = ImageOps.exif_transpose(raw_img).resize(
            (target_w, target_h), PILImage.Resampling.BILINEAR
        )
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format='JPEG', quality=80)
        return buf.getvalue(), target_w, target_h


def prepare_excel_image(filepath):
    if not filepath or pd.isna(filepath):
        return None
    fp = str(filepath).strip()
    if not fp or fp.lower() in ('nan', 'none', ''):
        return None

    resolved_path = None
    if os.path.exists(fp):
        resolved_path = fp
    else:
        base_name = os.path.basename(fp)
        p1 = os.path.join(IMG_DIR, base_name)
        p2 = os.path.join(BASE_DIR, base_name)
        if os.path.exists(p1):
            resolved_path = p1
        elif os.path.exists(p2):
            resolved_path = p2

    if not resolved_path or not os.path.exists(resolved_path):
        return None
    try:
        mtime = os.path.getmtime(resolved_path)
        img_bytes, w, h = _prepare_image_bytes(resolved_path, mtime)
        buf = io.BytesIO(img_bytes)
        xl_img = OpenpyxlImage(buf)
        xl_img.width = w
        xl_img.height = h
        xl_img._image_buffer = buf
        return xl_img
    except Exception:
        return None


def add_image_with_nudge(ws, xl_img, col, row, nx, ny):
    marker = AnchorMarker(col=col - 1, colOff=nx * 12700, row=row - 1, rowOff=ny * 12700)
    size = XDRPositiveSize2D(cx=pixels_to_EMU(xl_img.width), cy=pixels_to_EMU(xl_img.height))
    xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(xl_img)


# ✅ 개선⑭: 병합 셀 안전 쓰기 (MergedCell read-only 오류 방지)
def _safe_set(ws, row, col, value):
    """(row, col)이 병합 셀의 비앵커 위치면 병합범위 좌상단 앵커에 대신 쓴다."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                ws.cell(row=rng.min_row, column=rng.min_col).value = value
                return
        return  # 병합 정보를 못 찾으면 조용히 스킵
    cell.value = value


# ✅ 개선⑮: 라벨이 병합 셀이면 그 병합범위가 끝나는 '다음 열'을 값 칸으로 계산.
#           (라벨 바로 오른쪽 c+1은 병합 내부라, 값이 라벨을 덮어쓰던 버그를 해결)
def _value_col_after_label(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.max_col + 1
    return col + 1


def _merge_lines(v1, v2):
    s1, s2 = str(v1).strip(), str(v2).strip()
    if s1 and s2:
        return f"{s1}\n{s2}"
    return s1 or s2 or ""


# ✅ 개선⑯: 출하검사 표의 한 행(항목)을 채운다.
#   - 검사일자(M열 계열)와 판정(P열 계열)을 DB에서 가져와 기록
#   - 재검(2차) 값이 있으면 1차값과 줄바꿈으로 함께 표시 + 노란색 강조
#   - AGING 항목은 시작일 +2일 범위로 표시
def _fill_inspect_row(ws, row, col, v_nospace, row_data):
    for key, r1k, d1k, r2k, d2k, is_aging in INSPECT_ROW_MAP:
        if key not in v_nospace:
            continue

        r1 = str(row_data.get(r1k, '')).strip()
        r2 = str(row_data.get(r2k, '')).strip()
        has_retest = bool(r2)

        if is_aging:
            # AGING: 판정이 있으면 항상 +2일 범위로. 일자가 비면 접수일을 기준으로 사용.
            base1 = str(row_data.get(d1k, '')).strip() or str(row_data.get('접수일', '')).strip()
            base2 = str(row_data.get(d2k, '')).strip()
            d1 = calc_aging_48h(base1, multiline=True) if r1 else ""
            d2 = calc_aging_48h(base2, multiline=True) if r2 else ""
        else:
            d1 = str(row_data.get(d1k, '')).strip()
            d2 = str(row_data.get(d2k, '')).strip()

        verdict = _merge_lines(r1, r2)
        datestr = _merge_lines(d1, d2)

        # 항목 라벨(col) 기준으로 검사일자 열과 판정 열을 병합구조에서 계산
        # 라벨 병합 끝 다음 = 규격열 시작 → 규격 병합 끝 다음 = 검사일자 → 그 다음 = 판정
        spec_col = _value_col_after_label(ws, row, col)          # 규격(기준) 열
        date_col = _col_after_merge(ws, row, spec_col)           # 검사일자 열
        judge_col = _col_after_merge(ws, row, date_col)          # 판정 열

        _safe_set(ws, row, date_col, datestr if datestr else "-")
        _safe_set(ws, row, judge_col, verdict if verdict else "-")

        # 재검이 있으면 해당 항목행 전체(검사일자·판정 병합영역)를 노란색으로
        if has_retest:
            _highlight_row_cells(ws, row, [date_col, judge_col])
        return


def _col_after_merge(ws, row, col):
    """(row,col)이 병합이면 병합 끝+1, 아니면 col+1"""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.max_col + 1
    return col + 1


def _highlight_row_cells(ws, row, cols):
    """지정 열들이 속한 병합영역의 앵커 셀을 노란색으로 칠한다."""
    for col in cols:
        anchor_r, anchor_c = row, col
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                anchor_r, anchor_c = rng.min_row, rng.min_col
                break
        ws.cell(row=anchor_r, column=anchor_c).fill = YELLOW_FILL


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

    def merge_lines(v1, v2):
        s1, s2 = str(v1).strip(), str(v2).strip()
        if s1 and s2:
            return f"{s1}\n{s2}"
        return s1 or s2 or ""

    for idx, row_data in enumerate(selected_rows_data):
        no_val = str(row_data.get('NO.', idx + 1)).replace('/', '_').strip()
        sheet_title = f"Report_NO_{no_val}"[:30]
        ws = template_ws if idx == 0 else wb.copy_worksheet(template_ws)
        ws.title = sheet_title

        dc = str(row_data.get('불량원인', '')).strip()
        ra = str(row_data.get('수리내역', '')).strip()
        repair_desc = f"[불량원인] {dc}\n[수리내역] {ra}" if (dc and ra) else (f"[불량원인] {dc}" if dc else ra)

        # ✅ 개선⑧: 셀 순회를 iter_rows로, max 범위를 한 번만 계산
        max_r = min(ws.max_row, 50)
        max_c = min(ws.max_column, 15)
        # 순회 중 셀을 수정하므로, 매칭 위치를 먼저 수집한 뒤 일괄 처리
        matches = []
        for row_cells in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
            for cell in row_cells:
                val = cell.value
                if not val:
                    continue
                matches.append((cell.row, cell.column, str(val).strip()))

        for r, c, v_str in matches:
            v_nospace = v_str.replace(" ", "")
            # 라벨이 병합 셀이면 병합 끝 다음 열이 값 칸
            vc = _value_col_after_label(ws, r, c)

            if v_str == "접수일":
                _safe_set(ws, r, vc, clean_date_str(row_data.get('접수일', '')))
            elif v_str == "프로젝트":
                _safe_set(ws, r, vc, row_data.get('프로젝트', ''))
            elif v_str == "제품명":
                _safe_set(ws, r, vc, row_data.get('제품명', ''))
            elif "S/N" in v_str.upper():
                _safe_set(ws, r, vc, row_data.get('제품 S/N', ''))
            elif v_str == "접수내역":
                _safe_set(ws, r, vc, row_data.get('접수내역', ''))
            elif "불량증상" in v_nospace:
                # 불량증상/수리내역 값은 10번 열(J)에 기록
                _safe_set(ws, r, 10, row_data.get('확인내역', ''))
            elif "수리내역" in v_nospace and "사진" not in v_nospace:
                _safe_set(ws, r, 10, repair_desc)
            elif v_str == "비고":
                _safe_set(ws, r, vc, row_data.get('비고', ''))
            elif "확인내역_사진" in v_nospace:
                _safe_set(ws, r, c, "")
                img1 = prepare_excel_image(row_data.get('확인내역_사진'))
                if img1:
                    add_image_with_nudge(ws, img1, c, r, 4, 4)
            elif "수리내역_사진" in v_nospace:
                _safe_set(ws, r, c, "")
                img2 = prepare_excel_image(row_data.get('수리내역_사진'))
                if img2:
                    add_image_with_nudge(ws, img2, c, r, 4, 4)
            else:
                # ✅ 개선⑯: 출하검사 표 채우기 (검사일자·판정, 재검 병합, AGING +2일)
                _fill_inspect_row(ws, r, c, v_nospace, row_data)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ✅ 개선⑬: 검색용 합침 문자열을 연도·건수 단위로 캐시
@st.cache_data(show_spinner=False)
def build_haystack(cache_key, series_dict):
    # series_dict: {컬럼명: 값리스트} 형태로 넘겨 캐시 키를 안정적으로 만듦
    frame = pd.DataFrame(series_dict)
    return frame.astype(str).agg(' '.join, axis=1).str.lower()


# -------------------------------------------------------------
# [4. 세션 기반 인증 및 조회 UI 실행부]
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
        if input_id or input_pw:
            st.sidebar.error("❌ 비밀번호 불일치")
else:
    st.session_state["auth_status"] = False
    st.sidebar.info("👁️ 일반 사용자 모드 (조회 전용)")

is_master = st.session_state["auth_status"]

st.title("🏢 Smart AS Management & KPI System")

if "last_order" not in st.session_state:
    st.session_state["last_order"] = "최신순 (마지막 NO.부터)"
current_order = st.radio("대장 표시 순서", ["최신순 (마지막 NO.부터)", "과거순 (NO. 1부터)"], horizontal=True)


def ensure_select_col():
    if '선택' not in st.session_state["display_df"].columns:
        st.session_state["display_df"].insert(0, '선택', False)


if current_order != st.session_state["last_order"]:
    st.session_state["last_order"] = current_order
    clear_data_cache()
    st.session_state["display_df"] = load_db_data(current_order)
    ensure_select_col()

if "display_df" not in st.session_state or st.session_state["display_df"] is None:
    st.session_state["display_df"] = load_db_data(current_order)
    ensure_select_col()

if is_master:
    with st.expander("📥 [관리자 전용] 엑셀 파일 업로드 및 DB 동기화", expanded=False):
        uploaded_file = st.file_uploader("AS관리대장 엑셀 (.xlsx)", type=["xlsx"])
        if uploaded_file and st.button("🚀 DB 반영"):
            try:
                ex_df = pd.read_excel(uploaded_file)
                ex_df.columns = [str(c).strip() for c in ex_df.columns]
                for c in EXCEL_FIELDS:
                    if c not in ex_df.columns:
                        ex_df[c] = ""
                ex_df['NO.'] = [
                    str(safe_int_no(v)) if safe_int_no(v) is not None else str(i + 1)
                    for i, v in enumerate(ex_df['NO.'])
                ]
                ex_df = calculate_reception_counts(ex_df)
                ex_df[EXCEL_FIELDS].fillna("").astype(str).to_sql(
                    "as_data", conn, if_exists="replace", index=False
                )
                clear_data_cache()
                st.session_state["display_df"] = load_db_data(current_order)
                ensure_select_col()
                st.success("반영 완료")
                st.rerun()
            except Exception as e:
                st.error(f"오류: {e}")

try:
    tab1, tab2 = st.tabs(["📝 AS 관리대장 & 수리 REPORT", "📊 KPI 분석"])

    with tab1:
        full_df = st.session_state["display_df"].copy()
        full_df['접수년도'] = full_df['접수일'].astype(str).str[:4]
        available_years = sorted(
            [y for y in full_df['접수년도'].unique() if y.isdigit() and len(y) == 4],
            reverse=True
        )
        if not available_years:
            available_years = [str(datetime.now().year)]

        st.markdown("#### 📅 조회 연도 선택 및 필터")
        y_col1, sc1, sc2, sc3, sc4 = st.columns([1.5, 2.2, 2.2, 2.2, 2.2])
        with y_col1:
            sel_year = st.selectbox("조회 연도", ["전체 보기"] + available_years, index=0)

        if sel_year == "전체 보기":
            year_filtered_df = full_df
        else:
            year_filtered_df = full_df[full_df['접수년도'] == sel_year]

        with sc1:
            search_query = st.text_input("통합 검색", value="")
        with sc2:
            sel_proj = st.selectbox(
                "프로젝트",
                ["전체 프로젝트"] + sorted(
                    [str(p).strip() for p in year_filtered_df['프로젝트'].unique() if str(p).strip()]
                )
            )
        with sc3:
            sel_cost = st.selectbox("유/무상", ["전체", "무상", "유상"])
        with sc4:
            sel_res = st.selectbox(
                "처리결과",
                ["전체"] + sorted(
                    [str(r).strip() for r in year_filtered_df['처리결과'].unique() if str(r).strip()]
                )
            )

        v_df = year_filtered_df

        # ✅ 개선⑨+⑬: 검색용 합침 문자열은 연도·건수 단위로 캐시하고,
        #             타이핑 중에는 str.contains만 재실행되도록 함
        if search_query.strip():
            q = search_query.strip().lower()
            cache_key = f"{sel_year}|{len(year_filtered_df)}"
            series_dict = {c: year_filtered_df[c].astype(str).tolist() for c in SEARCH_COLS}
            haystack = build_haystack(cache_key, series_dict)
            haystack.index = year_filtered_df.index
            v_df = year_filtered_df[haystack.str.contains(q, regex=False, na=False)]

        if sel_proj != "전체 프로젝트":
            v_df = v_df[v_df['프로젝트'].astype(str).str.strip() == sel_proj]
        if sel_cost != "전체":
            v_df = v_df[v_df['유/무상'].astype(str).str.strip().str.contains(sel_cost, na=False)]
        if sel_res != "전체":
            v_df = v_df[v_df['처리결과'].astype(str).str.strip() == sel_res]

        st.caption(
            f"📌 **{sel_year}** 데이터 총 {len(year_filtered_df):,}건 중 "
            f"**검색/필터된 항목: {len(v_df):,}건** 표시 중"
        )

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

        edited_df = st.data_editor(
            v_df[['선택'] + EXCEL_FIELDS],
            column_config=column_config,
            hide_index=True,
            height=750,
            num_rows="dynamic" if is_master else "fixed",
            key="main_editor_v14"
        )

        # ✅ 개선⑰: 편집 내용 동기화는 '실제 변경이 있을 때만' 수행.
        #   매 실행마다 무조건 대입하면 data_editor가 재계산되며 스크롤이 위로 튕기고,
        #   리포트/다운로드 버튼 클릭 시 NO.가 엉뚱하게 이동하는 현상이 생김.
        edited_aligned = edited_df.copy()
        edited_aligned.index = v_df.index[:len(edited_df)]
        cols_to_sync = ['선택'] + EXCEL_FIELDS
        try:
            current_block = st.session_state["display_df"].loc[edited_aligned.index, cols_to_sync]
            if not current_block.reset_index(drop=True).equals(
                    edited_aligned[cols_to_sync].reset_index(drop=True)):
                st.session_state["display_df"].loc[edited_aligned.index, cols_to_sync] = \
                    edited_aligned[cols_to_sync]
        except Exception:
            st.session_state["display_df"].loc[edited_aligned.index, cols_to_sync] = \
                edited_aligned[cols_to_sync]

        b1, b2, b3 = st.columns([3, 3, 4])

        with b1:
            save_btn = st.button(
                "💾 표에서 수정한 내용 DB에 영구 저장",
                type="primary" if is_master else "secondary",
                disabled=not is_master
            )
            if save_btn and is_master:
                try:
                    raw_df = st.session_state["display_df"][EXCEL_FIELDS].copy()
                    if current_order.startswith("최신순"):
                        raw_df = raw_df.iloc[::-1].reset_index(drop=True)

                    today_str = datetime.now().strftime('%Y-%m-%d')
                    pair_items = list(INSPECT_1ST_PAIRS.items()) + list(INSPECT_2ND_PAIRS.items())
                    key_cols = ['접수일', '프로젝트', '제품명', '제품 S/N', '접수내역', '확인내역']

                    valid_rows = []
                    for r_dict in raw_df.to_dict(orient="records"):
                        if any(str(r_dict.get(c, '')).strip() for c in key_cols):
                            for tk, dk in pair_items:
                                if (str(r_dict.get(tk, '')).strip().upper() in ('PASS', 'FAIL')
                                        and not str(r_dict.get(dk, '')).strip()):
                                    r_dict[dk] = today_str
                            valid_rows.append(r_dict)

                    if valid_rows:
                        p_df = pd.DataFrame(valid_rows, columns=EXCEL_FIELDS)
                        p_df = calculate_reception_counts(p_df)
                        p_df[EXCEL_FIELDS].fillna("").astype(str).to_sql(
                            "as_data", conn, if_exists="replace", index=False
                        )
                        clear_data_cache()
                        st.session_state["display_df"] = load_db_data(current_order)
                        ensure_select_col()
                        st.toast("✅ DB 영구 저장 완료", icon="💾")
                        st.rerun()
                except Exception as e:
                    st.error(f"저장 오류: {e}")

        with b2:
            del_btn = st.button(
                "🗑️ 선택한 행 DB에서 완전 삭제",
                type="primary" if is_master else "secondary",
                disabled=not is_master
            )
            if del_btn and is_master:
                targets = edited_df[edited_df['선택'] == True]
                if not targets.empty:
                    target_orig_indices = v_df.index[:len(edited_df)][edited_df['선택'].values == True]
                    rem_df = st.session_state["display_df"].drop(index=target_orig_indices)[EXCEL_FIELDS]
                    rem_df.fillna("").astype(str).to_sql("as_data", conn, if_exists="replace", index=False)
                    clear_data_cache()
                    st.session_state["display_df"] = load_db_data(current_order)
                    ensure_select_col()
                    st.toast("🗑️ 삭제 완료", icon="✅")
                    st.rerun()

        with b3:
            # ✅ 개선⑪: 대장 엑셀은 버튼을 누를 때만 생성
            if st.button("📄 AS관리대장 엑셀 생성"):
                st.session_state["_ledger_bytes"] = make_fast_ledger_excel(
                    load_db_data(current_order)[EXCEL_FIELDS]
                )
            if st.session_state.get("_ledger_bytes"):
                st.download_button(
                    label="📥 AS관리대장 엑셀 다운로드",
                    data=st.session_state["_ledger_bytes"],
                    file_name=f"AS관리대장_{datetime.now().strftime('%y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        sel_rows = edited_df[edited_df['선택'] == True]
        if not sel_rows.empty:
            st.markdown("---")
            sel_nos = [str(x).strip() for x in sel_rows['NO.'].tolist()]
            st.subheader(f"📑 수리 REPORT 발행 (선택: NO. {', '.join(sel_nos)} / 총 {len(sel_rows)}건)")

            # ✅ 개선⑱: 리포트 생성을 form으로 감싸 버튼 클릭 시 상단 표가
            #   재계산·재스크롤되지 않도록 함. 선택 데이터는 즉시 스냅샷으로 확보.
            with st.form("report_form", clear_on_submit=False):
                submitted = st.form_submit_button(
                    f"🛠️ 선택한 {len(sel_rows)}건 REPORT 생성", type="primary"
                )
                if submitted:
                    try:
                        snapshot = sel_rows.to_dict(orient="records")
                        r_tuple = tuple(tuple(sorted(r.items())) for r in snapshot)
                        st.session_state["_report_bytes"] = generate_repair_report(r_tuple, TEMPLATE_FILE)
                        st.session_state["_report_nos"] = "_".join(sel_nos)[:40]
                        st.toast(f"✅ REPORT 생성 완료 (NO. {', '.join(sel_nos)})", icon="📑")
                    except Exception as e:
                        st.error(f"REPORT 생성 오류: {e}")

            if st.session_state.get("_report_bytes"):
                nos_tag = st.session_state.get("_report_nos", "")
                st.download_button(
                    label="📥 수리 REPORT 엑셀 다운로드",
                    data=st.session_state["_report_bytes"],
                    file_name=f"수리Report_NO{nos_tag}_{datetime.now().strftime('%y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

    with tab2:
        st.subheader("📊 2026년 품질 및 AS 운영 KPI 대시보드")
        kpi_df = load_db_data(current_order)
        st.metric("총 접수 건수", f"{len(kpi_df):,} 건")

except Exception as e:
    st.error(f"시스템 오류: {e}")
