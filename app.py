# =============================================
# [스텝-바이-스텝 알고리즘]
# 1) 사용자가 인사랑 '관내출장내역' 엑셀(.xls/.xlsx) 업로드
# 2) 헤더 검증: B4~W4 좌표가 EXPECTED_HEADERS와 일치하는지 확인
#    - 불일치 시 목록을 보여주고 처리 중단
# 3) 데이터 로드: 4행을 헤더로 간주(header=3), 문자열 통일(dtype=str)
# 4) 전처리:
#    - '순번','구분' 공백/NaN 제거
#    - '결재상태'에 '결재완료' 포함 행만 유지
# 5) 파생열 계산:
#    - X(4시간 구분): '총출장시간'에서 "4시간이상/4시간미만" 추출
#    - Y(1시간 미만): '총출장시간'이 "분"만 포함될 때 "1시간미만"
#    - Z(지급액): X와 '차량' 사용여부에 따라 0/1만/2만 원 산정
# 6) 정렬: X(이상→미만→공란), Z(내림차순) 기준
# 7) '정리결과' 시트 생성:
#    - 원본열 + [X,Y,Z] 순으로 기록
#    - 헤더 표시명 교체 및 서식(배경, 정렬, 금액 서식)
# 8) '요약' 시트 생성:
#    - 열 구성: [출장자, 지급단가, 합계금액, 출장일자(일 단위 나열)]
#    - 2만/1만만 집계, 합계금액 = 지급단가 × 일수
#    - 조건부서식: 합계금액 300,000 초과 분홍, 동일 이름 중복 표기 노란색
#    - 테두리, 너비 자동 조정, 첫 화면으로 설정
# 9) 메모리로 저장하여 Streamlit 다운로드 버튼 제공
# =============================================

import os
import re
from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import CellIsRule, FormulaRule

# ---- 시간대(KST) ----
try:
    from zoneinfo import ZoneInfo
    KST = ZoneInfo("Asia/Seoul")
except ImportError:
    from pytz import timezone
    KST = timezone("Asia/Seoul")

# =========================
# 상수
# =========================
EXPECTED_HEADERS = {
    "B4": "순번", "C4": "구분", "E4": "출발일자", "F4": "도착일자",
    "G4": "총출장시간", "I4": "출장지", "K4": "차량", "L4": "출장목적",
    "M4": "소속", "O4": "출장자", "P4": "지출", "R4": "여비",
    "S4": "지출일자", "T4": "결재상태", "U4": "여비등급",
    "V4": "비고", "W4": "업무대행",
}
APP_TITLE = "관내 출장내역 정리"
MANUAL_FILE = "인사랑 관내출장 내역 추출.pdf"

# =========================
# 공통 유틸
# =========================
def kst_now() -> datetime:
    return datetime.now(KST)

def kst_timestamp() -> str:
    return kst_now().strftime("%y%m%d_%H%M")

def read_headers_for_check(file_like_object) -> dict:
    """엑셀의 지정 셀(B4~W4)을 읽어 헤더 텍스트를 반환"""
    wb = load_workbook(file_like_object, data_only=True, read_only=True)
    ws = wb.active
    vals = {addr: (ws[addr].value if ws[addr].value is not None else "") for addr in EXPECTED_HEADERS}
    wb.close()
    return vals

def validate_headers(header_cells: dict) -> list:
    """기대한 헤더와 실제 값 비교. 불일치 메시지 목록 반환"""
    mismatches = []
    for addr, expected in EXPECTED_HEADERS.items():
        got = str(header_cells.get(addr, "")).strip()
        if got != expected:
            mismatches.append(f"  - {addr} 위치의 헤더 불일치 (기대: '{expected}', 실제: '{got}')")
    return mismatches

def time_bucket_x(s: str) -> str:
    """총출장시간 문자열 → '4시간이상'/'4시간미만'/'' 분류"""
    if not isinstance(s, str):
        s = "" if pd.isna(s) else str(s)
    s = s.replace(" ", "")
    if "일" in s:
        return "4시간이상"
    if "시간" in s:
        m = re.search(r"(\d+)\s*시간", s)
        if m and int(m.group(1)) >= 4:
            return "4시간이상"
        return "4시간미만"
    if "분" in s:
        return "4시간미만"
    return ""

def minute_flag_y(s: str) -> str:
    """총출장시간이 '분'만 있을 때 '1시간미만' 플래그"""
    if not isinstance(s, str):
        s = "" if pd.isna(s) else str(s)
    s = s.replace(" ", "")
    return "1시간미만" if ("분" in s and "시간" not in s) else ""

def payment_z(x_val: str, k_val: str) -> int:
    """X(4시간 구분)와 차량 사용여부(K)로 지급액 산정"""
    x_val = (x_val or "").strip()
    k_val = (k_val or "").strip()
    if x_val == "4시간이상" and k_val == "미사용": return 20000
    if x_val == "4시간이상" and k_val == "사용":   return 10000
    if x_val == "4시간미만" and k_val == "미사용": return 10000
    if x_val == "4시간미만" and k_val == "사용":   return 0
    return 0

def autofit_ws(ws: Worksheet, max_width=60):
    """열별 최대 텍스트 길이 기반 너비 자동 조정"""
    for col in ws.columns:
        first = next(iter(col), None)
        if first is None:
            continue
        col_letter = get_column_letter(first.column)
        length = 0
        for c in col:
            v = "" if c.value is None else str(c.value)
            # 개행 고려: 가장 긴 줄 기준
            line_max = max((len(line) for line in v.split("\n")), default=0)
            length = max(length, line_max)
        ws.column_dimensions[col_letter].width = min(max(12, length + 2), max_width)

# =========================
# 핵심 처리
# =========================
def create_processed_workbook(uploaded_file) -> BytesIO:
    """업로드 파일 처리 후 openpyxl 워크북을 바이트스트림으로 반환"""
    # 1) 데이터 로드
    engine = "xlrd" if uploaded_file.name.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(uploaded_file, header=3, dtype=str, engine=engine)

    # 2) 컬럼명 상수
    col_B, col_C = "순번", "구분"
    col_F, col_G, col_K, col_O, col_T = "도착일자", "총출장시간", "차량", "출장자", "결재상태"

    # 3) 전처리: 공란 제거
    if col_B in df.columns:
        df = df[~df[col_B].isna() & (df[col_B].astype(str).str.strip() != "")]
    if col_C in df.columns:
        df = df[~df[col_C].isna() & (df[col_C].astype(str).str.strip() != "")]

    # 4) 결재완료 필터
    if col_T in df.columns:
        df = df[df[col_T].astype(str).str.contains("결재완료", na=False)]
    else:
        raise RuntimeError("필수 열 '결재상태'를 찾을 수 없습니다.")

    # 5) 파생열 계산
    df["X"] = df[col_G].apply(time_bucket_x)
    df["Y"] = df[col_G].apply(minute_flag_y)
    df["Z"] = df.apply(lambda r: payment_z(r.get("X", ""), r.get(col_K, "")), axis=1)

    # 6) 날짜 파싱
    f_dt = pd.to_datetime(df[col_F], errors="coerce") if col_F in df.columns else pd.Series([pd.NaT]*len(df))

    # 7) 정렬 키 생성 및 정렬
    x_cat = pd.Categorical(df["X"], categories=["4시간이상", "4시간미만", ""], ordered=True)
    df_sorted = df.copy()
    df_sorted["X_cat"] = x_cat
    df_sorted["Z_num"] = pd.to_numeric(df_sorted["Z"], errors="coerce").fillna(0).astype(int)
    df_sorted = df_sorted.sort_values(["X_cat", "Z_num"], ascending=[True, False]).drop(columns=["X_cat", "Z_num"])

    # 8) 워크북 생성 및 '정리결과' 시트 작성
    wb = Workbook()
    ws = wb.active
    ws.title = "정리결과"

    cols = list(df_sorted.columns)
    for c in ["X", "Y", "Z"]:
        if c in cols:
            cols.remove(c)
    cols = cols + ["X", "Y", "Z"]  # 파생열을 끝으로 이동

    ws.append(cols)
    for _, row in df_sorted[cols].iterrows():
        ws.append(list(row.values))

    # 9) 헤더 이름 교체
    x_col_idx, y_col_idx, z_col_idx = cols.index("X") + 1, cols.index("Y") + 1, cols.index("Z") + 1
    ws.cell(row=1, column=x_col_idx).value = "4시간 구분"
    ws.cell(row=1, column=y_col_idx).value = "1시간 미만"
    ws.cell(row=1, column=z_col_idx).value = "지급액"

    # 10) 서식 적용
    red_font = Font(color="FF0000")
    pale_yellow = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    # 헤더 서식
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = bold
        cell.alignment = center

    # 데이터 서식
    for r in range(2, ws.max_row + 1):
        # 지급액 표기
        cell_z = ws.cell(row=r, column=z_col_idx)
        try:
            val = int(cell_z.value) if cell_z.value not in (None, "") else None
            if val is not None:
                cell_z.number_format = "#,##0"
                cell_z.font = red_font
        except (ValueError, TypeError):
            pass
        # 파생열 배경
        ws.cell(row=r, column=x_col_idx).fill = pale_yellow
        ws.cell(row=r, column=y_col_idx).fill = pale_yellow
        ws.cell(row=r, column=z_col_idx).fill = pale_yellow

    autofit_ws(ws)

    # 11) '요약' 시트 작성
    ws2 = wb.create_sheet("요약")
    ws2.append(["출장자", "지급단가", "합계금액", "출장일자"])

    df_work = pd.DataFrame({
        "출장자": df.get(col_O, ""),
        "지급액": df.get("Z", 0),
        "도착일자": f_dt,
    })
    # 유효 데이터만
    df_work = df_work[
        (~df_work["출장자"].isna()) &
        (df_work["출장자"].astype(str).str.strip() != "") &
        (~df_work["지급액"].isna())
    ]
    df_work["지급액"] = pd.to_numeric(df_work["지급액"], errors="coerce").fillna(0).astype(int)
    df_work = df_work[df_work["지급액"].isin([20000, 10000])]

    # 이름-단가별 날짜 나열
    names = sorted(df_work["출장자"].astype(str).unique())
    for name in names:
        for pay in [20000, 10000]:
            days = (
                df_work[(df_work["출장자"].astype(str) == name) & (df_work["지급액"] == pay)]["도착일자"]
                .dropna().sort_values().apply(lambda x: int(getattr(x, "day", x.day))).tolist()
            )
            if days:
                ws2.append([name, pay, None] + days)

    # 12) '요약' 서식 및 조건부서식
    pink = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
    pale_yellow2 = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
    thin = Side(style="thin", color="FF999999")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 헤더 서식
    for c in range(1, ws2.max_column + 1):
        ws2.cell(row=1, column=c).font = bold
        ws2.cell(row=1, column=c).alignment = center

    # 합계 수식, 정렬
    for r in range(2, ws2.max_row + 1):
        bcell = ws2.cell(row=r, column=2)  # 지급단가
        sum_cell = ws2.cell(row=r, column=3)
        if isinstance(bcell.value, (int, float)):
            bcell.number_format = "#,##0"
            bcell.alignment = center
        last_col_letter = get_column_letter(ws2.max_column)
        sum_cell.value = f"=B{r}*COUNT(D{r}:{last_col_letter}{r})"
        sum_cell.number_format = "#,##0"
        sum_cell.alignment = center
        # 날짜 셀 형식
        for c in range(4, ws2.max_column + 1):
            dcell = ws2.cell(row=r, column=c)
            if dcell.value not in (None, ""):
                dcell.number_format = "0"
                dcell.alignment = center

    # 조건부서식
    if ws2.max_row >= 2:
        ws2.conditional_formatting.add(f"C2:C{ws2.max_row}", CellIsRule(operator="greaterThan", formula=["300000"], fill=pink))
        ws2.conditional_formatting.add(f"A2:A{ws2.max_row}", FormulaRule(formula=[f'COUNTIF($A:$A,$A2)>1'], fill=pale_yellow2))

    # 테두리 및 너비
    for row in ws2.iter_rows():
        for cell in row:
            cell.border = border_all
    autofit_ws(ws2)

    # 13) 요약 시트를 첫 화면으로
    wb.active = wb.worksheets.index(ws2)

    # 14) 메모리 저장 후 반환
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =========================
# Streamlit UI
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(f"📑 {APP_TITLE}")
    st.markdown("✅ 인사랑 '관내출장내역' 파일을 업로드하면 여비를 계산하고 요약 시트를 생성합니다.")
    st.markdown("✅ '팀별 수합자료'와 '저장된 파일'을 비교하여 교차 검증하시면 됩니다.")
    st.markdown("✅ 다운로드 할 파일이 열리는 데 시간이 걸려, 잠시만 기다려주세요.")
    
    # 사이드바: 매뉴얼 다운로드
    with st.sidebar:
        st.header("안내 사항")
        st.info(
            """
            1. (매뉴얼 확인)
            
            하단에 사용법이 담긴 PDF 파일을 확인합니다.
            
            2. (파일 업로드)
            
            '인사랑'에서 내려받은 관내출장내역(.xlsx) 파일을 업로드합니다.
            
            3. (처리 및 다운로드)
            
            파일이 정상적으로 검증되면, 처리 버튼이 활성화됩니다. 버튼을 클릭하여 결과 파일을 다운로드하세요.
            """
        )
        if os.path.exists(MANUAL_FILE):
            with open(MANUAL_FILE, "rb") as pdf_file:
                st.download_button(
                    label="📜 참고 매뉴얼",
                    data=pdf_file,
                    file_name=MANUAL_FILE,
                    mime="application/pdf"
                )
        else:
            st.warning(f"'{MANUAL_FILE}'을 찾을 수 없습니다.")

    # 1. 업로드
    st.header("1. 엑셀 파일 업로드")
    uploaded_file = st.file_uploader(
        "인사랑 '관내출장내역' 엑셀 파일(.xlsx)을 선택하세요.",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            # 헤더 검증용 복사(파일 포인터 보호)
            file_buffer = BytesIO(uploaded_file.getvalue())
            st.info(f"✅ 업로드: **{uploaded_file.name}**")

            # 헤더 검증
            with st.spinner("서식 검증 중..."):
                headers = read_headers_for_check(file_buffer)
                mismatches = validate_headers(headers)

            if mismatches:
                st.error("❌ 파일 서식 오류")
                st.warning("아래 불일치 항목을 확인하세요.")
                st.code("\n".join(mismatches), language="diff")
                return

            st.success("✔️ 서식이 일치합니다. '처리 시작하기'를 누르세요.")
            st.header("2. 처리 및 정리내역 다운로드")

            if st.button("🚀 처리 시작하기", type="primary"):
                with st.spinner("처리 중..."):
                    try:
                        uploaded_file.seek(0)  # 실제 처리 전 포인터 초기화
                        processed_output = create_processed_workbook(uploaded_file)
                        st.success("✔️ 처리가 완료되었습니다. '정리내역'을 다운로드하세요.")
                        default_name = f"관내출장_정리내역_{kst_timestamp()}.xlsx"
                        st.download_button(
                            label="💾 정리내역 다운로드",
                            data=processed_output,
                            file_name=default_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"처리 오류: {e}")
        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")

if __name__ == "__main__":
    main()











