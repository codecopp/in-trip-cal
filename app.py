# -*- coding: utf-8 -*-
"""
관내출장 정리 자동화 (Streamlit 웹앱 버전)
- 기능 1: 프로그램 매뉴얼 다운로드
- 기능 2: 엑셀 파일 업로드 (서식 즉시 검증)
- 기능 3: 처리 및 결과 다운로드 (파일명 자동 제안)
- 요구사항: '요약' 시트를 첫 화면으로 표시(워크북 활성 시트로 설정)
"""

import os
import re
import platform
import streamlit as st
from io import BytesIO
from datetime import datetime

# tz
try:
    from zoneinfo import ZoneInfo
    KST = ZoneInfo("Asia/Seoul")
except ImportError:
    # Python < 3.9 에서는 pytz 사용
    from pytz import timezone
    KST = timezone("Asia/Seoul")


import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import CellIsRule, FormulaRule

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

APP_TITLE = "관내출장내역 정리 자동화"
MANUAL_FILE = "인사랑 관내출장 내역 추출.pdf"

# =========================
# 공통 유틸 (웹 환경에 맞게 일부 수정)
# =========================

def kst_now():
    return datetime.now(KST)

def kst_timestamp() -> str:
    return kst_now().strftime("%y%m%d_%H%M")

# file_like_object는 Streamlit의 UploadedFile 객체
def read_headers_for_check(file_like_object) -> dict:
    wb = load_workbook(file_like_object, data_only=True, read_only=True)
    ws = wb.active
    vals = {addr: (ws[addr].value if ws[addr].value is not None else "") for addr in EXPECTED_HEADERS}
    wb.close()
    return vals

def validate_headers(header_cells: dict) -> list:
    mismatches = []
    for addr, expected in EXPECTED_HEADERS.items():
        got = str(header_cells.get(addr, "")).strip()
        if got != expected:
            mismatches.append(f"  - {addr} 위치의 헤더 불일치 (기대: '{expected}', 실제: '{got}')")
    return mismatches

def time_bucket_x(s: str) -> str:
    if not isinstance(s, str):
        s = "" if pd.isna(s) else str(s)
    s = s.replace(" ", "")
    if "일" in s:
        return "4시간이상"
    if "시간" in s:
        m = re.search(r"(\d+)\s*시간", s)
        if m:
            h = int(m.group(1))
            return "4시간이상" if h >= 4 else "4시간미만"
        return "4시간미만"
    if "분" in s and "시간" not in s:
        return "4시간미만"
    return ""

def minute_flag_y(s: str) -> str:
    if not isinstance(s, str):
        s = "" if pd.isna(s) else str(s)
    s = s.replace(" ", "")
    return "1시간미만" if ("분" in s and "시간" not in s) else ""

def payment_z(x_val: str, k_val: str) -> int:
    x_val = (x_val or "").strip()
    k_val = (k_val or "").strip()
    if x_val == "4시간이상" and k_val == "미사용": return 20000
    if x_val == "4시간이상" and k_val == "사용":   return 10000
    if x_val == "4시간미만" and k_val == "미사용": return 10000
    if x_val == "4시간미만" and k_val == "사용":   return 0
    return 0

def autofit_ws(ws: Worksheet, max_width=60):
    for col in ws.columns:
        first = next(iter(col), None)
        if first is None:
            continue
        col_letter = get_column_letter(first.column)
        length = 0
        for c in col:
            v = c.value
            v = "" if v is None else str(v)
            length = max(length, len(v.split('\n'))) # 개행 문자 고려
            length = max(length, len(v))

        ws.column_dimensions[col_letter].width = min(max(12, length + 2), max_width)


# =========================
# 핵심 처리
# =========================
def create_processed_workbook(uploaded_file):
    # Streamlit의 UploadedFile 객체는 read 메서드를 가지고 있어 pandas/openpyxl에서 바로 사용 가능
    # .xls 파일도 처리하기 위해 xlrd 엔진 명시
    engine = "xlrd" if uploaded_file.name.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(uploaded_file, header=3, dtype=str, engine=engine)

    col_B, col_C = "순번", "구분"
    col_F, col_G, col_K, col_O, col_T = "도착일자", "총출장시간", "차량", "출장자", "결재상태"

    if col_B in df.columns:
        df = df[~df[col_B].isna() & (df[col_B].astype(str).str.strip() != "")]
    if col_C in df.columns:
        df = df[~df[col_C].isna() & (df[col_C].astype(str).str.strip() != "")]

    if col_T in df.columns:
        df = df[df[col_T].astype(str).str.contains("결재완료", na=False)]
    else:
        raise RuntimeError("필수 열 '결재상태'를 찾을 수 없습니다.")

    df["X"] = df[col_G].apply(time_bucket_x)
    df["Y"] = df[col_G].apply(minute_flag_y)
    df["Z"] = df.apply(lambda r: payment_z(r.get("X", ""), r.get(col_K, "")), axis=1)

    f_dt = pd.to_datetime(df[col_F], errors="coerce") if col_F in df.columns else pd.Series([pd.NaT]*len(df))

    x_cat = pd.Categorical(df["X"], categories=["4시간이상", "4시간미만", ""], ordered=True)
    df_sorted = df.copy()
    df_sorted["X_cat"] = x_cat
    df_sorted["Z_num"] = pd.to_numeric(df_sorted["Z"], errors="coerce").fillna(0).astype(int)
    df_sorted = df_sorted.sort_values(["X_cat", "Z_num"], ascending=[True, False]).drop(columns=["X_cat", "Z_num"])

    wb = Workbook()
    ws = wb.active
    ws.title = "정리결과"

    # 정리결과 시트 헤더 구성 및 이름 변경
    cols = list(df_sorted.columns)
    for c in ["X", "Y", "Z"]:
        if c in cols:
            cols.remove(c)
    cols = cols + ["X", "Y", "Z"]

    ws.append(cols)
    for _, row in df_sorted[cols].iterrows():
        ws.append(list(row.values))

    # 헤더 표시명 교체
    x_col_idx, y_col_idx, z_col_idx = cols.index("X") + 1, cols.index("Y") + 1, cols.index("Z") + 1
    ws.cell(row=1, column=x_col_idx).value = "4시간 구분"
    ws.cell(row=1, column=y_col_idx).value = "1시간 미만"
    ws.cell(row=1, column=z_col_idx).value = "지급액"

    # 서식 적용
    red_font = Font(color="FF0000")
    pale_yellow = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    for r in range(2, ws.max_row + 1):
        # 지급액 서식
        cell_z = ws.cell(row=r, column=z_col_idx)
        try:
            val = int(cell_z.value) if cell_z.value not in (None, "") else None
            if val is not None:
                cell_z.number_format = "#,##0"
                cell_z.font = red_font
        except (ValueError, TypeError):
            pass
        # X, Y, Z 열 배경색
        ws.cell(row=r, column=x_col_idx).fill = pale_yellow
        ws.cell(row=r, column=y_col_idx).fill = pale_yellow
        ws.cell(row=r, column=z_col_idx).fill = pale_yellow

    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = bold
        cell.alignment = center
    autofit_ws(ws)

    # 요약 시트
    ws2 = wb.create_sheet("요약")
    ws2.append(["출장자", "지급단가", "합계금액", "출장일자"])

    df_work = pd.DataFrame({
        "출장자": df.get(col_O, ""), "지급액": df.get("Z", 0), "도착일자": f_dt,
    })
    df_work = df_work[
        (~df_work["출장자"].isna()) &
        (df_work["출장자"].astype(str).str.strip() != "") &
        (~df_work["지급액"].isna())
    ]
    df_work["지급액"] = pd.to_numeric(df_work["지급액"], errors="coerce").fillna(0).astype(int)
    df_work = df_work[df_work["지급액"].isin([20000, 10000])]

    names = sorted(df_work["출장자"].astype(str).unique())
    for name in names:
        for pay in [20000, 10000]:
            days = (
                df_work[(df_work["출장자"].astype(str) == name) & (df_work["지급액"] == pay)]["도착일자"]
                .dropna().sort_values().apply(lambda x: int(getattr(x, "day", x.day))).tolist()
            )
            if days:
                ws2.append([name, pay, None] + days)

    # 요약 시트 서식
    pink = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
    pale_yellow2 = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
    thin = Side(style="thin", color="FF999999")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c in range(1, ws2.max_column + 1):
        ws2.cell(row=1, column=c).font = bold
        ws2.cell(row=1, column=c).alignment = center

    for r in range(2, ws2.max_row + 1):
        bcell, sum_cell = ws2.cell(row=r, column=2), ws2.cell(row=r, column=3)
        if isinstance(bcell.value, (int, float)):
            bcell.number_format = "#,##0"
            bcell.alignment = center
        last_col_letter = get_column_letter(ws2.max_column)
        sum_cell.value = f"=B{r}*COUNT(D{r}:{last_col_letter}{r})"
        sum_cell.number_format = "#,##0"
        sum_cell.alignment = center
        for c in range(4, ws2.max_column + 1):
            dcell = ws2.cell(row=r, column=c)
            if dcell.value not in (None, ""):
                dcell.number_format = "0"
                dcell.alignment = center
    
    # 조건부서식
    if ws2.max_row >= 2:
        ws2.conditional_formatting.add(f"C2:C{ws2.max_row}", CellIsRule(operator="greaterThan", formula=["300000"], fill=pink))
        ws2.conditional_formatting.add(f"A2:A{ws2.max_row}", FormulaRule(formula=[f'COUNTIF($A:$A,$A2)>1'], fill=pale_yellow2))

    # 테두리 및 너비 조정
    for row in ws2.iter_rows():
        for cell in row:
            cell.border = border_all
    autofit_ws(ws2)

    # ★ 요약 시트를 첫 화면으로 설정
    wb.active = wb.sheetnames.index(ws2.title)
    
    # 결과를 메모리 내 바이트 스트림으로 저장
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
    st.markdown("인사랑에서 내려받은 '관내출장내역' 엑셀 파일을 업로드하면, 여비를 자동으로 계산하고 요약 시트를 생성해 줍니다.")

    with st.sidebar:
        st.header("사용 방법")
        st.info(
            """
            1. **매뉴얼 확인 (선택)**: 프로그램 사용법이 담긴 PDF 파일을 확인합니다.
            2. **파일 업로드**: '인사랑'에서 내려받은 관내출장내역 `.xls` 또는 `.xlsx` 파일을 업로드합니다.
            3. **처리 및 다운로드**: 파일이 정상적으로 검증되면, 처리 버튼이 활성화됩니다. 버튼을 클릭하여 결과 파일을 다운로드하세요.
            """
        )
        
        # 매뉴얼 다운로드 버튼
        if os.path.exists(MANUAL_FILE):
            with open(MANUAL_FILE, "rb") as pdf_file:
                st.download_button(
                    label="📜 프로그램 매뉴얼 다운로드",
                    data=pdf_file,
                    file_name=MANUAL_FILE,
                    mime="application/pdf"
                )
        else:
            st.warning(f"'{MANUAL_FILE}'을 찾을 수 없습니다.")

    st.header("1. 엑셀 파일 업로드")
    uploaded_file = st.file_uploader(
        "여기를 클릭하여 '관내출장내역' 엑셀 파일을 선택하세요.",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            # 파일을 메모리에 유지하기 위해 복사본 사용
            file_buffer = BytesIO(uploaded_file.getvalue())
            
            st.info(f"✅ 파일이 업로드되었습니다: **{uploaded_file.name}**")
            
            # --- 서식 검증 ---
            with st.spinner("파일 서식을 검증하는 중입니다..."):
                headers = read_headers_for_check(file_buffer)
                mismatches = validate_headers(headers)
            
            if mismatches:
                st.error("❌ 파일 서식 오류")
                st.warning("업로드한 파일의 헤더가 예상과 다릅니다. 아래 불일치 항목을 확인해주세요.")
                st.code("\n".join(mismatches), language="diff")
            else:
                st.success("✔️ 파일 서식이 올바릅니다. 아래 버튼을 눌러 처리를 시작하세요.")
                
                st.header("2. 처리 및 결과 다운로드")
                
                # --- 처리 및 다운로드 ---
                if st.button("🚀 처리 시작하기", type="primary"):
                    with st.spinner("데이터를 처리하고 엑셀 파일을 생성 중입니다... 잠시만 기다려주세요."):
                        try:
                            # 검증 후 실제 처리를 위해 파일 포인터를 다시 처음으로
                            uploaded_file.seek(0)
                            processed_output = create_processed_workbook(uploaded_file)
                            
                            st.success("🎉 처리가 완료되었습니다! 아래 버튼을 클릭하여 결과 파일을 받으세요.")
                            
                            default_name = f"관내출장_정리내역_{kst_timestamp()}.xlsx"
                            st.download_button(
                                label="💾 결과 엑셀 파일 다운로드",
                                data=processed_output,
                                file_name=default_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"처리 중 오류가 발생했습니다: {e}")

        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")


if __name__ == "__main__":
    main()