import streamlit as st
import pandas as pd
import tempfile
import os
import datetime

# 기존 스케줄 연산 함수 임포트
from ScheduleV1 import (
    solve_global_schedule,
    verify_schedule_checklist,
    export_to_excel_single_sheet,
    LOUNGE_WORKER_BOUNDS,
    LOUNGE_LIST,
    LOUNGE_ALIAS
)

# =========================================================
# 🛑 관리자 설정: 구글 시트 CSV 게시 링크를 아래에 입력하세요.
# =========================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1hsHa9MvFwMs3mdDA4u6Fb5St7xEdQltSjcaY6OgcfbA/edit?usp=sharing"

# ---------------------------------------------------------
# [웹 페이지 기본 설정 및 커스텀 CSS]
# ---------------------------------------------------------
st.set_page_config(page_title="VIP Lounge Schedule System", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@300;400;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Noto Serif KR', serif !important; }
    .stApp { background-color: #121212; }
    h1, h2, h3, h4, h5, h6 { color: #D4AF37 !important; font-weight: 300 !important; letter-spacing: 1.5px; }
    p, span, label, div { color: #C0C0C0 !important; }
    [data-testid="stSidebar"] { background-color: #1A1A1A !important; border-right: 1px solid #333333 !important; }
    .stNumberInput > div > div > div, .stTextInput > div > div > div { border-color: #333333 !important; }
    .stNumberInput > div > div > div:focus-within, .stTextInput > div > div > div:focus-within {
        border-color: #D4AF37 !important; box-shadow: 0 0 5px rgba(212, 175, 55, 0.3) !important;
    }
    div.stButton > button:first-child {
        background-color: transparent !important; color: #D4AF37 !important; border: 1px solid #D4AF37 !important;
        border-radius: 0px !important; padding: 0.5rem 2rem !important; transition: all 0.4s ease-in-out !important;
    }
    div.stButton > button:first-child:hover { background-color: #D4AF37 !important; color: #121212 !important; }
    [data-testid="stDataFrame"] { border: 1px solid #333333 !important; }
    #MainMenu, footer {visibility: hidden;}
    hr { border-top: 1px solid #D4AF37 !important; opacity: 0.3; }
    
    /* 대시보드 메트릭 박스 스타일링 */
    [data-testid="stMetricValue"] { color: #D4AF37 !important; font-size: 1.8rem !important; }
    [data-testid="stMetricLabel"] { color: #C0C0C0 !important; font-size: 1rem !important; margin-bottom: 5px; }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; border-bottom: 1px solid #D4AF37; padding-bottom: 20px; margin-bottom: 20px;'>THE LOUNGE <br><span style='font-size: 0.5em;'>INTEGRATED SCHEDULE MANAGER</span></h1>", unsafe_allow_html=True)

# ---------------------------------------------------------
# [사이드바] 연산 설정
# ---------------------------------------------------------
st.sidebar.markdown("<h3>Operation Settings</h3><hr>", unsafe_allow_html=True)

# 다음 달 연산이 기본값이 되도록 자동 세팅
today = datetime.date.today()
default_month = today.month + 1 if today.month < 12 else 1
default_year = today.year if today.month < 12 else today.year + 1

year = st.sidebar.number_input("연도 (Year)", value=default_year, step=1)
month = st.sidebar.number_input("월 (Month)", value=default_month, min_value=1, max_value=12, step=1)
st.sidebar.markdown("<br>", unsafe_allow_html=True)
male_off = st.sidebar.number_input("남성 목표 휴무일수", value=11, step=1)
female_off = st.sidebar.number_input("여성 목표 휴무일수", value=12, step=1)
holidays_str = st.sidebar.text_input("공휴일 지정 (쉼표 구분)", value="3, 9")

public_holidays = [int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()]

# ---------------------------------------------------------
# [메인 화면] 실시간 데이터 로드 및 통계
# ---------------------------------------------------------
st.markdown("### Ⅰ. 실시간 제출 현황 (Live Dashboard)")

if SHEET_URL == "여기에_구글_시트_CSV_링크를_붙여넣으세요":
    st.error("app.py 코드 내의 SHEET_URL 변수에 구글 시트 링크를 입력해 주십시오.")
else:
    try:
        df_input = pd.read_csv(SHEET_URL)
        display_df = df_input.drop(columns=["타임스탬프"], errors="ignore")
        
        # 📊 [통계 대시보드 추가]
        total_submitted = len(display_df)
        lounge_col = next((c for c in display_df.columns if "라운지" in c), None)
        
        st.markdown(
            f"<div style='border:1px solid #333; padding:20px; text-align:center; margin-bottom:20px;'>"
            f"<span style='font-size:1.2em;'>전체 제출 인원</span><br>"
            f"<span style='font-size:2.5em; color:#D4AF37; font-weight:bold;'>{total_submitted}</span> 명"
            f"</div>", 
            unsafe_allow_html=True
        )

        if lounge_col and total_submitted > 0:
            counts = display_df[lounge_col].value_counts()
            cols = st.columns(len(counts))
            for i, (lounge_name, count) in enumerate(counts.items()):
                with cols[i]:
                    st.metric(label=f"📍 {lounge_name}", value=f"{count}명")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### Ⅱ. 직원 명단 상세 (Roster Details)")
        st.dataframe(display_df, use_container_width=True)
        st.markdown("<br><hr><br>", unsafe_allow_html=True)

        # ---------------------------------------------------------
        # [스케줄 최적화 연산]
        # ---------------------------------------------------------
        st.markdown("### Ⅲ. 스케줄 최적화 (Optimization)")
        if st.button("스케줄 자동 생성 시작", type="primary"):
            with st.spinner("최적화 엔진 연산 중입니다. 잠시만 기다려 주십시오..."):
                
                name_col = next((c for c in df_input.columns if "이름" in c or "성명" in c), df_input.columns[0])
                gender_col = next((c for c in df_input.columns if "성별" in c), df_input.columns[1])
                off_col = next((c for c in df_input.columns if "휴무" in c), None)
                m_col = next((c for c in df_input.columns if "생리" in c), None)
                rank_col = next((c for c in df_input.columns if "직급" in c), None)

                all_employees_flat = []
                lounge_employees = {lounge: [] for lounge in LOUNGE_LIST}

                for idx, row in df_input.iterrows():
                    name = str(row[name_col]).strip() if pd.notna(row[name_col]) else f"직원{idx+1}"
                    gender = str(row[gender_col]).strip() if pd.notna(row[gender_col]) else "여"
                    raw_lounge = str(row[lounge_col]).strip() if lounge_col and pd.notna(row[lounge_col]) else "YP"
                    lounge = LOUNGE_ALIAS.get(raw_lounge, raw_lounge)
                    rank = str(row[rank_col]).strip() if rank_col and pd.notna(row[rank_col]) else "사원"

                    req_off = []
                    if off_col and pd.notna(row[off_col]):
                        for item in str(row[off_col]).replace(";", ",").split(","):
                            if item.strip().isdigit(): req_off.append(int(item.strip()))

                    m_off = None
                    if m_col and pd.notna(row[m_col]):
                        if str(row[m_col]).strip().isdigit(): m_off = int(str(row[m_col]).strip())

                    emp_dict = {"name": name, "gender": gender, "lounge": lounge, "rank": rank, "req_off": req_off, "m_off": m_off}
                    all_employees_flat.append(emp_dict)
                    lounge_employees.setdefault(lounge, []).append(emp_dict)

                flat_labels = solve_global_schedule(all_employees_flat, year, month, male_off, female_off, LOUNGE_WORKER_BOUNDS, public_holidays)

                if flat_labels is None:
                    st.error("스케줄 생성 실패: 조건에 맞는 스케줄 조합을 찾을 수 없습니다.")
                else:
                    st.success("스케줄 생성이 성공적으로 완료되었습니다.")
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    st.markdown("### Ⅳ. 스케줄 검증 리포트 (Verification Report)")
                    checklist = verify_schedule_checklist(all_employees_flat, flat_labels, year, month, male_off, female_off)
                    chk_df = pd.DataFrame(checklist, columns=["점검 항목", "검증 기준", "점검 결과", "세부 보고 내용"])
                    st.dataframe(chk_df, use_container_width=True)
                    st.markdown("<br><hr><br>", unsafe_allow_html=True)

                    st.markdown("### Ⅴ. 결과물 다운로드 (Export)")
                    lounge_schedules = {lounge: [] for lounge in LOUNGE_LIST}
                    for i, e in enumerate(all_employees_flat):
                        lounge_schedules[e["lounge"]].append(flat_labels[i])

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                        export_to_excel_single_sheet(lounge_schedules, lounge_employees, all_employees_flat, year, month, tmp_excel.name, public_holidays, male_off, female_off)
                        with open(tmp_excel.name, "rb") as f:
                            st.download_button(
                                label="생성된 엑셀 파일 다운로드 (.xlsx)",
                                data=f.read(),
                                file_name=f"라운지_월간근무표_{year}년_{month}월.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
    except Exception as e:
        st.error(f"데이터를 불러오는 데 실패했습니다. (오류: {str(e)})")