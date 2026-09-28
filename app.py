import streamlit as st
import pandas as pd
import tempfile
import os

# 기존 스케줄 연산 함수 임포트
from ScheduleV1 import (
    solve_global_schedule,
    verify_schedule_checklist,
    export_to_excel_single_sheet,
    LOUNGE_WORKER_BOUNDS,
    LOUNGE_LIST,
    LOUNGE_ALIAS
)

# ---------------------------------------------------------
# [웹 페이지 기본 설정 및 커스텀 CSS (VIP 라운지 테마)]
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
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# [헤더 영역]
# ---------------------------------------------------------
st.markdown("<h1 style='text-align: center; border-bottom: 1px solid #D4AF37; padding-bottom: 20px; margin-bottom: 20px;'>THE LOUNGE <br><span style='font-size: 0.5em;'>INTEGRATED SCHEDULE MANAGER</span></h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.1em; margin-bottom: 40px;'>구글 시트 실시간 연동 기반 스케줄링 시스템입니다.</p>", unsafe_allow_html=True)

# ---------------------------------------------------------
# [사이드바] 연산 설정 및 DB 연동
# ---------------------------------------------------------
st.sidebar.markdown("<h3>Operation Settings</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<hr>", unsafe_allow_html=True)

# 구글 시트 CSV 게시 링크 입력란 (관리자용)
sheet_url = st.sidebar.text_input("🔗 구글 시트 DB 링크 (CSV)", help="구글 시트에서 '웹에 게시(CSV)'로 생성한 링크를 입력하세요.")

st.sidebar.markdown("<br>", unsafe_allow_html=True)
year = st.sidebar.number_input("연도 (Year)", value=2026, step=1)
month = st.sidebar.number_input("월 (Month)", value=10, min_value=1, max_value=12, step=1)
male_off = st.sidebar.number_input("남성 목표 휴무일수", value=11, step=1)
female_off = st.sidebar.number_input("여성 목표 휴무일수", value=12, step=1)
holidays_str = st.sidebar.text_input("공휴일 지정 (쉼표 구분)", value="3, 9")

public_holidays = [int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()]

# ---------------------------------------------------------
# [메인 화면] 실시간 데이터 로드 및 연산
# ---------------------------------------------------------
st.markdown("### Ⅰ. 실시간 직원 명단 (Live Roster DB)")

if not sheet_url:
    st.info("💡 좌측 톱니바퀴 메뉴(Operation Settings)에 **구글 시트 DB 링크**를 입력하면 명단이 자동으로 로드됩니다.")
else:
    try:
        # 구글 시트에서 실시간으로 데이터 읽어오기
        df_input = pd.read_csv(sheet_url)
        
        # '타임스탬프' 등 불필요한 열이 있으면 숨기고 화면에 표시
        display_df = df_input.drop(columns=["타임스탬프"], errors="ignore")
        st.dataframe(display_df, use_container_width=True)
        st.markdown("<br><hr><br>", unsafe_allow_html=True)

        st.markdown("### Ⅱ. 스케줄 최적화 (Optimization)")
        if st.button("스케줄 자동 생성 시작", type="primary"):
            with st.spinner("최적화 엔진 연산 중입니다. 잠시만 기다려 주십시오..."):
                
                # 구글 폼에서 들어온 데이터를 ScheduleV1 규격에 맞게 매핑
                name_col = next((c for c in df_input.columns if "이름" in c or "성명" in c), df_input.columns[0])
                gender_col = next((c for c in df_input.columns if "성별" in c), df_input.columns[1])
                off_col = next((c for c in df_input.columns if "휴무" in c), None)
                lounge_col = next((c for c in df_input.columns if "라운지" in c), None)
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

                # OR-Tools 연산
                flat_labels = solve_global_schedule(
                    all_employees_flat, year, month, male_off, female_off, LOUNGE_WORKER_BOUNDS, public_holidays
                )

                if flat_labels is None:
                    st.error("스케줄 생성 실패: 조건에 맞는 스케줄 조합을 찾을 수 없습니다. 휴무 조건을 완화하여 폼을 수정해 주십시오.")
                else:
                    st.success("스케줄 생성이 성공적으로 완료되었습니다.")
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    st.markdown("### Ⅲ. 스케줄 검증 리포트 (Verification Report)")
                    checklist = verify_schedule_checklist(all_employees_flat, flat_labels, year, month, male_off, female_off)
                    chk_df = pd.DataFrame(checklist, columns=["점검 항목", "검증 기준", "점검 결과", "세부 보고 내용"])
                    st.dataframe(chk_df, use_container_width=True)
                    st.markdown("<br><hr><br>", unsafe_allow_html=True)

                    st.markdown("### Ⅳ. 결과물 다운로드 (Export)")
                    lounge_schedules = {lounge: [] for lounge in LOUNGE_LIST}
                    for i, e in enumerate(all_employees_flat):
                        lounge_schedules[e["lounge"]].append(flat_labels[i])

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                        export_to_excel_single_sheet(
                            lounge_schedules, lounge_employees, all_employees_flat,
                            year, month, tmp_excel.name, public_holidays, male_off, female_off
                        )
                        
                        with open(tmp_excel.name, "rb") as f:
                            st.download_button(
                                label="생성된 엑셀 파일 다운로드 (.xlsx)",
                                data=f.read(),
                                file_name=f"라운지_월간근무표_{year}년_{month}월.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
    except Exception as e:
        st.error(f"구글 시트 데이터를 불러오는 데 실패했습니다. 링크가 올바른지, 게시 설정이 CSV로 되어있는지 확인해주세요. (오류 메시지: {str(e)})")