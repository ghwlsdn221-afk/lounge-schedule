import streamlit as st
import pandas as pd
import tempfile
import os

# 기존에 만드신 스케줄 연산 함수들을 가져옵니다.
from ScheduleV1 import (
    solve_global_schedule,
    verify_schedule_checklist,
    export_to_excel_single_sheet,
    LOUNGE_WORKER_BOUNDS,
    LOUNGE_LIST,
    LOUNGE_ALIAS
)

# ---------------------------------------------------------
# [웹 페이지 기본 설정]
# ---------------------------------------------------------
st.set_page_config(
    page_title="VIP Lounge Schedule System", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# [커스텀 CSS 주입] 백화점 VIP 라운지 테마 (블랙 & 샴페인 골드)
# ---------------------------------------------------------
st.markdown("""
<style>
    /* 구글 폰트(명조체) 불러오기 */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@300;400;700&display=swap');
    
    /* 전체 폰트 적용 및 배경색 (차분한 다크톤) */
    html, body, [class*="css"]  {
        font-family: 'Noto Serif KR', serif !important;
    }
    
    .stApp {
        background-color: #121212;
    }

    /* 텍스트 및 헤더 컬러 (샴페인 골드 & 실버 그레이) */
    h1, h2, h3, h4, h5, h6 {
        color: #D4AF37 !important; 
        font-weight: 300 !important;
        letter-spacing: 1.5px;
    }
    
    p, span, label, div {
        color: #C0C0C0 !important;
    }

    /* 사이드바 스타일링 */
    [data-testid="stSidebar"] {
        background-color: #1A1A1A !important;
        border-right: 1px solid #333333 !important;
    }
    
    /* 사이드바 입력창 테두리 골드 포인트 */
    .stNumberInput > div > div > div, .stTextInput > div > div > div {
        border-color: #333333 !important;
    }
    .stNumberInput > div > div > div:focus-within, .stTextInput > div > div > div:focus-within {
        border-color: #D4AF37 !important;
        box-shadow: 0 0 5px rgba(212, 175, 55, 0.3) !important;
    }

    /* 고급스러운 버튼 스타일링 (명품 브랜드 스타일 - 테두리 얇게, 호버시 반전) */
    div.stButton > button:first-child {
        background-color: transparent !important;
        color: #D4AF37 !important;
        border: 1px solid #D4AF37 !important;
        border-radius: 0px !important; /* 각진 테두리로 모던함 강조 */
        padding: 0.5rem 2rem !important;
        transition: all 0.4s ease-in-out !important;
        font-weight: 400 !important;
        letter-spacing: 1px;
    }
    
    div.stButton > button:first-child:hover {
        background-color: #D4AF37 !important;
        color: #121212 !important;
        box-shadow: 0px 4px 15px rgba(212, 175, 55, 0.3) !important;
    }

    /* 파일 업로더 점선 테두리 */
    [data-testid="stFileUploadDropzone"] {
        border: 1px dashed #D4AF37 !important;
        background-color: #181818 !important;
        border-radius: 0px !important;
    }
    
    /* 데이터 프레임 컨테이너 */
    [data-testid="stDataFrame"] {
        border: 1px solid #333333 !important;
    }

    /* Streamlit 기본 브랜딩 숨기기 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 구분선 골드 컬러 */
    hr {
        border-top: 1px solid #D4AF37 !important;
        opacity: 0.3;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# [헤더 영역] 타이틀 중앙 정렬 및 서브 타이틀
# ---------------------------------------------------------
st.markdown("<h1 style='text-align: center; border-bottom: 1px solid #D4AF37; padding-bottom: 20px; margin-bottom: 20px;'>THE LOUNGE <br><span style='font-size: 0.5em;'>INTEGRATED SCHEDULE MANAGER</span></h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.1em; margin-bottom: 40px;'>최상의 고객 서비스를 위한 프리미엄 라운지 통합 스케줄링 시스템입니다.</p>", unsafe_allow_html=True)


# ---------------------------------------------------------
# [사이드바] 연산 설정 영역
# ---------------------------------------------------------
st.sidebar.markdown("<h3>Operation Settings</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<hr>", unsafe_allow_html=True)

year = st.sidebar.number_input("연도 (Year)", value=2026, step=1)
month = st.sidebar.number_input("월 (Month)", value=10, min_value=1, max_value=12, step=1)

st.sidebar.markdown("<br>", unsafe_allow_html=True)

male_off = st.sidebar.number_input("남성 목표 휴무일수", value=11, step=1)
female_off = st.sidebar.number_input("여성 목표 휴무일수", value=12, step=1)
holidays_str = st.sidebar.text_input("공휴일 지정 (쉼표 구분)", value="3, 9")

# 공휴일 텍스트를 리스트로 변환
public_holidays = [int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()]

st.sidebar.markdown("<br><br><br><p style='text-size:0.8em; color:#666 !important; text-align:center;'>Powered by OR-Tools Engine</p>", unsafe_allow_html=True)

# ---------------------------------------------------------
# [메인 화면] 파일 업로드 및 스케줄 생성
# ---------------------------------------------------------
st.markdown("### Ⅰ. 직원 명단 업로드 (Upload Roster)")
uploaded_file = st.file_uploader("직원 데이터 CSV 파일을 업로드해 주십시오.", type=["csv"])

if uploaded_file is not None:
    # 1. 업로드된 CSV 데이터 미리보기
    try:
        df_input = pd.read_csv(uploaded_file, encoding="utf-8-sig")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df_input = pd.read_csv(uploaded_file, encoding="cp949")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### Ⅱ. 불러온 직원 목록 (Employee Roster)")
    st.dataframe(df_input, use_container_width=True)
    
    st.markdown("<br><hr><br>", unsafe_allow_html=True)

    # 2. 스케줄 생성 버튼 (디자인 적용됨)
    st.markdown("### Ⅲ. 스케줄 최적화 (Optimization)")
    if st.button("스케줄 자동 생성 시작", type="primary"):
        with st.spinner("최적화 엔진 연산 중입니다. 잠시만 기다려 주십시오..."):
            
            # 직원 데이터 구조화 (ScheduleV1 형식 변환)
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

            # OR-Tools 스케줄 연산 실행
            flat_labels = solve_global_schedule(
                all_employees_flat, year, month, male_off, female_off, LOUNGE_WORKER_BOUNDS, public_holidays
            )

            if flat_labels is None:
                st.error("스케줄 생성 실패: 조건에 맞는 스케줄 조합을 찾을 수 없습니다. (휴무 조건 완화 요망)")
            else:
                st.success("스케줄 생성이 성공적으로 완료되었습니다.")

                st.markdown("<br>", unsafe_allow_html=True)
                
                # 3. 검증 체크리스트 표시
                st.markdown("### Ⅳ. 스케줄 검증 리포트 (Verification Report)")
                checklist = verify_schedule_checklist(
                    all_employees_flat, flat_labels, year, month, male_off, female_off
                )
                
                chk_df = pd.DataFrame(checklist, columns=["점검 항목", "검증 기준", "점검 결과", "세부 보고 내용"])
                st.dataframe(chk_df, use_container_width=True)

                st.markdown("<br><hr><br>", unsafe_allow_html=True)

                # 4. 엑셀 파일 다운로드 제공
                st.markdown("### Ⅴ. 결과물 다운로드 (Export)")
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