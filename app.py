import streamlit as st
import pandas as pd
import tempfile
import os
import datetime
import calendar
import pickle

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
# 🛑 관리자 설정
# =========================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTsq8nya6v_Nf8hOCQC70GsP9dhdtbWZV2pnyTNeozrmJ2ye4vhzVKNEr-8fWV7NSV_WkZ3bL6GIP8K/pub?output=csv"
CACHE_FILE = "cached_schedule.pkl"

# [웹 페이지 기본 설정 및 커스텀 CSS]
st.set_page_config(page_title="VIP Lounge Schedule System", page_icon="👑", layout="wide", initial_sidebar_state="expanded")

# =========================================================
# 🔄 세션 상태 및 로컬 파일 연동 (여기서부터 바로 스케줄 앱 시작)
# =========================================================
if "schedule_generated" not in st.session_state:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "rb") as f:
                st.session_state.schedule_data = pickle.load(f)
            st.session_state.schedule_generated = True
        except Exception:
            st.session_state.schedule_generated = False
            st.session_state.schedule_data = {}
    else:
        st.session_state.schedule_generated = False
        st.session_state.schedule_data = {}

# CSS: 스케줄 앱 메인 화면용 CSS
st.markdown("""
    <style>
    /* 데이터프레임 컨테이너 배경 */
    [data-testid="stDataFrame"] {
        background-color: transparent !important;
    }
    
    /* 셀 내부 배경색 및 테두리 */
    [data-testid="stDataFrame"] div[data-testid="StyledFullScreenButton"] {
        display: none; /* 전체화면 버튼 숨기기 (깔끔한 UI를 위해) */
    }
    
    /* 테이블 헤더 및 셀 강제 다크모드 적용 */
    thead tr th, tbody tr td {
        background-color: #161616 !important;
        color: #E0E0E0 !important;
        border-color: #333333 !important;
    }
    
    /* 헤더 골드 포인트 */
    thead tr th {
        color: #D4AF37 !important;
        border-bottom: 2px solid #D4AF37 !important;
    }
    </style>
""", unsafe_allow_html=True)

# [VIP Lounge] 하단 하얀색 테두리 및 Streamlit 워터마크 강제 제거 CSS
hide_streamlit_ui = """
<style>
    /* 1. 기본 헤더 및 푸터 완벽 숨김 */
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 2. 'Built with Streamlit' 워터마크 및 전체화면 버튼 강제 숨김 */
    .viewerBadge_container__1QSob {display: none !important;}
    .viewerBadge_link__1S137 {display: none !important;}
    div[class^="viewerBadge"] {display: none !important;}
    
    /* 3. 하단 여백 및 하얀색 바(Bottom Bar) 제거 */
    div[data-testid="stBottom"] {display: none !important;}
    
    /* 4. 앱 내부 상하좌우 여백 최소화 (HTML과 자연스럽게 연결되도록) */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0rem !important;
    }
</style>
"""
st.markdown(hide_streamlit_ui, unsafe_allow_html=True)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@300;400;700&display=swap');
    
    html, body, [class*="css"]  { font-family: 'Noto Serif KR', serif !important; }
    .stApp { background-color: #0d0d0d; }
    h1, h2, h3, h4, h5, h6 { color: #D4AF37 !important; font-weight: 400 !important; letter-spacing: 1.5px; }
    p, span, label, div { color: #E0E0E0 !important; }
    
    [data-testid="stSidebar"] { background-color: #161616 !important; border-right: 1px solid #2a2a2a !important; }
    
    .stNumberInput > div > div > div, .stTextInput > div > div > div { 
        border-color: #333333 !important; background-color: #1a1a1a !important; color: #fff !important; 
    }
    .stNumberInput > div > div > div:focus-within, .stTextInput > div > div > div:focus-within {
        border-color: #D4AF37 !important; box-shadow: 0 0 5px rgba(212, 175, 55, 0.4) !important;
    }
    
    div.stButton > button:first-child {
        background-color: #1a1a1a !important; color: #D4AF37 !important; border: 1px solid #D4AF37 !important;
        border-radius: 4px !important; padding: 0.6rem 2rem !important; transition: all 0.3s ease !important;
        font-weight: bold !important; letter-spacing: 1px;
    }
    div.stButton > button:first-child:hover { 
        background-color: #D4AF37 !important; color: #0d0d0d !important; box-shadow: 0 4px 12px rgba(212,175,55,0.3) !important; 
    }
    
    [data-testid="stDataFrame"] { border: 1px solid #333333 !important; border-radius: 8px; overflow: hidden; }
    
    #MainMenu, footer {visibility: hidden;}
    hr { border-top: 1px solid #D4AF37 !important; opacity: 0.2; margin: 2rem 0; }
    
    .total-box {
        background: linear-gradient(145deg, #1a1a1a, #121212); border: 1px solid #333;
        border-top: 3px solid #D4AF37; border-radius: 10px; padding: 30px;
        text-align: center; box-shadow: 0 8px 16px rgba(0,0,0,0.4); margin-bottom: 30px;
    }
    .total-title { font-size:1.1em; color:#A0A0A0; letter-spacing:1px; }
    .total-count { font-size:3.5em; color:#D4AF37; font-weight:700; line-height:1; }
    
    .lounge-card {
        background-color: #161616; border: 1px solid #2a2a2a; border-top: 3px solid #D4AF37; 
        border-radius: 8px; padding: 20px 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .lounge-title { color: #A0A0A0; font-size: 14px; font-weight: 400; margin-bottom: 10px; letter-spacing: 1px; }
    .lounge-count { color: #D4AF37; font-size: 26px; font-weight: bold; }
    .lounge-unit { font-size: 13px; color: #666; font-weight: normal; }

    @media (max-width: 768px) {
        h1 { font-size: 1.8rem !important; }
        .total-box { padding: 15px; margin-bottom: 15px; }
        .total-title { font-size: 0.9em; }
        .total-count { font-size: 2.5em; }
        
        .lounge-card { padding: 15px 5px; margin-bottom: 10px; }
        .lounge-title { font-size: 12px; margin-bottom: 5px; }
        .lounge-count { font-size: 20px; }
        .lounge-unit { font-size: 11px; }
        
        div.stButton > button:first-child { width: 100% !important; padding: 1rem !important; font-size: 1.1rem !important; }
        
        .mobile-scroll-hint { display: block !important; color: #888; font-size: 0.85em; text-align: right; margin-bottom: 5px; }
    }
    
    .mobile-scroll-hint { display: none; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style='text-align: center; border-bottom: 1px solid rgba(212, 175, 55, 0.3); padding-bottom: 25px; margin-bottom: 30px;'>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# [사이드바] 연산 설정
# ---------------------------------------------------------
st.sidebar.markdown("<h3>Operation Settings</h3><hr style='margin: 1rem 0;'>", unsafe_allow_html=True)

today = datetime.date.today()
default_month = today.month + 1 if today.month < 12 else 1
default_year = today.year if today.month < 12 else today.year + 1

year = st.sidebar.number_input("연도 (Year)", value=default_year, step=1)
month = st.sidebar.number_input("월 (Month)", value=default_month, min_value=1, max_value=12, step=1)
st.sidebar.markdown("<br>", unsafe_allow_html=True)
male_off = st.sidebar.number_input("남성 목표 휴무일수", value=11, step=1)
female_off = st.sidebar.number_input("여성 목표 휴무일수", value=12, step=1)
holidays_str = st.sidebar.text_input("공휴일 지정 (쉼표 구분)", value="3, 9")
closed_days_str = st.sidebar.text_input("휴점일 지정 (쉼표 구분)", value="19")  # 💡 휴점일 입력 추가

public_holidays = [int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()]
store_closed_days = [int(x.strip()) for x in closed_days_str.split(",") if x.strip().isdigit()]  # 💡 휴점일 리스트 변환

def color_schedule_cells(val):
    val_str = str(val).strip()
    if "휴점" in val_str:  # 💡 휴점일 셀 강조 색상 추가 (짙은 회색)
        return 'color: #FFFFFF; font-weight: bold; background-color: #555555;'
    elif any(keyword in val_str for keyword in ["휴무", "생휴", "연차", "공휴", "반휴", "휴"]):
        return 'color: #FF6B6B; font-weight: bold; background-color: #3A1C1C;'
    elif any(keyword in val_str for keyword in ["근무", "주", "야", "오픈", "마감", "미들"]):
        return 'color: #4D96FF; font-weight: bold; background-color: #1C2A3A;'
    return ''

target_order = ["자데", "자홀", "블랙", "블루", "세이지", "YP"]

# 빈칸 제거 및 전체 셀에 다크모드 배경색을 입히는 함수
def apply_dark_style(df):
    return df.fillna("").style.set_properties(**{
        'background-color': '#161616',
        'color': '#E0E0E0',
        'border-color': '#333333'
    })

def get_order_weight(name):
    for idx, target in enumerate(target_order):
        if target in str(name):
            return idx
    return 999 

# ---------------------------------------------------------
# [메인 화면] 실시간 데이터 로드 및 통계
# ---------------------------------------------------------
st.markdown("### Ⅰ. 총원 (Dashboard)")

if SHEET_URL == "여기에_구글_시트_CSV_링크를_붙여넣으세요":
    st.error("app.py 코드 내의 SHEET_URL 변수에 구글 시트 링크를 입력해 주십시오.")
else:
    try:
        df_input = pd.read_csv(SHEET_URL)
        display_df = df_input.drop(columns=["타임스탬프"], errors="ignore")
        
        total_submitted = len(display_df)
        lounge_col = next((c for c in display_df.columns if "라운지" in c), None)
        
        st.markdown(
            f"<div class='total-box'>"
            f"<span class='total-title'>등록 인원 (Saved Personnel)</span><br><br>"
            f"<span class='total-count'>{total_submitted}</span>"
            f"<span style='font-size:1.5em; color:#666 !important; margin-left:10px;'>명</span>"
            f"</div>", 
            unsafe_allow_html=True
        )

        if lounge_col and total_submitted > 0:
            lounge_counts = display_df[lounge_col].value_counts().to_dict()
            sorted_lounges = sorted(lounge_counts.keys(), key=get_order_weight)
            
            cols = st.columns(len(sorted_lounges))
            for i, lounge_name in enumerate(sorted_lounges):
                count = lounge_counts[lounge_name]
                with cols[i]:
                    card_html = f"""
                    <div class="lounge-card">
                        <div class="lounge-title">{lounge_name}</div>
                        <div class="lounge-count">{count}<span class="lounge-unit"> 명</span></div>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)
        
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### Ⅱ. 직원 명단 상세 (Roster Details)")
        st.markdown("<span class='mobile-scroll-hint'>👉 표를 좌우로 스크롤하여 확인하세요</span>", unsafe_allow_html=True)
        st.dataframe(apply_dark_style(display_df), use_container_width=True)
        st.markdown("<br><hr>", unsafe_allow_html=True)

        # ---------------------------------------------------------
        # [스케줄 최적화 연산]
        # ---------------------------------------------------------
        st.markdown("### Ⅲ. 스케줄 최적화 (Optimization)")
        if st.button("✨ 스케줄 자동 생성 시작", type="primary"):
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

                    emp_dict = {"name": name, "gender": gender, "lounge": lounge, "rank": rank, "req_off": req_off, "m_off": m_off, "raw_lounge": raw_lounge}
                    all_employees_flat.append(emp_dict)
                    
                    if lounge in lounge_employees:
                        lounge_employees[lounge].append(emp_dict)

                # 💡 함수 호출 시 store_closed_days 인자 추가 전달
                flat_labels = solve_global_schedule(
                    all_employees_flat, year, month, male_off, female_off, 
                    LOUNGE_WORKER_BOUNDS, public_holidays, store_closed_days
                )

                if flat_labels is None:
                    st.error("❌ 스케줄 생성 실패: 조건에 맞는 스케줄 조합을 찾을 수 없습니다.")
                    st.session_state.schedule_generated = False
                else:
                    st.success("✅ 스케줄 생성이 성공적으로 완료되었습니다.")
                    
                    new_schedule_data = {
                        "all_employees_flat": all_employees_flat,
                        "lounge_employees": lounge_employees,
                        "flat_labels": flat_labels,
                        "year": year,
                        "month": month,
                        "male_off": male_off,
                        "female_off": female_off,
                        "public_holidays": public_holidays,
                        "store_closed_days": store_closed_days  # 💡 세션에 휴점일 데이터 보존
                    }
                    st.session_state.schedule_generated = True
                    st.session_state.schedule_data = new_schedule_data
                    
                    with open(CACHE_FILE, "wb") as f:
                        pickle.dump(new_schedule_data, f)

        # ---------------------------------------------------------
        # [결과 출력 영역]
        # ---------------------------------------------------------
        if st.session_state.schedule_generated:
            st.markdown("<br>", unsafe_allow_html=True)
            
            data = st.session_state.schedule_data
            all_emp = data["all_employees_flat"]
            l_emp = data["lounge_employees"]
            labels = data["flat_labels"]
            s_year = data["year"]
            s_month = data["month"]
            
            _, num_days = calendar.monthrange(s_year, s_month)
            day_columns = [f"{d}일" for d in range(1, num_days + 1)]
            
            st.markdown("### Ⅳ. 생성된 스케줄 결과 (Generated Schedule)")
            st.markdown("<span class='mobile-scroll-hint'>👉 표를 좌우로 스크롤하여 전체 근무표를 확인하세요</span>", unsafe_allow_html=True)
            
            schedule_data = []
            for i, emp in enumerate(all_emp):
                schedule_row = list(labels[i])
                if len(schedule_row) < num_days:
                    schedule_row += [""] * (num_days - len(schedule_row))
                elif len(schedule_row) > num_days:
                    schedule_row = schedule_row[:num_days]
                    
                row_data = [emp["raw_lounge"], emp["name"], emp["gender"], emp["rank"]] + schedule_row
                schedule_data.append(row_data)

            columns = ["라운지", "이름", "성별", "직급"] + day_columns
            res_df = pd.DataFrame(schedule_data, columns=columns)
            
            for lounge_kw in target_order:
                df_lounge = res_df[res_df['라운지'].str.contains(lounge_kw, na=False)]
                if not df_lounge.empty:
                    st.markdown(f"<h5 style='color: #D4AF37; margin-top: 20px; border-left: 4px solid #D4AF37; padding-left: 10px;'>{lounge_kw}</h5>", unsafe_allow_html=True)
                    st.dataframe(apply_dark_style(df_lounge).map(color_schedule_cells, subset=day_columns), use_container_width=True)

            other_mask = ~res_df['라운지'].str.contains('|'.join(target_order), na=False)
            df_other = res_df[other_mask]
            if not df_other.empty:
                st.markdown(f"<h5 style='color: #D4AF37; margin-top: 20px; border-left: 4px solid #D4AF37; padding-left: 10px;'>기타 라운지</h5>", unsafe_allow_html=True)
                st.dataframe(apply_dark_style(df_other).map(color_schedule_cells, subset=day_columns), use_container_width=True)
                
            st.markdown("<br><hr>", unsafe_allow_html=True)

            # ---------------------------------------------------------
            # Ⅴ. 스케줄 요약 및 통계
            # ---------------------------------------------------------
            st.markdown("### Ⅴ. 스케줄 요약 및 일자별 통계 (Summary & Daily Stats)")
            
            st.markdown("<h5 style='color: #D4AF37; margin-top: 20px;'>1. 개인별 근무 요약</h5>", unsafe_allow_html=True)
            summary_data = []
            for i, row in res_df.iterrows():
                work_days = row[day_columns].isin(["근무", "주", "야", "오픈", "마감", "미들"]).sum()
                off_days = row[day_columns].isin(["휴무", "휴", "반휴", "휴점"]).sum()  # 💡 통계 합산에 휴점 포함
                m_off_days = row[day_columns].isin(["생휴"]).sum()
                
                summary_data.append({
                    "라운지": row["라운지"],
                    "이름": row["이름"],
                    "성별": row["성별"],
                    "총 근무일수": work_days,
                    "총 휴무일수 (일반)": off_days,
                    "생리휴가 사용일수": m_off_days
                })
                
            summary_df = pd.DataFrame(summary_data)
            
            def highlight_stats(val):
                if isinstance(val, int) and val > 0:
                    return 'color: #D4AF37; font-weight: bold;'
                return ''
                
            st.dataframe(summary_df.style.map(highlight_stats, subset=["총 근무일수", "총 휴무일수 (일반)", "생리휴가 사용일수"]), use_container_width=True)
            
            st.markdown("<h5 style='color: #D4AF37; margin-top: 30px;'>2. 일자별 전체 근무 통계</h5>", unsafe_allow_html=True)
            
            daily_stats = []
            total_working = [res_df[day].isin(["근무", "주", "야", "오픈", "마감", "미들"]).sum() for day in day_columns]
            total_off = [res_df[day].isin(["휴무", "휴", "생휴", "반휴", "휴점"]).sum() for day in day_columns]  # 💡 일자별 휴무 인원에 휴점 포함
            
            daily_stats.append(["총 출근 인원"] + total_working)
            daily_stats.append(["총 휴무 인원"] + total_off)
            
            for lounge_kw in target_order:
                df_l = res_df[res_df['라운지'].str.contains(lounge_kw, na=False)]
                if not df_l.empty:
                    lounge_working = [df_l[day].isin(["근무", "주", "야", "오픈", "마감", "미들"]).sum() for day in day_columns]
                    daily_stats.append([f"{lounge_kw} 출근 인원"] + lounge_working)
            
            daily_stats_df = pd.DataFrame(daily_stats, columns=["구분"] + day_columns)
            daily_stats_df = daily_stats_df.set_index("구분")
            
            def style_daily_stats(val):
                if isinstance(val, (int, float)):
                    if val < 2:  
                        return 'color: #FF6B6B; font-weight: bold;'
                    elif val >= 4: 
                        return 'color: #4D96FF; font-weight: bold;'
                return ''
                
            st.dataframe(apply_dark_style(daily_stats_df).map(style_daily_stats), use_container_width=True)
            st.markdown("<br><hr>", unsafe_allow_html=True)
            
            # ---------------------------------------------------------
            # Ⅵ. 스케줄 검증 리포트
            # ---------------------------------------------------------
            st.markdown("### Ⅵ. 스케줄 검증 리포트 (Verification Report)")
            st.markdown("<span class='mobile-scroll-hint'>👉 표를 좌우로 스크롤하여 확인하세요</span>", unsafe_allow_html=True)
            
            # 💡 체크리스트 검증 함수에 휴점일 데이터 전달
            checklist = verify_schedule_checklist(all_emp, labels, s_year, s_month, data["male_off"], data["female_off"], data.get("store_closed_days", []))
            chk_df = pd.DataFrame(checklist, columns=["점검 항목", "검증 기준", "점검 결과", "세부 보고 내용"])
            
            def highlight_result(val):
                if val == "PASS":
                    return 'color: #00FF00; font-weight: bold;'
                elif val == "FAIL":
                    return 'color: #FF4B4B; font-weight: bold;'
                elif val == "WARNING":
                    return 'color: #FFA500; font-weight: bold;'
                return ''
            
            st.dataframe(apply_dark_style(chk_df).map(highlight_result, subset=['점검 결과']), use_container_width=True)
            st.markdown("<br><hr>", unsafe_allow_html=True)

            # ---------------------------------------------------------
            # Ⅶ. 엑셀 다운로드
            # ---------------------------------------------------------
            st.markdown("### Ⅶ. 엑셀 다운로드 (Export to Excel)")
            lounge_schedules = {lounge: [] for lounge in LOUNGE_LIST}
            for i, e in enumerate(all_emp):
                if e["lounge"] in lounge_schedules:
                    lounge_schedules[e["lounge"]].append(labels[i])

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                # 💡 엑셀 추출 함수에 휴점일 데이터 전달
                export_to_excel_single_sheet(
                    lounge_schedules, l_emp, all_emp, s_year, s_month, 
                    tmp_excel.name, data["public_holidays"], data["male_off"], data["female_off"], data.get("store_closed_days", [])
                )
                with open(tmp_excel.name, "rb") as f:
                    st.download_button(
                        label="💾 생성된 엑셀 파일 다운로드 (.xlsx)",
                        data=f.read(),
                        file_name=f"라운지_월간근무표_{s_year}년_{s_month}월.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

    except Exception as e:
        st.error(f"데이터 연산 중 오류가 발생했습니다. (상세 오류: {str(e)})")