import streamlit as st
import pandas as pd
import tempfile
import os
import datetime
import calendar
import pickle  # 데이터를 로컬 파일로 저장하고 불러오기 위한 모듈 추가

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
SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTsq8nya6v_Nf8hOCQC70GsP9dhdtbWZV2pnyTNeozrmJ2ye4vhzVKNEr-8fWV7NSV_WkZ3bL6GIP8K/pub?output=csv"
CACHE_FILE = "cached_schedule.pkl"  # 생성된 스케줄을 저장할 내부 파일명

# ---------------------------------------------------------
# [웹 페이지 기본 설정 및 커스텀 CSS]
# ---------------------------------------------------------
st.set_page_config(page_title="VIP Lounge Schedule System", layout="wide", initial_sidebar_state="expanded")

# =========================================================
# 🔄 세션 상태 및 로컬 파일 연동 (새로고침 방어 로직)
# =========================================================
if "schedule_generated" not in st.session_state:
    # 앱이 처음 켜졌거나 새로고침 되었을 때, 로컬에 저장된 캐시 파일이 있는지 확인
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
        background: linear-gradient(145deg, #1a1a1a, #121212);
        border: 1px solid #333;
        border-top: 3px solid #D4AF37;
        border-radius: 10px;
        padding: 30px;
        text-align: center;
        box-shadow: 0 8px 16px rgba(0,0,0,0.4);
        margin-bottom: 30px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style='text-align: center; border-bottom: 1px solid rgba(212, 175, 55, 0.3); padding-bottom: 25px; margin-bottom: 30px;'>
    <h1 style='margin-bottom: 0;'>현대백화점 판교점</h1>
    <span style='font-size: 1.2em; color: #888 !important; letter-spacing: 3px; font-weight: 300;'>VIP LOUNGE SCHEDULE MANAGER</span>
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

public_holidays = [int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()]

# 스케줄 텍스트 색상 결정 함수
def color_schedule_cells(val):
    val_str = str(val).strip()
    if any(keyword in val_str for keyword in ["휴무", "생휴", "연차", "공휴", "반휴", "휴"]):
        return 'color: #FF6B6B; font-weight: bold; background-color: #3A1C1C;'
    elif any(keyword in val_str for keyword in ["근무", "주", "야", "오픈", "마감", "미들"]):
        return 'color: #4D96FF; font-weight: bold; background-color: #1C2A3A;'
    return ''

target_order = ["자데", "자홀", "블랙", "블루", "세이지"]

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
            f"<span style='font-size:1.1em; color:#A0A0A0 !important; letter-spacing:1px;'>등록 인원 (Saved Personnel)</span><br><br>"
            f"<span style='font-size:3.5em; color:#D4AF37; font-weight:700; line-height:1;'>{total_submitted}</span>"
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
                    <div style="background-color: #161616; border: 1px solid #2a2a2a; border-top: 3px solid #D4AF37; border-radius: 8px; padding: 20px 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                        <div style="color: #A0A0A0; font-size: 14px; font-weight: 400; margin-bottom: 10px; letter-spacing: 1px;">{lounge_name}</div>
                        <div style="color: #D4AF37; font-size: 26px; font-weight: bold;">{count}<span style="font-size: 13px; color: #666; font-weight: normal;"> 명</span></div>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)
        
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### Ⅱ. 직원 명단 상세 (Roster Details)")
        st.dataframe(display_df, use_container_width=True)
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

                flat_labels = solve_global_schedule(all_employees_flat, year, month, male_off, female_off, LOUNGE_WORKER_BOUNDS, public_holidays)

                if flat_labels is None:
                    st.error("❌ 스케줄 생성 실패: 조건에 맞는 스케줄 조합을 찾을 수 없습니다.")
                    st.session_state.schedule_generated = False
                else:
                    st.success("✅ 스케줄 생성이 성공적으로 완료되었습니다.")
                    
                    # 새 결과를 세션 상태에 저장
                    new_schedule_data = {
                        "all_employees_flat": all_employees_flat,
                        "lounge_employees": lounge_employees,
                        "flat_labels": flat_labels,
                        "year": year,
                        "month": month,
                        "male_off": male_off,
                        "female_off": female_off,
                        "public_holidays": public_holidays
                    }
                    st.session_state.schedule_generated = True
                    st.session_state.schedule_data = new_schedule_data
                    
                    # 💡 핵심: 앱 구동 환경(로컬)에 결과를 파일로 덮어씌워 영구 저장
                    with open(CACHE_FILE, "wb") as f:
                        pickle.dump(new_schedule_data, f)

        # ---------------------------------------------------------
        # [결과 출력 영역] - 세션/캐시에 데이터가 있으면 항상 출력
        # ---------------------------------------------------------
        if st.session_state.schedule_generated:
            st.markdown("<br>", unsafe_allow_html=True)
            
            data = st.session_state.schedule_data
            all_emp = data["all_employees_flat"]
            l_emp = data["lounge_employees"]
            labels = data["flat_labels"]
            s_year = data["year"]
            s_month = data["month"]
            
            st.markdown("### Ⅳ. 생성된 스케줄 결과 (Generated Schedule)")
            
            _, num_days = calendar.monthrange(s_year, s_month)
            day_columns = [f"{d}일" for d in range(1, num_days + 1)]
            
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
                    st.dataframe(df_lounge.style.map(color_schedule_cells, subset=day_columns), use_container_width=True)

            other_mask = ~res_df['라운지'].str.contains('|'.join(target_order), na=False)
            df_other = res_df[other_mask]
            if not df_other.empty:
                st.markdown(f"<h5 style='color: #D4AF37; margin-top: 20px; border-left: 4px solid #D4AF37; padding-left: 10px;'>기타 라운지</h5>", unsafe_allow_html=True)
                st.dataframe(df_other.style.map(color_schedule_cells, subset=day_columns), use_container_width=True)
                
            st.markdown("<br><hr>", unsafe_allow_html=True)
            
            st.markdown("### Ⅴ. 스케줄 검증 리포트 (Verification Report)")
            checklist = verify_schedule_checklist(all_emp, labels, s_year, s_month, data["male_off"], data["female_off"])
            chk_df = pd.DataFrame(checklist, columns=["점검 항목", "검증 기준", "점검 결과", "세부 보고 내용"])
            
            def highlight_result(val):
                if val == "PASS":
                    return 'color: #00FF00; font-weight: bold;'
                elif val == "FAIL":
                    return 'color: #FF4B4B; font-weight: bold;'
                elif val == "WARNING":
                    return 'color: #FFA500; font-weight: bold;'
                return ''
            
            st.dataframe(chk_df.style.map(highlight_result, subset=['점검 결과']), use_container_width=True)
            st.markdown("<br><hr>", unsafe_allow_html=True)

            st.markdown("### Ⅵ. 엑셀 다운로드 (Export to Excel)")
            lounge_schedules = {lounge: [] for lounge in LOUNGE_LIST}
            for i, e in enumerate(all_emp):
                if e["lounge"] in lounge_schedules:
                    lounge_schedules[e["lounge"]].append(labels[i])

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                export_to_excel_single_sheet(lounge_schedules, l_emp, all_emp, s_year, s_month, tmp_excel.name, data["public_holidays"], data["male_off"], data["female_off"])
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