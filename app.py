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

# 웹 페이지 기본 설정 (넓은 화면 모드)
st.set_page_config(page_title="라운지 근무표 생성기", layout="wide")

# 헤더 제목
st.title("📅 전 라운지 통합 근무 스케줄 생성기")
st.write("파이썬 OR-Tools 최적화 엔진으로 100% 조건에 맞는 스케줄을 자동 생성합니다.")

# ---------------------------------------------------------
# [사이드바] 연산 설정 영역
# ---------------------------------------------------------
st.sidebar.header("⚙️ 기본 연산 설정")
year = st.sidebar.number_input("연도", value=2026, step=1)
month = st.sidebar.number_input("월", value=10, min_value=1, max_value=12, step=1)
male_off = st.sidebar.number_input("남성 목표 휴무일수", value=11, step=1)
female_off = st.sidebar.number_input("여성 목표 휴무일수", value=12, step=1)
holidays_str = st.sidebar.text_input("공휴일 지정 (쉼표 구분)", value="3, 9")

# 공휴일 텍스트를 리스트로 변환
public_holidays = [int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()]

# ---------------------------------------------------------
# [메인 화면] 파일 업로드 및 스케줄 생성
# ---------------------------------------------------------
uploaded_file = st.file_uploader("📂 직원 CSV 파일 업로드", type=["csv"])

if uploaded_file is not None:
    # 1. 업로드된 CSV 데이터 미리보기
    try:
        df_input = pd.read_csv(uploaded_file, encoding="utf-8-sig")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df_input = pd.read_csv(uploaded_file, encoding="cp949")

    st.subheader("👥 불러온 직원 목록")
    st.dataframe(df_input, use_container_width=True)

    # 2. 스케줄 생성 버튼
    if st.button("⚡ 스케줄 자동 생성 (OR-Tools 최적화)", type="primary"):
        with st.spinner("최적화 엔진 연산 중... 잠시만 기다려주세요."):
            
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
                st.error("❌ 스케줄 생성 실패: 조건에 맞는 스케줄 조합을 찾을 수 없습니다.")
            else:
                st.success("🎉 스케줄 생성이 완료되었습니다!")

                # 3. 검증 체크리스트 표시
                st.subheader("📋 스케줄 최적화 결과 검증 체크리스트")
                checklist = verify_schedule_checklist(
                    all_employees_flat, flat_labels, year, month, male_off, female_off
                )
                
                chk_df = pd.DataFrame(checklist, columns=["점검 항목", "검증 기준", "점검 결과", "세부 보고 내용"])
                st.dataframe(chk_df, use_container_width=True)

                # 4. 엑셀 파일 다운로드 제공
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
                            label="📊 생성된 엑셀 파일 다운로드 (.xlsx)",
                            data=f.read(),
                            file_name=f"월간근무표_{year}년_{month}월.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )