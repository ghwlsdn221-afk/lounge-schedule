import calendar
import datetime
import os
import random

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from ortools.sat.python.cp_model import FEASIBLE, OPTIMAL, CpModel, CpSolver

# 7개 라운지 목록 (자데 최상단 배치, 바리스타 추가)
LOUNGE_LIST = ["자데", "블랙", "자홀", "블루", "세이지", "YP", "바리스타"]

LOUNGE_ALIAS = {
    "자스민 데스크": "자데",
    "자스민데스크": "자데",
    "자스민 홀": "자홀",
    "자스민홀": "자홀",
}

# 라운지별 근무 인원 기준 (최소, 최대)
LOUNGE_WORKER_BOUNDS = {
    "블랙": (3, 5),
    "자데": (2, 4),  
    "자홀": (6, 9),
    "블루": (2, 4),
    "세이지": (2, 3),
    "YP": (3, 5),
    "바리스타": (2, 4)  # 바리스타 최소 인원 세팅
}

def solve_global_schedule(
    emp_list, year, month, male_off_days, female_off_days, worker_bounds, public_holidays, store_closed_days=None,
    min_total_weekday=24, min_total_weekend=27
):
    if store_closed_days is None:
        store_closed_days = []

    num_days = calendar.monthrange(year, month)[1]
    num_emp = len(emp_list)

    if num_emp == 0:
        return None

    model = CpModel()
    penalty_vars = []

    # [변수 선언]
    is_off = {}
    assign = {}

    for e in range(num_emp):
        home_lounge = emp_list[e]["lounge"]
        for d in range(1, num_days + 1):
            is_off[e, d] = model.NewBoolVar(f"off_{e}_{d}")
            for l in LOUNGE_LIST:
                assign[e, d, l] = model.NewBoolVar(f"assign_{e}_{d}_{l}")

            model.Add(is_off[e, d] + sum(assign[e, d, l] for l in LOUNGE_LIST) == 1)

            # [지원(Support) 금지 룰 설정]
            if home_lounge == "YP":
                for l in LOUNGE_LIST:
                    if l != "YP":
                        model.Add(assign[e, d, l] == 0)
            else:
                model.Add(assign[e, d, "YP"] == 0)

            # ★ 바리스타 신규 룰 1: 타 라운지 소속은 바리스타 라운지로 지원 금지 (인바운드 원천 차단)
            if home_lounge != "바리스타":
                model.Add(assign[e, d, "바리스타"] == 0)

            # ★ 바리스타 신규 룰 2: 자데, 자홀, 바리스타 외에는 타 라운지 지원 불가 (아웃바운드 차단)
            # -> 즉, 바리스타는 타 라운지 지원(아웃바운드)이 '가능'하도록 허용됨
            if home_lounge not in ["자데", "자홀", "바리스타"] and home_lounge != "YP":
                for l in LOUNGE_LIST:
                    if l != home_lounge:
                        model.Add(assign[e, d, l] == 0)

    # 0. 휴점일 강제 전원 휴무
    for d in store_closed_days:
        for e in range(num_emp):
            model.Add(is_off[e, d] == 1)

    m_off_vars = {}

    for e, emp in enumerate(emp_list):
        target_off = male_off_days if emp["gender"] == "남" else female_off_days
        
        annual_leaves = emp.get("req_annual", [])
        valid_annual = [d for d in annual_leaves if d not in store_closed_days]
        
        # 1. 월간 목표 휴무일 + 연차일수 = 총 쉬는 날로 강제
        actual_off = sum(is_off[e, d] for d in range(1, num_days + 1))
        model.Add(actual_off == target_off + len(valid_annual))
        
        for ad in valid_annual:
            if 1 <= ad <= num_days:
                model.Add(is_off[e, ad] == 1)

        # 2. 신청 휴무일 반영
        for roff in emp["req_off"]:
            if 1 <= roff <= num_days and roff not in store_closed_days and roff not in valid_annual:
                penalty_vars.append((1 - is_off[e, roff]) * 3000)

        # 3. 생휴 제약
        if emp["gender"] == "여":
            valid_m_days = [
                d for d in range(1, num_days + 1)
                if datetime.date(year, month, d).weekday() < 4
                and d not in emp["req_off"]
                and d not in store_closed_days
                and d not in valid_annual
            ]

            day_m_vars = []
            for d in valid_m_days:
                is_m = model.NewBoolVar(f"is_m_{e}_{d}")
                model.Add(is_off[e, d] == 1).OnlyEnforceIf(is_m)
                day_m_vars.append((d, is_m))

                adj_off = []
                if d > 1: adj_off.append(is_off[e, d - 1])
                if d < num_days: adj_off.append(is_off[e, d + 1])
                
                if adj_off: model.AddBoolOr(adj_off).OnlyEnforceIf(is_m)
                else: model.Add(is_m == 0)

            if day_m_vars:
                model.AddExactlyOne([v[1] for v in day_m_vars])
                m_off_vars[e] = day_m_vars

                if emp["m_off"] and 1 <= emp["m_off"] <= num_days and emp["m_off"] in valid_m_days:
                    req_is_m = next((v[1] for v in day_m_vars if v[0] == emp["m_off"]), None)
                    if req_is_m is not None:
                        penalty_vars.append((1 - req_is_m) * 1000)

# 4. 연속 근무 및 휴무 패턴 제한
    for e in range(num_emp):
        # 1. 5연속 근무 절대 금지 (강력 차단 복구)
        for d in range(1, num_days - 3):
            model.Add(sum(is_off[e, d + offset] for offset in range(5)) >= 1)

        # 2. 4연속 근무 페널티 (가급적 3연속 이하로 짜도록 유도)
        for d in range(1, num_days - 2):
            worked_4 = model.NewBoolVar(f"w4_{e}_{d}")
            sum_off = sum(is_off[e, d + offset] for offset in range(4))
            model.Add(sum_off == 0).OnlyEnforceIf(worked_4)
            model.Add(sum_off >= 1).OnlyEnforceIf(worked_4.Not())
            penalty_vars.append(worked_4 * 100)

        # 3. 3연속 휴무 페널티
        for d in range(1, num_days - 1):
            off_3 = model.NewBoolVar(f"o3_{e}_{d}")
            sum_off = sum(is_off[e, d + offset] for offset in range(3))
            model.Add(sum_off == 3).OnlyEnforceIf(off_3)
            model.Add(sum_off <= 2).OnlyEnforceIf(off_3.Not())
            penalty_vars.append(off_3 * 300)
            
        # 👇 [신규 추가] 4. 이틀 연속 휴무 페널티 (휴무 분산 유도)
        # 생휴 등 필수적인 연속 휴무를 위해 벌점을 너무 높지 않게(200점) 설정하여,
        # 꼭 필요할 때만 2연속을 허용하고 가급적 1일 단위로 휴무를 쪼개도록 유도합니다.
        for d in range(1, num_days):
            off_2 = model.NewBoolVar(f"o2_{e}_{d}")
            sum_off = is_off[e, d] + is_off[e, d + 1]
            model.Add(sum_off == 2).OnlyEnforceIf(off_2)
            model.Add(sum_off <= 1).OnlyEnforceIf(off_2.Not())
            penalty_vars.append(off_2 * 200)

    # 5. 라운지별 일일 최소/최대 근무 인원 제약
    for d in range(1, num_days + 1):
        if d in store_closed_days: continue

        w_idx = datetime.date(year, month, d).weekday()
        is_weekend_or_holiday = (w_idx >= 5) or (d in public_holidays)

        for l in LOUNGE_LIST:
            working_workers = sum(assign[e, d, l] for e in range(num_emp))
            min_w, max_w = worker_bounds.get(l, (2, 3))

            shortfall = model.NewIntVar(0, num_emp, f"short_{d}_{l}")
            overage = model.NewIntVar(0, num_emp, f"over_{d}_{l}")

            if l == "자데":
                model.Add(working_workers + shortfall >= 2)
                model.Add(working_workers - overage <= 2)
                penalty_vars.append(shortfall * 10000)
                penalty_vars.append(overage * 2000)
                
            elif l == "YP":
                is_yp_busy = (w_idx >= 4) or (d in public_holidays)
                if is_yp_busy:
                    model.Add(working_workers + shortfall >= 4)
                else:
                    model.Add(working_workers + shortfall >= 3)
                model.Add(working_workers - overage <= 5)
                penalty_vars.append(shortfall * 10000)
                penalty_vars.append(overage * 2000)

            elif l == "바리스타": # ★ 바리스타 모든 요일 최소 2명 강제 룰
                model.Add(working_workers + shortfall >= 2)
                model.Add(working_workers - overage <= max_w)
                penalty_vars.append(shortfall * 10000)
                penalty_vars.append(overage * 1000)
                
            else:
                if is_weekend_or_holiday:
                    model.Add(working_workers + shortfall >= max_w)
                    model.Add(working_workers - overage <= max_w)
                    penalty_vars.append(shortfall * 10000)
                    penalty_vars.append(overage * 2000)
                else:
                    model.Add(working_workers + shortfall >= min_w)
                    model.Add(working_workers - overage <= max_w)
                    penalty_vars.append(shortfall * 10000)
                    penalty_vars.append(overage * 1000)

    # 6. 전사 매니저 최소 2명 출근 제약
    managers = [e for e, emp in enumerate(emp_list) if emp["rank"] == "매니저"]
    if len(managers) >= 2:
        for d in range(1, num_days + 1):
            if d in store_closed_days: continue
            mgrs_working = sum(1 - is_off[m, d] for m in managers)
            shortfall = model.NewIntVar(0, len(managers), f"mgr_short_{d}")
            model.Add(mgrs_working + shortfall >= 2)
            penalty_vars.append(shortfall * 1000)

    # 7. 라운지별 핵심 책임자 동시 휴무 금지
    for l in LOUNGE_LIST:
        if l == "바리스타": continue # ★ 바리스타 신규 룰 3: 책임자 교차 휴무 제외

        lounge_emps = [e for e, emp in enumerate(emp_list) if emp["lounge"] == l]
        if not lounge_emps: continue

        l_managers = [e for e in lounge_emps if emp_list[e]["rank"] == "매니저"]
        l_seniors = [e for e in lounge_emps if emp_list[e]["rank"] == "선임"]

        leaders = []
        if len(l_managers) >= 2: leaders = l_managers[:2]
        elif len(l_managers) == 1 and len(l_seniors) >= 1: leaders = [l_managers[0], l_seniors[0]]
        elif len(l_managers) == 0 and len(l_seniors) >= 2: leaders = l_seniors[:2]

        if leaders:
            for d in range(1, num_days + 1):
                if d in store_closed_days: continue
                model.Add(sum(is_off[ldr, d] for ldr in leaders) <= 1)

        for d in range(1, num_days + 1):
            if d in store_closed_days: continue
            model.Add(sum(is_off[e, d] for e in lounge_emps) <= len(lounge_emps) - 1)

    # 8. 타 라운지 지원 근무 벌점
    for e in range(num_emp):
        home_lounge = emp_list[e]["lounge"]
        for d in range(1, num_days + 1):
            if d in store_closed_days: continue
            for l in LOUNGE_LIST:
                if l != home_lounge:
                    penalty_vars.append(assign[e, d, l] * 100) # 바리스타가 타 라운지 갈 때도 페널티 부여(가급적 본인 라운지 지키도록 유도)

    # 9. 전사 일일 총 출근 인원 제약 (평일 / 주말·공휴일)
    for d in range(1, num_days + 1):
        if d in store_closed_days: continue
        
        w_idx = datetime.date(year, month, d).weekday()
        is_weekend_or_holiday = (w_idx >= 4) or (d in public_holidays)
        
        total_working_day = sum(assign[e, d, l] for e in range(num_emp) for l in LOUNGE_LIST)
        shortfall = model.NewIntVar(0, num_emp, f"global_short_{d}")
        
        if is_weekend_or_holiday:
            model.Add(total_working_day + shortfall >= min_total_weekend)
        else:
            model.Add(total_working_day + shortfall >= min_total_weekday)
            
        penalty_vars.append(shortfall * 50000)

    model.Minimize(sum(penalty_vars))
    solver = CpSolver()
    solver.parameters.random_seed = random.randint(1, 10000)
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status not in (OPTIMAL, FEASIBLE):
        return None

    flat_labels = []
    for e in range(num_emp):
        emp = emp_list[e]
        chosen_m_day = None
        if emp["gender"] == "여" and e in m_off_vars:
            for d, is_m in m_off_vars[e]:
                if solver.Value(is_m) == 1:
                    chosen_m_day = d
                    break

        valid_annual = [d for d in emp.get("req_annual", []) if d not in store_closed_days]

        emp_labels = []
        for d in range(1, num_days + 1):
            if solver.Value(is_off[e, d]) == 1:
                if d in store_closed_days:
                    emp_labels.append("휴점")
                elif d in valid_annual:
                    emp_labels.append("연차")
                elif d in emp["req_off"]:
                    emp_labels.append("신청휴")
                elif d == chosen_m_day:
                    emp_labels.append("생휴")
                else:
                    emp_labels.append("휴무")
            else:
                emp_labels.append("근무")

        flat_labels.append(emp_labels)

    return flat_labels

def verify_schedule_checklist(emp_list, flat_labels, year, month, male_off_days, female_off_days, store_closed_days=None, min_total_weekday=24, min_total_weekend=27, public_holidays=None):
    if store_closed_days is None: store_closed_days = []
    if public_holidays is None: public_holidays = []
    
    num_days = calendar.monthrange(year, month)[1]
    checklist = []
    off_keywords = ["휴무", "신청휴", "생휴", "휴점", "연차"]

    # 1. 매니저 최소 출근
    managers = [e for e, emp in enumerate(emp_list) if emp["rank"] == "매니저"]
    if len(managers) >= 2:
        mgr_short_days = []
        for d in range(1, num_days + 1):
            if d in store_closed_days: continue
            working_cnt = sum(1 for m in managers if flat_labels[m][d - 1] == "근무")
            if working_cnt < 2: mgr_short_days.append(d)
        if not mgr_short_days: chk1 = ("매니저 최소 출근", "전사 매니저 매일 2명 이상 출근", "✅ 적합", "전일 매니저 2명 이상 출근 달성")
        else: chk1 = ("매니저 최소 출근", "전사 매니저 매일 2명 이상 출근", "❌ 미흡", f"미달 일자: {mgr_short_days}일")
    else: chk1 = ("매니저 최소 출근", "전사 매니저 매일 2명 이상 출근", "ℹ️ 해당없음", f"전사 매니저 총 {len(managers)}명")
    checklist.append(chk1)

    # 2. 책임자 휴무 교차 여부
    overlap_list = []
    for l in LOUNGE_LIST:
        if l == "바리스타": continue # ★ 체크리스트에서도 바리스타 교차 검증 패스

        lounge_emps = [e for e, emp in enumerate(emp_list) if emp["lounge"] == l]
        if not lounge_emps: continue
        l_managers = [e for e in lounge_emps if emp_list[e]["rank"] == "매니저"]
        l_seniors = [e for e in lounge_emps if emp_list[e]["rank"] == "선임"]
        leaders = []
        if len(l_managers) >= 2: leaders = l_managers[:2]
        elif len(l_managers) == 1 and len(l_seniors) >= 1: leaders = [l_managers[0], l_seniors[0]]
        elif len(l_managers) == 0 and len(l_seniors) >= 2: leaders = l_seniors[:2]

        if leaders:
            for d in range(1, num_days + 1):
                if d in store_closed_days: continue
                off_cnt = sum(1 for ldr in leaders if flat_labels[ldr][d - 1] in off_keywords)
                if off_cnt > 1: overlap_list.append(f"{l}({d}일)")

    if not overlap_list: chk2 = ("책임자 휴무 교차", "라운지별 책임자(매니저/선임) 동시 휴무 불가", "✅ 적합", "전 라운지 책임자 동시 휴무 미발생")
    else: chk2 = ("책임자 휴무 교차", "라운지별 책임자(매니저/선임) 동시 휴무 불가", "❌ 미흡", f"중첩 발생: {', '.join(overlap_list)}")
    checklist.append(chk2)

    # 3. 생휴 제약
    m_issues = []
    for e, emp in enumerate(emp_list):
        if emp["gender"] == "여":
            for d in range(1, num_days + 1):
                if flat_labels[e][d - 1] == "생휴":
                    w_idx = datetime.date(year, month, d).weekday()
                    is_mon_thu = (w_idx < 4)
                    has_adj_off = False
                    if d > 1 and flat_labels[e][d - 2] in off_keywords: has_adj_off = True
                    if d < num_days and flat_labels[e][d] in off_keywords: has_adj_off = True
                    if not (is_mon_thu and has_adj_off):
                        m_issues.append(f"{emp['name']}({d}일)")

    if not m_issues: chk3 = ("생리휴가 규정 준수", "월~목 배정 및 이틀 연속 휴무 반드시 포함", "✅ 적합", "전원 생휴 연속 휴무 조건 충족")
    else: chk3 = ("생리휴가 규정 준수", "월~목 배정 및 이틀 연속 휴무 반드시 포함", "❌ 미흡", f"미충족: {', '.join(m_issues)}")
    checklist.append(chk3)

    # 4. 4연속 근무
    consecutive4_list = []
    for e, emp in enumerate(emp_list):
        emp_labels = flat_labels[e]
        w_count = 0
        ranges = []
        for d in range(1, num_days + 1):
            if emp_labels[d - 1] == "근무": w_count += 1
            else:
                if w_count >= 4: ranges.append(f"{d - w_count}~{d - 1}일({w_count}연속)")
                w_count = 0
        if w_count >= 4: ranges.append(f"{num_days - w_count + 1}~{num_days}일({w_count}연속)")
        if ranges: consecutive4_list.append(f"{emp['name']}[{emp['lounge']}]({', '.join(ranges)})")

    if not consecutive4_list: chk4 = ("4연속 근무 점검", "최대한 3연속 이하 근무 유도 (4연속 기피)", "✅ 미발생", "4연속 이상 근무자 없음")
    else: chk4 = ("4연속 근무 점검", "최대한 3연속 이하 근무 유도 (4연속 기피)", "⚠️ 발생", f"발생 인원: {', '.join(consecutive4_list)}")
    checklist.append(chk4)

    # 5. 목표 휴무일수
    off_mismatch = []
    for e, emp in enumerate(emp_list):
        target_off = male_off_days if emp["gender"] == "남" else female_off_days
        actual_regular_off_cnt = sum(1 for d in range(1, num_days + 1) if flat_labels[e][d - 1] in ["휴무", "신청휴", "생휴", "휴점"])
        if actual_regular_off_cnt != target_off:
            off_mismatch.append(f"{emp['name']}(기본휴무 {actual_regular_off_cnt}일)")

    if not off_mismatch: chk5 = ("월 목표 휴무일수", f"남성 {male_off_days}일 / 여성 {female_off_days}일 정확히 준수 (연차 별도)", "✅ 적합", "전 직원 연차 제외 기본 휴무일수 100% 달성")
    else: chk5 = ("월 목표 휴무일수", f"남성 {male_off_days}일 / 여성 {female_off_days}일 정확히 준수 (연차 별도)", "❌ 미흡", f"일수 불일치: {', '.join(off_mismatch)}")
    checklist.append(chk5)

    # 6. 전사 총 출근 인원 점검
    global_short_days = []
    for d in range(1, num_days + 1):
        if d in store_closed_days: continue
        w_idx = datetime.date(year, month, d).weekday()
        is_weekend_or_holiday = (w_idx >= 4) or (d in public_holidays)
        
        working_cnt = sum(1 for e in range(len(emp_list)) if flat_labels[e][d - 1] == "근무")
        target = min_total_weekend if is_weekend_or_holiday else min_total_weekday
        
        if working_cnt < target:
            global_short_days.append(f"{d}일({working_cnt}명/최소{target}명)")

    if not global_short_days: 
        chk6 = ("전사 총 출근 인원", f"평일 {min_total_weekday}명 / 주말·공휴일 {min_total_weekend}명 이상", "✅ 적합", "모든 영업일 최소 출근 인원 충족")
    else: 
        chk6 = ("전사 총 출근 인원", f"평일 {min_total_weekday}명 / 주말·공휴일 {min_total_weekend}명 이상", "❌ 미흡", f"미달 일자: {', '.join(global_short_days)}")
    checklist.append(chk6)

    return checklist

def export_to_excel_single_sheet(
    lounge_schedules, lounge_employees, all_employees_flat, year, month, output_excel,
    public_holidays, male_off_days, female_off_days, store_closed_days=None
):
    if store_closed_days is None: store_closed_days = []
    num_days = calendar.monthrange(year, month)[1]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "월간근무표(통합)"
    ws.sheet_view.showGridLines = True

    font_title = Font(name="맑은 고딕", size=16, bold=True, color="1F4E78")
    font_section = Font(name="맑은 고딕", size=13, bold=True, color="1F4E78")
    font_hdr = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_hdr_red = Font(name="맑은 고딕", size=10, bold=True, color="FF0000")

    fill_navy = PatternFill(start_color="1F4E78", fill_type="solid")
    fill_sat = PatternFill(start_color="D9E1F2", fill_type="solid")
    fill_sun_hol = PatternFill(start_color="FCE4D6", fill_type="solid")
    fill_att = PatternFill(start_color="D9E1F2", fill_type="solid")
    font_att = Font(name="맑은 고딕", size=9, bold=True, color="1F4E78")
    fill_sup_row = PatternFill(start_color="FCE4D6", fill_type="solid")
    font_sup_row = Font(name="맑은 고딕", size=9, bold=True, color="C65911")
    fill_act = PatternFill(start_color="E2EFDA", fill_type="solid")
    font_act = Font(name="맑은 고딕", size=10, bold=True, color="276A3C")
    border_thin = Side(border_style="thin", color="BFBFBF")
    border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    fill_map_excel = {
        "근무": PatternFill(start_color="E3F2FD", fill_type="solid"),
        "휴무": PatternFill(start_color="E0E0E0", fill_type="solid"),
        "신청휴": PatternFill(start_color="C8E6C9", fill_type="solid"),
        "생휴": PatternFill(start_color="FFE0B2", fill_type="solid"),
        "연차": PatternFill(start_color="E1BEE7", fill_type="solid"), 
        "휴점": PatternFill(start_color="9E9E9E", fill_type="solid"),
    }
    font_map_excel = {
        "근무": Font(name="맑은 고딕", size=9, bold=True, color="0D47A1"),
        "휴무": Font(name="맑은 고딕", size=9, color="424242"),
        "신청휴": Font(name="맑은 고딕", size=9, bold=True, color="1B5E20"),
        "생휴": Font(name="맑은 고딕", size=9, bold=True, color="E65100"),
        "연차": Font(name="맑은 고딕", size=9, bold=True, color="4A148C"),
        "휴점": Font(name="맑은 고딕", size=9, bold=True, color="FFFFFF"),
    }

    weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]

    ws.cell(row=1, column=1, value=f"📅 {year}년 {month}월 전 라운지 통합 근무표").font = font_title
    curr_r = 3

    for lounge_name in LOUNGE_LIST:
        if lounge_name not in lounge_schedules or not lounge_employees[lounge_name]: continue

        ws.cell(row=curr_r, column=1, value=f"■ {lounge_name} 라운지").font = font_section
        curr_r += 1

        headers = (
            ["라운지", "직급", "성명", "성별"]
            + [f"{d}일" for d in range(1, num_days + 1)]
            + ["근무", "휴무", "신청휴", "생휴", "연차", "휴점", "휴무총합"]
        )
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=curr_r, column=i, value=h)
            c.font = font_hdr
            c.fill = fill_navy
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_box
        curr_r += 1

        for d in range(1, num_days + 1):
            col_idx = d + 4
            w_idx = datetime.date(year, month, d).weekday()
            is_holiday = d in public_holidays or d in store_closed_days

            c_day = ws.cell(row=curr_r, column=col_idx, value=weekdays_kr[w_idx])
            c_day.font = font_hdr_red if (w_idx >= 5 or is_holiday) else font_hdr

            if w_idx == 5 and not is_holiday: c_day.fill = fill_sat
            elif w_idx == 6 or is_holiday: c_day.fill = fill_sun_hol
            else: c_day.fill = fill_navy
            c_day.alignment = Alignment(horizontal="center", vertical="center")
            c_day.border = border_box

        for c in list(range(1, 5)) + list(range(num_days + 5, num_days + 12)):
            cell = ws.cell(row=curr_r, column=c)
            cell.fill = fill_navy
            cell.border = border_box
        curr_r += 1

        emp_list = lounge_employees[lounge_name]
        labels_list = lounge_schedules[lounge_name]
        start_emp_r = curr_r

        for e, emp in enumerate(emp_list):
            ws.cell(row=curr_r, column=1, value=emp["lounge"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_r, column=2, value=emp["rank"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_r, column=3, value=emp["name"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_r, column=4, value=emp["gender"]).alignment = Alignment(horizontal="center")
            for i in range(1, 5): ws.cell(row=curr_r, column=i).border = border_box

            for d in range(1, num_days + 1):
                lbl = labels_list[e][d - 1]
                cell = ws.cell(row=curr_r, column=d + 4, value=lbl)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = fill_map_excel.get(lbl, fill_map_excel["근무"])
                cell.font = font_map_excel.get(lbl, font_map_excel["근무"])
                cell.border = border_box

            sum_col_start = num_days + 5
            start_letter = get_column_letter(5)
            end_letter = get_column_letter(num_days + 4)

            for i, h in enumerate(["근무", "휴무", "신청휴", "생휴", "연차", "휴점"]):
                c_sum = ws.cell(
                    row=curr_r,
                    column=sum_col_start + i,
                    value=f'=COUNTIF({start_letter}{curr_r}:{end_letter}{curr_r}, "{h}")',
                )
                c_sum.alignment = Alignment(horizontal="center")
                c_sum.border = border_box

            # 휴무총합 (휴무, 신청휴, 생휴, 휴점만 합산 / 연차 제외)
            col_off_start = get_column_letter(sum_col_start + 1) # 휴무
            col_off_end = get_column_letter(sum_col_start + 3)   # 생휴
            col_closed = get_column_letter(sum_col_start + 5)    # 휴점
            
            c_tot_off = ws.cell(
                row=curr_r,
                column=sum_col_start + 6,
                value=f"=SUM({col_off_start}{curr_r}:{col_off_end}{curr_r}) + {col_closed}{curr_r}",
            )
            c_tot_off.alignment = Alignment(horizontal="center")
            c_tot_off.border = border_box

            curr_r += 1

        last_emp_r = curr_r - 1
        r_att, r_sup, r_act = curr_r, curr_r + 1, curr_r + 2

        for r_idx, title, fill, font in [(r_att, "출근인원", fill_att, font_att), 
                                         (r_sup, "타 접점 지원", fill_sup_row, font_sup_row), 
                                         (r_act, "실제 근무인원", fill_act, font_act)]:
            for c in range(1, 5):
                cell = ws.cell(row=r_idx, column=c)
                cell.fill, cell.font, cell.border = fill, font, border_box
            ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=4)
            ws.cell(row=r_idx, column=1, value=title).alignment = Alignment(horizontal="center", vertical="center")

        for d in range(1, num_days + 1):
            col_idx = d + 4
            col_let = get_column_letter(col_idx)

            c_att = ws.cell(row=r_att, column=col_idx, value=f'=COUNTIF({col_let}{start_emp_r}:{col_let}{last_emp_r}, "근무")')
            c_att.fill, c_att.font, c_att.border = fill_att, font_att, border_box
            c_att.alignment = Alignment(horizontal="center", vertical="center")

            c_sup = ws.cell(row=r_sup, column=col_idx, value="")
            c_sup.fill, c_sup.font, c_sup.border = fill_sup_row, font_sup_row, border_box
            c_sup.alignment = Alignment(horizontal="center", vertical="center")

            c_act = ws.cell(row=r_act, column=col_idx, value=f'={col_let}{r_att} + IF({col_let}{r_sup}="", 0, {col_let}{r_sup})')
            c_act.fill, c_act.font, c_act.border = fill_act, font_act, border_box
            c_act.alignment = Alignment(horizontal="center", vertical="center")

        curr_r = r_act + 3

    ws.cell(row=curr_r, column=1, value="📋 스케줄 최적화 연산 결과 검증 체크리스트").font = font_section
    curr_r += 1

    chk_headers = ["점검 항목", "검증 기준", "점검 결과", "세부 보고 및 미달 내용"]
    chk_col_spans = [(1, 2), (3, 4), (5, 6), (7, num_days + 12)]

    for idx, h_text in enumerate(chk_headers):
        s_col, e_col = chk_col_spans[idx]
        for c in range(s_col, e_col + 1):
            cell = ws.cell(row=curr_r, column=c)
            cell.fill, cell.font, cell.border = fill_navy, font_hdr, border_box
        ws.merge_cells(start_row=curr_r, start_column=s_col, end_row=curr_r, end_column=e_col)
        ws.cell(row=curr_r, column=s_col, value=h_text).alignment = Alignment(horizontal="center", vertical="center")
    curr_r += 1

    flat_labels_all = []
    for emp in all_employees_flat:
        l = emp["lounge"]
        e_idx = next(i for i, e in enumerate(lounge_employees[l]) if e["name"] == emp["name"])
        flat_labels_all.append(lounge_schedules[l][e_idx])

    # 💡 엑셀 출력 함수 내부에서 체크리스트 호출할 때도 24, 27 강제 전달 대신, public_holidays 포함하여 넘김
    checklist_results = verify_schedule_checklist(
        all_employees_flat, flat_labels_all, year, month, male_off_days, female_off_days, 
        store_closed_days, 24, 27, public_holidays
    )

    font_pass, fill_pass = Font(name="맑은 고딕", size=10, bold=True, color="1B5E20"), PatternFill(start_color="C8E6C9", fill_type="solid")
    font_fail, fill_fail = Font(name="맑은 고딕", size=10, bold=True, color="C62828"), PatternFill(start_color="FFCDD2", fill_type="solid")

    for item_name, criteria, status_str, detail_str in checklist_results:
        is_pass = "적합" in status_str or "미발생" in status_str
        res_font, res_fill = (font_pass, fill_pass) if is_pass else (font_fail, fill_fail)

        for s_col, e_col, val, al, bg, ft in [
            (1, 2, item_name, "center", None, None),
            (3, 4, criteria, "left", None, None),
            (5, 6, status_str, "center", res_fill, res_font),
            (7, num_days + 12, detail_str, "left", None, None)
        ]:
            for c in range(s_col, e_col + 1):
                cell = ws.cell(row=curr_r, column=c)
                cell.border = border_box
                if bg: cell.fill = bg
                if ft: cell.font = ft
            ws.merge_cells(start_row=curr_r, start_column=s_col, end_row=curr_r, end_column=e_col)
            ws.cell(row=curr_r, column=s_col, value=val).alignment = Alignment(horizontal=al, vertical="center")
        curr_r += 1

    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 10, 8
    ws.column_dimensions["C"].width, ws.column_dimensions["D"].width = 10, 6
    for d in range(1, num_days + 1):
        ws.column_dimensions[get_column_letter(d + 4)].width = 7

    wb.save(output_excel)