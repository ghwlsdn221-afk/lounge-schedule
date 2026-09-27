import calendar
import datetime
import os
import tempfile
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

from ortools.sat.python.cp_model import FEASIBLE, OPTIMAL, CpModel, CpSolver
import streamlit as st

# ==============================================================================
# 1. 상수 및 라운지 설정
# ==============================================================================
LOUNGE_LIST = ["자데", "블랙", "자홀", "블루", "세이지", "YP"]

LOUNGE_ALIAS = {
    "자스민 데스크": "자데",
    "자스민데스크": "자데",
    "자스민 홀": "자홀",
    "자스민홀": "자홀",
}

LOUNGE_WORKER_BOUNDS = {
    "블랙": (3, 4),
    "자데": (2, 3),
    "자홀": (6, 8),
    "블루": (2, 3),
    "세이지": (2, 3),
}

SAMPLE_CSV_DATA = """이름,성별,신청휴무일,라운지,생리휴가일,직급
김연진M,여,"2, 16",자데,,매니저
백은정,여,"4, 18",자데,,선임
이지민,여,"6, 20",자데,,사원
권태검,남,"8, 22",자데,,사원
김연재M,여,"10, 24",자홀,,매니저
이시은,여,"3, 17",자홀,,선임
박유빈,여,"8, 22",자홀,,사원
임채은,여,"5, 19",자홀,,사원
김시연,여,"12, 26",자홀,,사원
허승규,남,"1, 15",자홀,,사원
김미나,여,,자홀,,사원
이연주,여,,자홀,,사원
정현영,남,,자홀,,사원
조예원,여,,자홀,,사원
하석영,남,,자홀,,사원
정현주M,여,,블랙,,매니저
김유정M,여,,블랙,,매니저
이수연,여,,블랙,,사원
김유진,여,,블랙,,사원
조유림,여,,블랙,,사원
이예지,여,,블랙,,사원
임은지,여,,블루,,선임
김서연,여,,블루,,선임
김해인,여,,블루,,사원
이식우,남,,블루,,사원
오유경,여,,세이지,,선임
서은서,여,,세이지,,선임
박정은,여,,세이지,,사원"""


# ==============================================================================
# 2. 구글 OR-Tools 스케줄 최적화 연산 엔진
# ==============================================================================
def solve_global_schedule(
    emp_list,
    year,
    month,
    male_off_days,
    female_off_days,
    worker_bounds,
    public_holidays,
):
    num_days = calendar.monthrange(year, month)[1]
    num_emp = len(emp_list)

    if num_emp == 0:
        return None

    model = CpModel()
    penalty_vars = []

    is_off = {}
    assign = {}

    for e in range(num_emp):
        home_lounge = emp_list[e]["lounge"]
        for d in range(1, num_days + 1):
            is_off[e, d] = model.NewBoolVar(f"off_{e}_{d}")
            for l in LOUNGE_LIST:
                assign[e, d, l] = model.NewBoolVar(f"assign_{e}_{d}_{l}")

            model.Add(
                is_off[e, d] + sum(assign[e, d, l] for l in LOUNGE_LIST) == 1
            )

            if home_lounge == "YP":
                for l in LOUNGE_LIST:
                    if l != "YP":
                        model.Add(assign[e, d, l] == 0)
            else:
                model.Add(assign[e, d, "YP"] == 0)

            if home_lounge not in ["자데", "자홀"] and home_lounge != "YP":
                for l in LOUNGE_LIST:
                    if l != home_lounge:
                        model.Add(assign[e, d, l] == 0)

    m_off_vars = {}

    for e, emp in enumerate(emp_list):
        target_off = male_off_days if emp["gender"] == "남" else female_off_days
        actual_off = sum(is_off[e, d] for d in range(1, num_days + 1))
        model.Add(actual_off == target_off)

        for roff in emp["req_off"]:
            if 1 <= roff <= num_days:
                penalty_vars.append((1 - is_off[e, roff]) * 3000)

        if emp["gender"] == "여":
            valid_m_days = [
                d
                for d in range(1, num_days + 1)
                if datetime.date(year, month, d).weekday() < 4
                and d not in emp["req_off"]
            ]

            day_m_vars = []
            for d in valid_m_days:
                is_m = model.NewBoolVar(f"is_m_{e}_{d}")
                model.Add(is_off[e, d] == 1).OnlyEnforceIf(is_m)
                day_m_vars.append((d, is_m))

            if day_m_vars:
                model.AddExactlyOne([v[1] for v in day_m_vars])
                m_off_vars[e] = day_m_vars

                if (
                    emp["m_off"]
                    and 1 <= emp["m_off"] <= num_days
                    and emp["m_off"] in valid_m_days
                ):
                    req_is_m = next(
                        (v[1] for v in day_m_vars if v[0] == emp["m_off"]), None
                    )
                    if req_is_m is not None:
                        penalty_vars.append((1 - req_is_m) * 1000)

                for d, is_m in day_m_vars:
                    adj_off = []
                    if d > 1:
                        adj_off.append(is_off[e, d - 1])
                    if d < num_days:
                        adj_off.append(is_off[e, d + 1])

                    if adj_off:
                        sum_adj = sum(adj_off)
                        iso = model.NewBoolVar(f"m_iso_{e}_{d}")
                        model.Add(sum_adj == 0).OnlyEnforceIf(iso)
                        model.Add(sum_adj >= 1).OnlyEnforceIf(iso.Not())

                        b_and = model.NewBoolVar(f"m_and_{e}_{d}")
                        model.AddBoolOr([is_m.Not(), iso.Not(), b_and])
                        model.AddImplication(b_and, is_m)
                        model.AddImplication(b_and, iso)

                        penalty_vars.append(b_and * 500)

    for e in range(num_emp):
        for d in range(1, num_days - 3):
            model.Add(sum(is_off[e, d + offset] for offset in range(5)) >= 1)

        for d in range(1, num_days - 2):
            worked_4 = model.NewBoolVar(f"w4_{e}_{d}")
            sum_off = sum(is_off[e, d + offset] for offset in range(4))
            model.Add(sum_off == 0).OnlyEnforceIf(worked_4)
            model.Add(sum_off >= 1).OnlyEnforceIf(worked_4.Not())
            penalty_vars.append(worked_4 * 800)

        for d in range(1, num_days - 1):
            off_3 = model.NewBoolVar(f"o3_{e}_{d}")
            sum_off = sum(is_off[e, d + offset] for offset in range(3))
            model.Add(sum_off == 3).OnlyEnforceIf(off_3)
            model.Add(sum_off <= 2).OnlyEnforceIf(off_3.Not())
            penalty_vars.append(off_3 * 300)

    for d in range(1, num_days + 1):
        is_weekend_or_holiday = (
            datetime.date(year, month, d).weekday() >= 5
        ) or (d in public_holidays)

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

    managers = [e for e, emp in enumerate(emp_list) if emp["rank"] == "매니저"]
    if len(managers) >= 2:
        for d in range(1, num_days + 1):
            mgrs_working = sum(1 - is_off[m, d] for m in managers)
            shortfall = model.NewIntVar(0, len(managers), f"mgr_short_{d}")
            model.Add(mgrs_working + shortfall >= 2)
            penalty_vars.append(shortfall * 1000)

    for l in LOUNGE_LIST:
        lounge_emps = [
            e for e, emp in enumerate(emp_list) if emp["lounge"] == l
        ]
        if not lounge_emps:
            continue

        l_managers = [
            e for e in lounge_emps if emp_list[e]["rank"] == "매니저"
        ]
        l_seniors = [e for e in lounge_emps if emp_list[e]["rank"] == "선임"]

        leaders = []
        if len(l_managers) >= 2:
            leaders = l_managers[:2]
        elif len(l_managers) == 1 and len(l_seniors) >= 1:
            leaders = [l_managers[0], l_seniors[0]]
        elif len(l_managers) == 0 and len(l_seniors) >= 2:
            leaders = l_seniors[:2]

        if leaders:
            for d in range(1, num_days + 1):
                model.Add(sum(is_off[ldr, d] for ldr in leaders) <= 1)

        for d in range(1, num_days + 1):
            model.Add(
                sum(is_off[e, d] for e in lounge_emps) <= len(lounge_emps) - 1
            )

    for e in range(num_emp):
        home_lounge = emp_list[e]["lounge"]
        for d in range(1, num_days + 1):
            for l in LOUNGE_LIST:
                if l != home_lounge:
                    penalty_vars.append(assign[e, d, l] * 100)

    model.Minimize(sum(penalty_vars))

    solver = CpSolver()
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

        emp_labels = []
        for d in range(1, num_days + 1):
            if solver.Value(is_off[e, d]) == 1:
                if d in emp["req_off"]:
                    emp_labels.append("신청휴")
                elif d == chosen_m_day:
                    emp_labels.append("생휴")
                else:
                    emp_labels.append("휴무")
            else:
                assigned_l = emp["lounge"]
                for l in LOUNGE_LIST:
                    if solver.Value(assign[e, d, l]) == 1:
                        assigned_l = l
                        break
                if assigned_l != emp["lounge"]:
                    emp_labels.append(f"지원({assigned_l})")
                else:
                    emp_labels.append("근무")

        flat_labels.append(emp_labels)

    return flat_labels


# ==============================================================================
# 3. 체크리스트 검증 함수
# ==============================================================================
def verify_schedule_checklist(
    emp_list, flat_labels, year, month, male_off_days, female_off_days
):
    num_days = calendar.monthrange(year, month)[1]
    checklist = []

    managers = [e for e, emp in enumerate(emp_list) if emp["rank"] == "매니저"]
    if len(managers) >= 2:
        mgr_short_days = []
        for d in range(1, num_days + 1):
            working_cnt = sum(
                1
                for m in managers
                if "근무" in flat_labels[m][d - 1]
                or "지원" in flat_labels[m][d - 1]
            )
            if working_cnt < 2:
                mgr_short_days.append(d)

        if not mgr_short_days:
            chk1 = (
                "매니저 최소 출근",
                "전사 매니저 매일 2명 이상 출근",
                "✅ 적합",
                "전일 매니저 2명 이상 출근 달성",
            )
        else:
            chk1 = (
                "매니저 최소 출근",
                "전사 매니저 매일 2명 이상 출근",
                "❌ 미흡",
                f"미달 일자: {mgr_short_days}일",
            )
    else:
        chk1 = (
            "매니저 최소 출근",
            "전사 매니저 매일 2명 이상 출근",
            "ℹ️ 해당없음",
            f"전사 매니저 총 {len(managers)}명",
        )
    checklist.append(chk1)

    overlap_list = []
    for l in LOUNGE_LIST:
        lounge_emps = [
            e for e, emp in enumerate(emp_list) if emp["lounge"] == l
        ]
        if not lounge_emps:
            continue
        l_managers = [
            e for e in lounge_emps if emp_list[e]["rank"] == "매니저"
        ]
        l_seniors = [e for e in lounge_emps if emp_list[e]["rank"] == "선임"]

        leaders = []
        if len(l_managers) >= 2:
            leaders = l_managers[:2]
        elif len(l_managers) == 1 and len(l_seniors) >= 1:
            leaders = [l_managers[0], l_seniors[0]]
        elif len(l_managers) == 0 and len(l_seniors) >= 2:
            leaders = l_seniors[:2]

        if leaders:
            for d in range(1, num_days + 1):
                off_cnt = sum(
                    1
                    for ldr in leaders
                    if flat_labels[ldr][d - 1] in ["휴무", "신청휴", "생휴"]
                )
                if off_cnt > 1:
                    overlap_list.append(f"{l}({d}일)")

    if not overlap_list:
        chk2 = (
            "책임자 휴무 교차",
            "라운지별 책임자(매니저/선임) 동시 휴무 불가",
            "✅ 적합",
            "전 라운지 책임자 동시 휴무 미발생",
        )
    else:
        chk2 = (
            "책임자 휴무 교차",
            "라운지별 책임자(매니저/선임) 동시 휴무 불가",
            "❌ 미흡",
            f"중첩 발생: {', '.join(overlap_list)}",
        )
    checklist.append(chk2)

    m_issues = []
    for e, emp in enumerate(emp_list):
        if emp["gender"] == "여":
            for d in range(1, num_days + 1):
                if flat_labels[e][d - 1] == "생휴":
                    w_idx = datetime.date(year, month, d).weekday()
                    is_mon_thu = w_idx < 4

                    has_adj_off = False
                    if d > 1 and flat_labels[e][d - 2] in [
                        "휴무",
                        "신청휴",
                        "생휴",
                    ]:
                        has_adj_off = True
                    if d < num_days and flat_labels[e][d] in [
                        "휴무",
                        "신청휴",
                        "생휴",
                    ]:
                        has_adj_off = True

                    if not (is_mon_thu and has_adj_off):
                        m_issues.append(f"{emp['name']}({d}일)")

    if not m_issues:
        chk3 = (
            "생리휴가 규정 준수",
            "월~목요일 배정 및 이틀 연속 휴무 보장",
            "✅ 적합",
            "전원 월~목 배정 및 연속 휴무 조건 충족",
        )
    else:
        chk3 = (
            "생리휴가 규정 준수",
            "월~목요일 배정 및 이틀 연속 휴무 보장",
            "❌ 미흡",
            f"미충족: {', '.join(m_issues)}",
        )
    checklist.append(chk3)

    consecutive4_list = []
    for e, emp in enumerate(emp_list):
        emp_labels = flat_labels[e]
        w_count = 0
        ranges = []
        for d in range(1, num_days + 1):
            if "근무" in emp_labels[d - 1] or "지원" in emp_labels[d - 1]:
                w_count += 1
            else:
                if w_count >= 4:
                    ranges.append(f"{d - w_count}~{d - 1}일({w_count}연속)")
                w_count = 0
        if w_count >= 4:
            ranges.append(f"{num_days - w_count + 1}~{num_days}일({w_count}연속)")

        if ranges:
            consecutive4_list.append(
                f"{emp['name']}[{emp['lounge']}]({', '.join(ranges)})"
            )

    if not consecutive4_list:
        chk4 = (
            "4연속 근무 점검",
            "최대한 3연속 이하 근무 유도 (4연속 기피)",
            "✅ 미발생",
            "4연속 이상 근무자 없음",
        )
    else:
        chk4 = (
            "4연속 근무 점검",
            "최대한 3연속 이하 근무 유도 (4연속 기피)",
            "⚠️ 발생",
            f"발생 인원: {', '.join(consecutive4_list)}",
        )
    checklist.append(chk4)

    off_mismatch = []
    for e, emp in enumerate(emp_list):
        target_off = male_off_days if emp["gender"] == "남" else female_off_days
        actual_off_cnt = sum(
            1
            for d in range(1, num_days + 1)
            if flat_labels[e][d - 1] in ["휴무", "신청휴", "생휴"]
        )
        if actual_off_cnt != target_off:
            off_mismatch.append(f"{emp['name']}({actual_off_cnt}일)")

    if not off_mismatch:
        chk5 = (
            "월 목표 휴무일수",
            f"남성 {male_off_days}일 / 여성 {female_off_days}일 정확히 준수",
            "✅ 적합",
            "전 직원 지정 휴무일수 100% 달성",
        )
    else:
        chk5 = (
            "월 목표 휴무일수",
            f"남성 {male_off_days}일 / 여성 {female_off_days}일 정확히 준수",
            "❌ 미흡",
            f"일수 불일치: {', '.join(off_mismatch)}",
        )
    checklist.append(chk5)

    return checklist


# ==============================================================================
# 4. 엑셀 내보내기 함수
# ==============================================================================
def export_to_excel_single_sheet(
    lounge_schedules,
    lounge_employees,
    all_employees_flat,
    year,
    month,
    output_excel,
    public_holidays,
    male_off_days,
    female_off_days,
):
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
    border_box = Border(
        left=border_thin, right=border_thin, top=border_thin, bottom=border_thin
    )

    fill_map_excel = {
        "근무": PatternFill(start_color="E3F2FD", fill_type="solid"),
        "휴무": PatternFill(start_color="E0E0E0", fill_type="solid"),
        "신청휴": PatternFill(start_color="C8E6C9", fill_type="solid"),
        "생휴": PatternFill(start_color="FFE0B2", fill_type="solid"),
        "지원": PatternFill(start_color="FCE4D6", fill_type="solid"),
    }

    font_map_excel = {
        "근무": Font(name="맑은 고딕", size=9, bold=True, color="0D47A1"),
        "휴무": Font(name="맑은 고딕", size=9, color="424242"),
        "신청휴": Font(name="맑은 고딕", size=9, bold=True, color="1B5E20"),
        "생휴": Font(name="맑은 고딕", size=9, bold=True, color="E65100"),
        "지원": Font(name="맑은 고딕", size=9, bold=True, color="C65911"),
    }

    weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]

    ws.cell(
        row=1, column=1, value=f"📅 {year}년 {month}월 전 라운지 통합 근무표"
    ).font = font_title
    curr_r = 3

    for lounge_name in LOUNGE_LIST:
        if (
            lounge_name not in lounge_schedules
            or not lounge_employees[lounge_name]
        ):
            continue

        ws.cell(
            row=curr_r, column=1, value=f"■ {lounge_name} 라운지"
        ).font = font_section
        curr_r += 1

        headers = (
            ["라운지", "직급", "성명", "성별"]
            + [f"{d}일" for d in range(1, num_days + 1)]
            + ["근무", "휴무", "신청휴", "생휴", "휴무총합"]
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
            is_holiday = d in public_holidays

            c_day = ws.cell(
                row=curr_r, column=col_idx, value=weekdays_kr[w_idx]
            )
            c_day.font = (
                font_hdr_red if (w_idx >= 5 or is_holiday) else font_hdr
            )

            if w_idx == 5 and not is_holiday:
                c_day.fill = fill_sat
            elif w_idx == 6 or is_holiday:
                c_day.fill = fill_sun_hol
            else:
                c_day.fill = fill_navy

            c_day.alignment = Alignment(horizontal="center", vertical="center")
            c_day.border = border_box

        for c in list(range(1, 5)) + list(
            range(num_days + 5, num_days + 10)
        ):
            cell = ws.cell(row=curr_r, column=c)
            cell.fill = fill_navy
            cell.border = border_box

        curr_r += 1

        emp_list = lounge_employees[lounge_name]
        labels_list = lounge_schedules[lounge_name]
        start_emp_r = curr_r

        for e, emp in enumerate(emp_list):
            ws.cell(row=curr_r, column=1, value=emp["lounge"]).alignment = (
                Alignment(horizontal="center")
            )
            ws.cell(row=curr_r, column=2, value=emp["rank"]).alignment = (
                Alignment(horizontal="center")
            )
            ws.cell(row=curr_r, column=3, value=emp["name"]).alignment = (
                Alignment(horizontal="center")
            )
            ws.cell(row=curr_r, column=4, value=emp["gender"]).alignment = (
                Alignment(horizontal="center")
            )

            for i in range(1, 5):
                ws.cell(row=curr_r, column=i).border = border_box

            for d in range(1, num_days + 1):
                lbl = labels_list[e][d - 1]
                cell = ws.cell(row=curr_r, column=d + 4, value=lbl)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

                lookup_key = "지원" if "지원" in lbl else lbl
                cell.fill = fill_map_excel.get(
                    lookup_key, fill_map_excel["근무"]
                )
                cell.font = font_map_excel.get(
                    lookup_key, font_map_excel["근무"]
                )
                cell.border = border_box

            sum_col_start = num_days + 5
            start_letter = get_column_letter(5)
            end_letter = get_column_letter(num_days + 4)

            for i, h in enumerate(["근무", "휴무", "신청휴", "생휴"]):
                match_val = '="*근무*"' if h == "근무" else f'="{h}"'
                c_sum = ws.cell(
                    row=curr_r,
                    column=sum_col_start + i,
                    value=f"=COUNTIF({start_letter}{curr_r}:{end_letter}{curr_r}, {match_val})",
                )
                c_sum.alignment = Alignment(horizontal="center")
                c_sum.border = border_box

            off_start_col = get_column_letter(sum_col_start + 1)
            off_end_col = get_column_letter(sum_col_start + 3)
            c_tot_off = ws.cell(
                row=curr_r,
                column=sum_col_start + 4,
                value=f"=SUM({off_start_col}{curr_r}:{off_end_col}{curr_r})",
            )
            c_tot_off.alignment = Alignment(horizontal="center")
            c_tot_off.border = border_box

            curr_r += 1

        last_emp_r = curr_r - 1

        r_att = curr_r
        r_sup = curr_r + 1
        r_act = curr_r + 2

        for c in range(1, 5):
            cell = ws.cell(row=r_att, column=c)
            cell.fill = fill_att
            cell.font = font_att
            cell.border = border_box
        ws.merge_cells(
            start_row=r_att, start_column=1, end_row=r_att, end_column=4
        )
        ws.cell(row=r_att, column=1, value="출근인원").alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for c in range(1, 5):
            cell = ws.cell(row=r_sup, column=c)
            cell.fill = fill_sup_row
            cell.font = font_sup_row
            cell.border = border_box
        ws.merge_cells(
            start_row=r_sup, start_column=1, end_row=r_sup, end_column=4
        )
        ws.cell(row=r_sup, column=1, value="타 접점 지원").alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for c in range(1, 5):
            cell = ws.cell(row=r_act, column=c)
            cell.fill = fill_act
            cell.font = font_act
            cell.border = border_box
        ws.merge_cells(
            start_row=r_act, start_column=1, end_row=r_act, end_column=4
        )
        ws.cell(row=r_act, column=1, value="실제 근무인원").alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for d in range(1, num_days + 1):
            col_idx = d + 4
            col_let = get_column_letter(col_idx)

            c_att = ws.cell(
                row=r_att,
                column=col_idx,
                value=f'=COUNTIF({col_let}{start_emp_r}:{col_let}{last_emp_r}, "*근무*") + COUNTIF({col_let}{start_emp_r}:{col_let}{last_emp_r}, "*지원*")',
            )
            c_att.fill = fill_att
            c_att.font = font_att
            c_att.border = border_box
            c_att.alignment = Alignment(
                horizontal="center", vertical="center"
            )

            c_sup = ws.cell(row=r_sup, column=col_idx, value="")
            c_sup.fill = fill_sup_row
            c_sup.font = font_sup_row
            c_sup.border = border_box
            c_sup.alignment = Alignment(
                horizontal="center", vertical="center"
            )

            c_act = ws.cell(
                row=r_act,
                column=col_idx,
                value=f'={col_let}{r_att} + IF({col_let}{r_sup}="", 0, {col_let}{r_sup})',
            )
            c_act.fill = fill_act
            c_act.font = font_act
            c_act.border = border_box
            c_act.alignment = Alignment(
                horizontal="center", vertical="center"
            )

        curr_r = r_act + 3

    ws.cell(
        row=curr_r, column=1, value="📋 스케줄 최적화 연산 결과 검증 체크리스트"
    ).font = font_section
    curr_r += 1

    chk_headers = [
        "점검 항목",
        "검증 기준",
        "점검 결과",
        "세부 보고 및 미달 내용",
    ]
    chk_col_spans = [(1, 2), (3, 4), (5, 6), (7, num_days + 8)]

    for idx, h_text in enumerate(chk_headers):
        s_col, e_col = chk_col_spans[idx]
        for c in range(s_col, e_col + 1):
            cell = ws.cell(row=curr_r, column=c)
            cell.fill = fill_navy
            cell.font = font_hdr
            cell.border = border_box
        ws.merge_cells(
            start_row=curr_r,
            start_column=s_col,
            end_row=curr_r,
            end_column=e_col,
        )
        ws.cell(row=curr_r, column=s_col, value=h_text).alignment = Alignment(
            horizontal="center", vertical="center"
        )

    curr_r += 1

    flat_labels_all = []
    for emp in all_employees_flat:
        l = emp["lounge"]
        e_idx = next(
            i
            for i, e in enumerate(lounge_employees[l])
            if e["name"] == emp["name"]
        )
        flat_labels_all.append(lounge_schedules[l][e_idx])

    checklist_results = verify_schedule_checklist(
        all_employees_flat,
        flat_labels_all,
        year,
        month,
        male_off_days,
        female_off_days,
    )

    font_pass = Font(name="맑은 고딕", size=10, bold=True, color="1B5E20")
    fill_pass = PatternFill(start_color="C8E6C9", fill_type="solid")

    font_fail = Font(name="맑은 고딕", size=10, bold=True, color="C62828")
    fill_fail = PatternFill(start_color="FFCDD2", fill_type="solid")

    for item_name, criteria, status_str, detail_str in checklist_results:
        is_pass = "적합" in status_str or "미발생" in status_str
        res_font = font_pass if is_pass else font_fail
        res_fill = fill_pass if is_pass else fill_fail

        for c in range(1, 3):
            cell = ws.cell(row=curr_r, column=c)
            cell.border = border_box
        ws.merge_cells(
            start_row=curr_r, start_column=1, end_row=curr_r, end_column=2
        )
        ws.cell(row=curr_r, column=1, value=item_name).alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for c in range(3, 5):
            cell = ws.cell(row=curr_r, column=c)
            cell.border = border_box
        ws.merge_cells(
            start_row=curr_r, start_column=3, end_row=curr_r, end_column=4
        )
        ws.cell(row=curr_r, column=3, value=criteria).alignment = Alignment(
            horizontal="left", vertical="center"
        )

        for c in range(5, 7):
            cell = ws.cell(row=curr_r, column=c)
            cell.fill = res_fill
            cell.font = res_font
            cell.border = border_box
        ws.merge_cells(
            start_row=curr_r, start_column=5, end_row=curr_r, end_column=6
        )
        ws.cell(row=curr_r, column=5, value=status_str).alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for c in range(7, num_days + 9):
            cell = ws.cell(row=curr_r, column=c)
            cell.border = border_box
        ws.merge_cells(
            start_row=curr_r,
            start_column=7,
            end_row=curr_r,
            end_column=num_days + 8,
        )
        ws.cell(row=curr_r, column=7, value=detail_str).alignment = Alignment(
            horizontal="left", vertical="center"
        )

        curr_r += 1

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 6
    for d in range(1, num_days + 1):
        ws.column_dimensions[get_column_letter(d + 4)].width = 7

    wb.save(output_excel)


# ==============================================================================
# 5. Streamlit 메인 UI 대시보드 (다크모드 고 대비 CSS 적용)
# ==============================================================================
st.set_page_config(
    page_title="라운지 근무 스케줄 최적화 시스템",
    page_icon="📅",
    layout="wide",
)

# 다크 모드 및 라이트 모드 공통 시시성 보장 강제 CSS 설정
st.markdown(
    """
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Pretendard', sans-serif !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    
    /* 요약 카드 디자인 강제 고정 */
    [data-testid="stMetric"] {
        background-color: #ffffff !important;
        padding: 16px !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1) !important;
        border: 1px solid #e2e8f0 !important;
    }
    [data-testid="stMetricLabel"] p, [data-testid="stMetricValue"] div {
        color: #0f172a !important;
    }

    /* 근무 표 스타일 강제 고정 (다크모드 문자 안보임 방지) */
    .schedule-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
        text-align: center;
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    .schedule-table th {
        background-color: #1e293b !important;
        color: #ffffff !important;
        border: 1px solid #334155 !important;
        padding: 8px 4px !important;
    }
    .schedule-table td {
        border: 1px solid #cbd5e1 !important;
        padding: 6px 3px !important;
        background-color: #ffffff !important;
        color: #0f172a !important;
    }

    /* 뱃지 시각화 스타일 */
    .badge-work { background-color: #e0f2fe !important; color: #0369a1 !important; font-weight: bold; padding: 3px 6px; border-radius: 4px; display: inline-block; }
    .badge-off { background-color: #f1f5f9 !important; color: #475569 !important; padding: 3px 6px; border-radius: 4px; display: inline-block; }
    .badge-req { background-color: #dcfce7 !important; color: #15803d !important; font-weight: bold; padding: 3px 6px; border-radius: 4px; display: inline-block; }
    .badge-m { background-color: #fef3c7 !important; color: #b45309 !important; font-weight: bold; padding: 3px 6px; border-radius: 4px; display: inline-block; }
    .badge-sup { background-color: #ffedd5 !important; color: #c2410c !important; font-weight: bold; padding: 3px 6px; border-radius: 4px; display: inline-block; }
    </style>
""",
    unsafe_allow_html=True,
)

# 헤더
st.title("📅 전 라운지 통합 근무 스케줄 생성기")
st.caption(
    "Google OR-Tools 최적화 연산 엔진 기반 / 지원 근무 및 8대 핵심 제약 조건 100% 반영"
)

# 사이드바 설정 영역
st.sidebar.header("⚙️ 연산 기본 설정")
year = st.sidebar.number_input("연도", value=2026, step=1)
month = st.sidebar.number_input(
    "월", value=10, min_value=1, max_value=12, step=1
)
male_off = st.sidebar.number_input("남성 목표 휴무일수", value=11, step=1)
female_off = st.sidebar.number_input("여성 목표 휴무일수", value=12, step=1)
holidays_str = st.sidebar.text_input("공휴일 지정 (쉼표 구분)", value="3, 9")

public_holidays = [
    int(x.strip()) for x in holidays_str.split(",") if x.strip().isdigit()
]

# CSV 파일 처리
st.sidebar.divider()
st.sidebar.header("📂 직원 데이터")
uploaded_file = st.sidebar.file_uploader(
    "CSV 파일 업로드", type=["csv"], help="employees_input.csv 파일을 선택하세요"
)

use_sample = st.sidebar.checkbox("기본 샘플 데이터 사용", value=True)

if uploaded_file is not None:
    try:
        df_input = pd.read_csv(uploaded_file, encoding="utf-8-sig")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df_input = pd.read_csv(uploaded_file, encoding="cp949")
elif use_sample:
    import io

    df_input = pd.read_csv(io.StringIO(SAMPLE_CSV_DATA))
else:
    df_input = None

all_employees_flat = []
lounge_employees = {lounge: [] for lounge in LOUNGE_LIST}

if df_input is not None:
    name_col = next(
        (c for c in df_input.columns if "이름" in c or "성명" in c),
        df_input.columns[0],
    )
    gender_col = next(
        (c for c in df_input.columns if "성별" in c), df_input.columns[1]
    )
    off_col = next((c for c in df_input.columns if "휴무" in c), None)
    lounge_col = next((c for c in df_input.columns if "라운지" in c), None)
    m_col = next((c for c in df_input.columns if "생리" in c), None)
    rank_col = next((c for c in df_input.columns if "직급" in c), None)

    for idx, row in df_input.iterrows():
        name = (
            str(row[name_col]).strip()
            if pd.notna(row[name_col])
            else f"직원{idx+1}"
        )
        gender = (
            str(row[gender_col]).strip() if pd.notna(row[gender_col]) else "여"
        )
        raw_lounge = (
            str(row[lounge_col]).strip()
            if lounge_col and pd.notna(row[lounge_col])
            else "YP"
        )
        lounge = LOUNGE_ALIAS.get(raw_lounge, raw_lounge)
        rank = (
            str(row[rank_col]).strip()
            if rank_col and pd.notna(row[rank_col])
            else "사원"
        )

        req_off = []
        if off_col and pd.notna(row[off_col]):
            for item in str(row[off_col]).replace(";", ",").split(","):
                if item.strip().isdigit():
                    req_off.append(int(item.strip()))

        m_off = None
        if m_col and pd.notna(row[m_col]):
            if str(row[m_col]).strip().isdigit():
                m_off = int(str(row[m_col]).strip())

        emp_dict = {
            "name": name,
            "gender": gender,
            "lounge": lounge,
            "rank": rank,
            "req_off": req_off,
            "m_off": m_off,
        }
        all_employees_flat.append(emp_dict)
        lounge_employees.setdefault(lounge, []).append(emp_dict)

# 요약 지표 카드
m1, m2, m3, m4 = st.columns(4)
m1.metric("👥 총 인원", f"{len(all_employees_flat)}명")
m2.metric("📅 대상 연월", f"{year}년 {month}월")
m3.metric("🎯 목표 휴무", f"남 {male_off}일 / 여 {female_off}일")
m4.metric("🎈 공휴일", f"{len(public_holidays)}일 지정 ({public_holidays})")

st.divider()

# 메인 탭 생성
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "🗓️ 월간 통합 근무표",
        "📋 검증 체크리스트",
        "👥 직원 데이터 관리",
        "📥 엑셀 내보내기",
    ]
)

# ------------------------------------------------------------------------------
# TAB 1: 월간 통합 근무표
# ------------------------------------------------------------------------------
with tab1:
    col_btn, col_info = st.columns([1, 4])
    with col_btn:
        run_btn = st.button(
            "⚡ 스케줄 자동 생성", type="primary", use_container_width=True
        )

    if run_btn or "flat_labels" in st.session_state:
        if run_btn:
            with st.spinner("구글 OR-Tools 최적화 연산 진행 중..."):
                flat_labels = solve_global_schedule(
                    all_employees_flat,
                    year,
                    month,
                    male_off,
                    female_off,
                    LOUNGE_WORKER_BOUNDS,
                    public_holidays,
                )
                st.session_state["flat_labels"] = flat_labels
        else:
            flat_labels = st.session_state["flat_labels"]

        if flat_labels is None:
            st.error("❌ 조건에 맞는 스케줄 조합을 찾을 수 없습니다.")
        else:
            st.success("🎉 최적 스케줄 연산 완료!")

            lounge_schedules = {lounge: [] for lounge in LOUNGE_LIST}
            for i, e in enumerate(all_employees_flat):
                lounge_schedules[e["lounge"]].append(flat_labels[i])

            num_days = calendar.monthrange(year, month)[1]
            weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]

            for lounge_name in LOUNGE_LIST:
                if (
                    lounge_name not in lounge_schedules
                    or not lounge_employees[lounge_name]
                ):
                    continue

                st.subheader(f"■ {lounge_name} 라운지")

                html = '<div style="overflow-x:auto;"><table class="schedule-table">'
                html += '<thead><tr style="background-color:#1e293b; color:white;">'
                html += "<th>직급</th><th>성명</th><th>성별</th>"

                for d in range(1, num_days + 1):
                    w = datetime.date(year, month, d).weekday()
                    is_hol = d in public_holidays or w >= 5
                    bg_col = "#ef4444" if is_hol else "#1e293b"
                    html += f'<th style="background-color:{bg_col} !important; color:#ffffff !important;">{d}<br><span style="font-size:10px;">{weekdays_kr[w]}</span></th>'

                html += "<th>근무</th><th>휴무</th><th>신청휴</th><th>생휴</th><th>휴무총합</th></tr></thead><tbody>"

                l_emps = lounge_employees[lounge_name]
                l_labels = lounge_schedules[lounge_name]

                for e_idx, emp in enumerate(l_emps):
                    labels = l_labels[e_idx]
                    work_cnt = sum(
                        1
                        for x in labels
                        if "근무" in x or "지원" in x
                    )
                    off_cnt = sum(1 for x in labels if x == "휴무")
                    req_cnt = sum(1 for x in labels if x == "신청휴")
                    m_cnt = sum(1 for x in labels if x == "생휴")

                    html += f'<tr><td style="font-weight:bold; color:#0f172a !important;">{emp["rank"]}</td><td style="font-weight:bold; color:#0f172a !important;">{emp["name"]}</td><td style="color:#0f172a !important;">{emp["gender"]}</td>'

                    for d in range(1, num_days + 1):
                        lbl = labels[d - 1]
                        badge_cls = "badge-work"
                        if "지원" in lbl:
                            badge_cls = "badge-sup"
                        elif lbl == "휴무":
                            badge_cls = "badge-off"
                        elif lbl == "신청휴":
                            badge_cls = "badge-req"
                        elif lbl == "생휴":
                            badge_cls = "badge-m"

                        html += f'<td><span class="{badge_cls}">{lbl}</span></td>'

                    html += f'<td style="font-weight:bold; color:#0f172a !important;">{work_cnt}</td><td style="color:#0f172a !important;">{off_cnt}</td><td style="color:#1b5e20 !important; font-weight:bold;">{req_cnt}</td><td style="color:#e65100 !important; font-weight:bold;">{m_cnt}</td><td style="background-color:#f1f5f9 !important; color:#0f172a !important; font-weight:bold;">{off_cnt+req_cnt+m_cnt}</td></tr>'

                # 일별 인원 수치 명확한 배경/글자색 고정
                att_counts, sup_counts, act_counts = [], [], []
                for d in range(1, num_days + 1):
                    home_cnt = sum(
                        1
                        for x in l_labels
                        if "근무" in x[d - 1] or "지원" in x[d - 1]
                    )
                    out_cnt = sum(1 for x in l_labels if "지원" in x[d - 1])

                    in_cnt = 0
                    for other_l in LOUNGE_LIST:
                        if other_l == lounge_name:
                            continue
                        for other_lbls in lounge_schedules[other_l]:
                            if f"지원({lounge_name})" in other_lbls[d - 1]:
                                in_cnt += 1

                    net_sup = in_cnt - out_cnt
                    att_counts.append(home_cnt)
                    sup_counts.append(
                        "-"
                        if net_sup == 0
                        else f"+{net_sup}"
                        if net_sup > 0
                        else f"{net_sup}"
                    )
                    act_counts.append(home_cnt + net_sup)

                html += f'<tr style="background-color:#eff6ff !important; font-weight:bold; color:#1e40af !important;"><td colspan="3" style="background-color:#eff6ff !important; color:#1e40af !important;">출근 인원</td>'
                for c in att_counts:
                    html += f'<td style="background-color:#eff6ff !important; color:#1e40af !important;">{c}</td>'
                html += '<td colspan="5" style="background-color:#eff6ff !important;"></td></tr>'

                html += f'<tr style="background-color:#fff7ed !important; font-weight:bold; color:#c65911 !important;"><td colspan="3" style="background-color:#fff7ed !important; color:#c65911 !important;">타 접점 지원</td>'
                for c in sup_counts:
                    html += f'<td style="background-color:#fff7ed !important; color:#c65911 !important;">{c}</td>'
                html += '<td colspan="5" style="background-color:#fff7ed !important;"></td></tr>'

                html += f'<tr style="background-color:#f0fdf4 !important; font-weight:bold; color:#15803d !important;"><td colspan="3" style="background-color:#f0fdf4 !important; color:#15803d !important;">실제 근무인원</td>'
                for c in act_counts:
                    html += f'<td style="background-color:#f0fdf4 !important; color:#15803d !important;">{c}</td>'
                html += '<td colspan="5" style="background-color:#f0fdf4 !important;"></td></tr>'

                html += "</tbody></table></div><br>"
                st.markdown(html, unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 2: 검증 체크리스트
# ------------------------------------------------------------------------------
with tab2:
    st.subheader("📋 스케줄 최적화 연산 결과 검증 체크리스트")

    if "flat_labels" in st.session_state and st.session_state["flat_labels"]:
        checklist = verify_schedule_checklist(
            all_employees_flat,
            st.session_state["flat_labels"],
            year,
            month,
            male_off,
            female_off,
        )

        chk_df = pd.DataFrame(
            checklist,
            columns=[
                "점검 항목",
                "검증 기준",
                "점검 결과",
                "세부 보고 내용",
            ],
        )
        st.dataframe(chk_df, use_container_width=True)
    else:
        st.info("먼저 [월간 통합 근무표] 탭에서 [⚡ 스케줄 자동 생성] 버튼을 클릭하세요.")

# ------------------------------------------------------------------------------
# TAB 3: 직원 데이터 관리
# ------------------------------------------------------------------------------
with tab3:
    st.subheader("👥 불러온 직원 데이터 목록")
    if df_input is not None:
        st.dataframe(df_input, use_container_width=True)

        col_l, col_r = st.columns(2)
        with col_l:
            st.write("#### 라운지별 인원 현황")
            st.bar_chart(df_input["라운지"].value_counts())
        with col_r:
            st.write("#### 직급별 인원 현황")
            st.bar_chart(df_input["직급"].value_counts())

# ------------------------------------------------------------------------------
# TAB 4: 엑셀 내보내기 & 양식
# ------------------------------------------------------------------------------
with tab4:
    st.subheader("📥 엑셀 내보내기 및 양식 다운로드")

    if "flat_labels" in st.session_state and st.session_state["flat_labels"]:
        lounge_schedules = {lounge: [] for lounge in LOUNGE_LIST}
        for i, e in enumerate(all_employees_flat):
            lounge_schedules[e["lounge"]].append(
                st.session_state["flat_labels"][i]
            )

        with tempfile.NamedTemporaryFile(
            delete=False, suffix=".xlsx"
        ) as tmp_excel:
            export_to_excel_single_sheet(
                lounge_schedules,
                lounge_employees,
                all_employees_flat,
                year,
                month,
                tmp_excel.name,
                public_holidays,
                male_off,
                female_off,
            )

            with open(tmp_excel.name, "rb") as f:
                st.download_button(
                    label="📊 생성된 엑셀 파일 다운로드 (.xlsx)",
                    data=f.read(),
                    file_name=f"월간근무표_{year}년_{month}월.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
    else:
        st.info("스케줄을 먼저 생성하면 엑셀 다운로드 버튼이 활성화됩니다.")

    st.divider()
    st.download_button(
        label="📄 CSV 입력 샘플 양식 다운로드",
        data=SAMPLE_CSV_DATA.encode("utf-8-sig"),
        file_name="employees_sample_template.csv",
        mime="text/csv",
    )