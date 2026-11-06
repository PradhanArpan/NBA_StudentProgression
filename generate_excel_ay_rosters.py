import json
import math
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def get_last_active_sem(s):
    active = []
    for sm_idx in range(1, 9):
        sm_k = f"sem{sm_idx}"
        v = s.get('semesters', {}).get(sm_k, {})
        has_marks = (v.get('sgpa') is not None) or (v.get('pct') is not None)
        has_stat = v.get('status') in ['Pass', 'Fail', 'FC-D', 'FC', 'SC', 'P']
        has_src = bool(v.get('sources'))
        if has_marks or has_stat or has_src:
            active.append(sm_idx)
    return max(active) if active else 0

def get_nba_tag(s):
    prog = s.get('progressionStatus', 'Normal')
    cat = s.get('category', 'Regular')
    disc_sem = s.get('discontinuedFromSem')
    yb_sem = s.get('yearBackFromSem')
    yb_batch = s.get('yearBackJuniorBatch')
    
    if prog == 'Discontinued' or s.get('gradStatus') == 'Discontinued / Transferred':
        return f"N6 - Discontinued (Left from Sem {disc_sem or 1})"
    elif prog == 'Year-Back' or s.get('isYearBackExit'):
        return f"YB - Year-Back (Detained at Sem {yb_sem or 4} -> Batch {yb_batch or 'Junior'})"
    elif prog == 'Readmitted' or 'Readmitted' in cat:
        origin_b = s.get('yearBackOrigin', {}).get('batch') or ''
        return f"N5 - Readmitted" + (f" (from Batch {origin_b})" if origin_b else "")
    elif cat == 'Lateral':
        return "N2 - Lateral Entry"
    else:
        return "N1 - Regular Admission"

def get_ay_year_roster(students, ay_start, year_level):
    base_batch = str(ay_start - (year_level - 1))
    odd_key = f"sem{year_level * 2 - 1}"
    even_key = f"sem{year_level * 2}"
    roster = []
    included_regnos = set()

    # 1. Base batch students
    for s in students:
        if s['batch'] != base_batch:
            continue

        cat = s.get('category', 'Regular')
        prog = s.get('progressionStatus', 'Normal')
        disc_sem = s.get('discontinuedFromSem')
        yb_sem = s.get('yearBackFromSem')
        yb_batch = s.get('yearBackJuniorBatch') or str(int(base_batch) + 1)

        # Lateral entry admitted directly in Year II (Semester 3)
        if year_level == 1 and cat == 'Lateral':
            continue

        # Readmitted students admitted after gap year in Year II or III
        if prog == 'Readmitted' or 'Readmitted' in cat:
            if year_level == 1:
                continue
            origin_reg = s.get('yearBackOrigin', {}).get('regNo') if s.get('yearBackOrigin') else None
            origin_s = next((st for st in students if st['regNo'] == origin_reg), None) if origin_reg else None
            origin_yb_sem = origin_s.get('yearBackFromSem') if origin_s else (4 if s.get('semesters', {}).get('sem4', {}).get('status') == 'Pass' else 2)
            join_year = math.ceil((origin_yb_sem or 4) / 2) + 1
            if year_level < join_year:
                continue

        sem_odd = s.get('semesters', {}).get(odd_key, {})
        sem_even = s.get('semesters', {}).get(even_key, {})

        odd_sgpa = sem_odd.get('sgpa')
        odd_pct = sem_odd.get('pct')
        odd_stat = sem_odd.get('status')

        even_sgpa = sem_even.get('sgpa')
        even_pct = sem_even.get('pct')
        even_stat = sem_even.get('status')

        odd_str = f"SGPA {float(odd_sgpa):.2f} ({odd_pct}%)" if odd_sgpa else (odd_stat or "-")
        even_str = f"SGPA {float(even_sgpa):.2f} ({even_pct}%)" if even_sgpa else (even_stat or "-")

        is_disc = (prog == 'Discontinued' or s.get('gradStatus') == 'Discontinued / Transferred')
        is_yb = (prog == 'Year-Back' or s.get('isYearBackExit'))

        nba_tag = get_nba_tag(s)
        is_active = True
        is_discontinued = False
        is_off_campus = False

        if is_disc:
            eff_disc_sem = disc_sem or 1
            max_attended_year = math.ceil(eff_disc_sem / 2)
            is_discontinued = True
            if year_level > max_attended_year:
                # Retained on cohort roster as requested for complete N1-N6 audit trail
                odd_str = "— (Discontinued)"
                even_str = "— (Discontinued)"
                status_desc = f"Discontinued (Left from Sem {eff_disc_sem})"
                remarks_desc = f"Discontinued in Semester {eff_disc_sem} (N6 Exited Program)"
                is_active = False
            else:
                status_desc = f"Discontinued (Left from Sem {eff_disc_sem})"
                remarks_desc = f"Exited program from Semester {eff_disc_sem}"
                is_active = False
        elif is_yb:
            detained_sem = yb_sem or 4
            detained_year = math.ceil(detained_sem / 2)
            if year_level > detained_year:
                # Off-campus gap year; not taking courses with base batch
                odd_str = "— (Detained)"
                even_str = "— (Detained)"
                status_desc = f"Off-Campus (Detained at Sem {detained_sem} -> Batch {yb_batch})"
                remarks_desc = f"Detained at Sem {detained_sem}; off-campus before rejoining Batch {yb_batch} at Sem {detained_sem + 1}"
                is_off_campus = True
                is_active = False
            elif year_level == detained_year:
                status_desc = f"Year-Back (Detained at Sem {detained_sem} -> Batch {yb_batch})"
                remarks_desc = f"Detained at Sem {detained_sem}; rejoining Batch {yb_batch}"
                is_active = True
            else:
                status_desc = "Pass (All Cleared)" if (odd_stat == 'Pass' and even_stat == 'Pass') else "Enrolled / Active"
                remarks_desc = "Active progression in initial study year"
                is_active = True
        else:
            # Normal / Regular / Lateral / Readmitted active progression
            if sem_odd.get('status') == 'Fail' or sem_even.get('status') == 'Fail':
                status_desc = 'With Backlog'
            elif sem_odd.get('status') == 'Pass' and sem_even.get('status') == 'Pass':
                status_desc = 'Pass (All Cleared)'
            elif sem_odd.get('status') == 'Pass' or sem_even.get('status') == 'Pass':
                status_desc = 'In Progress / Part Verified'
            elif s.get('gradStatus') and 'No Backlog' in s.get('gradStatus') and year_level == 4:
                status_desc = 'Graduated (Clear)'
            elif s.get('gradStatus') and 'With Backlog' in s.get('gradStatus') and year_level == 4:
                status_desc = 'Graduated (With Backlog)'
            else:
                status_desc = 'Enrolled / Active'

            remarks_parts = []
            if cat == 'Lateral':
                remarks_parts.append('Lateral Entry in S3')
            elif prog == 'Readmitted' or 'Readmitted' in cat:
                orig_b = s.get('yearBackOrigin', {}).get('batch') or ''
                remarks_parts.append(f'Readmitted after Year-Back (Orig. Batch {orig_b})')
            else:
                remarks_parts.append('Regular Admission')
            remarks_desc = '; '.join(remarks_parts)
            is_active = True

        roster.append({
            'regNo': s['regNo'],
            'name': s['name'],
            'gender': s.get('gender', '-'),
            'nba_tag': nba_tag,
            'category': cat,
            'batch': f"Batch {s['batch']}",
            'odd_sem': f"{odd_key.upper()}: {odd_str}",
            'even_sem': f"{even_key.upper()}: {even_str}",
            'status': status_desc,
            'remarks': remarks_desc,
            'is_active': is_active,
            'is_discontinued': is_discontinued,
            'is_off_campus': is_off_campus
        })
        included_regnos.add(s['regNo'])

    # 2. Check for students detained from a senior batch who transfer into this junior cohort
    for s in students:
        prog = s.get('progressionStatus')
        target_batch = s.get('yearBackJuniorBatch') or (s.get('yearBackTransferredTo', {}).get('batch') if s.get('yearBackTransferredTo') else None)
        readmit_regno = s.get('readmissionRegNo') or (s.get('yearBackTransferredTo', {}).get('regNo') if s.get('yearBackTransferredTo') else None)

        if (prog == 'Year-Back' or s.get('isYearBackExit')) and target_batch == base_batch:
            has_readmit_record = bool(readmit_regno and any(st['regNo'] == readmit_regno for st in students))
            if not has_readmit_record and s['regNo'] not in included_regnos:
                yb_sem = s.get('yearBackFromSem', 4)
                join_year_level = math.ceil(yb_sem / 2) + 1
                if year_level >= join_year_level:
                    sem_odd = s.get('semesters', {}).get(odd_key, {})
                    sem_even = s.get('semesters', {}).get(even_key, {})

                    odd_sgpa = sem_odd.get('sgpa')
                    odd_pct = sem_odd.get('pct')
                    odd_stat = sem_odd.get('status')

                    even_sgpa = sem_even.get('sgpa')
                    even_pct = sem_even.get('pct')
                    even_stat = sem_even.get('status')

                    odd_str = f"SGPA {float(odd_sgpa):.2f} ({odd_pct}%)" if odd_sgpa else (odd_stat or "-")
                    even_str = f"SGPA {float(even_sgpa):.2f} ({even_pct}%)" if even_sgpa else (even_stat or "-")

                    roster.append({
                        'regNo': s['regNo'],
                        'name': s['name'],
                        'gender': s.get('gender', '-'),
                        'nba_tag': f"N5 - Readmitted (from Batch {s['batch']})",
                        'category': 'Readmitted',
                        'batch': f"Batch {base_batch} (ex-Batch {s['batch']})",
                        'odd_sem': f"{odd_key.upper()}: {odd_str}",
                        'even_sem': f"{even_key.upper()}: {even_str}",
                        'status': f"Year-Back (Rejoined Batch {base_batch} in Sem {yb_sem + 1})",
                        'remarks': f"Year-Back from Batch {s['batch']}; joined Batch {base_batch} at Sem {yb_sem + 1}",
                        'is_active': True,
                        'is_discontinued': False,
                        'is_off_campus': False
                    })
                    included_regnos.add(s['regNo'])

    # Sort: active students first, then off-campus/discontinued, by Reg No
    roster.sort(key=lambda x: (1 if (x.get('is_discontinued') or x.get('is_off_campus')) else 0, x['regNo']))
    return roster

def generate_ay_excel(json_path=None, output_path=None):
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    if json_path is None:
        json_path = os.path.join(cur_dir, 'student_data.json')
    if output_path is None:
        output_path = os.path.join(cur_dir, 'Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx')

    with open(json_path, 'r', encoding='utf-8') as f:
        db = json.load(f)

    students = db['students']

    academic_years = [
        ("AY 2026-27", 2026, "Current Academic Year (CAY)"),
        ("AY 2025-26", 2025, "CAYm1"),
        ("AY 2024-25", 2024, "CAYm2"),
        ("AY 2023-24", 2023, "CAYm3 / Benchmark Graduated Batch (2020)"),
        ("AY 2022-23", 2022, "CAYm4 / LYG"),
        ("AY 2021-22", 2021, "CAYm5 / LYGm1"),
        ("AY 2020-21", 2020, "CAYm6 / LYGm2")
    ]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    NAVY = "1F497D"
    LIGHT_BLUE = "DCE6F1"
    WHITE = "FFFFFF"
    BORDER_GRAY = "CBD5E1"
    LIGHT_CYAN = "EDF2F8"

    font_title = Font(name="Calibri", size=16, bold=True, color="1F497D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
    font_sec_hdr = Font(name="Calibri", size=11, bold=True, color=WHITE)
    font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color=WHITE)
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    font_italic = Font(name="Calibri", size=9, italic=True, color="64748B")

    # NBA Tag Font Colors
    font_tag_n1 = Font(name="Calibri", size=9, bold=True, color="1E3A8A") # Navy
    font_tag_n2 = Font(name="Calibri", size=9, bold=True, color="065F46") # Green
    font_tag_n5 = Font(name="Calibri", size=9, bold=True, color="581C87") # Purple
    font_tag_n6 = Font(name="Calibri", size=9, bold=True, color="991B1B") # Red
    font_tag_yb = Font(name="Calibri", size=9, bold=True, color="92400E") # Amber

    # Row Fills
    fill_disc_row = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid") # Rose/red tint
    fill_yb_row = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid") # Amber tint
    fill_alt_row = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Slate tint
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    fill_navy = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    fill_light_blue = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
    fill_light_cyan = PatternFill(start_color=LIGHT_CYAN, end_color=LIGHT_CYAN, fill_type="solid")

    thin_border_side = Side(style='thin', color=BORDER_GRAY)
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # ================= 1. EXECUTIVE SUMMARY SHEET =================
    ws_sum = wb.create_sheet(title="Executive Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.cell(row=2, column=2, value="B.Tech Civil Engineering - Academic Year Cohort Progression & NBA Criterion 4 Audit").font = font_title
    ws_sum.cell(row=3, column=2, value="Official Academic Year Rosters with N1, N2, N5, N6, and YB Cohort Accounting (AY 2020-21 to AY 2026-27)").font = font_subtitle

    sum_headers = [
        "Academic Year", 
        "NBA Status", 
        "I Year (Active / Admitted)", 
        "II Year (Active / Admitted)", 
        "III Year (Active / Admitted)", 
        "IV Year (Active / Admitted)", 
        "Total Active Enrolled",
        "Total Cohort Admitted"
    ]
    for col_idx, h in enumerate(sum_headers, start=2):
        cell = ws_sum.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = thin_border
    ws_sum.row_dimensions[5].height = 28

    row_idx = 6
    for ay_label, ay_start, ay_status in academic_years:
        r1 = get_ay_year_roster(students, ay_start, 1)
        r2 = get_ay_year_roster(students, ay_start, 2)
        r3 = get_ay_year_roster(students, ay_start, 3)
        r4 = get_ay_year_roster(students, ay_start, 4)

        act1 = sum(1 for s in r1 if s.get('is_active', True))
        act2 = sum(1 for s in r2 if s.get('is_active', True))
        act3 = sum(1 for s in r3 if s.get('is_active', True))
        act4 = sum(1 for s in r4 if s.get('is_active', True))

        tot_act = act1 + act2 + act3 + act4
        tot_cohort = len(r1) + len(r2) + len(r3) + len(r4)

        ws_sum.cell(row=row_idx, column=2, value=ay_label).font = font_bold
        ws_sum.cell(row=row_idx, column=2).alignment = align_center
        ws_sum.cell(row=row_idx, column=2).border = thin_border

        ws_sum.cell(row=row_idx, column=3, value=ay_status).font = font_regular
        ws_sum.cell(row=row_idx, column=3).alignment = align_left
        ws_sum.cell(row=row_idx, column=3).border = thin_border

        vals = [
            f"{act1} / {len(r1)}" if len(r1) > 0 else "-",
            f"{act2} / {len(r2)}" if len(r2) > 0 else "-",
            f"{act3} / {len(r3)}" if len(r3) > 0 else "-",
            f"{act4} / {len(r4)}" if len(r4) > 0 else "-",
            tot_act,
            tot_cohort
        ]

        for c_idx, val in enumerate(vals, start=4):
            cell = ws_sum.cell(row=row_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [8, 9] else font_regular
            cell.alignment = align_center
            cell.border = thin_border
            if c_idx == 8:
                cell.fill = fill_light_cyan
            elif c_idx == 9:
                cell.fill = fill_light_blue

        ws_sum.row_dimensions[row_idx].height = 22
        row_idx += 1

    # Total Row
    ws_sum.cell(row=row_idx, column=2, value="Cumulative Over 7 AYs").font = font_bold
    ws_sum.cell(row=row_idx, column=2).alignment = align_center
    ws_sum.cell(row=row_idx, column=2).border = thin_border
    ws_sum.cell(row=row_idx, column=3, value="Student-Years Audited").font = font_italic
    ws_sum.cell(row=row_idx, column=3).border = thin_border

    ws_sum.cell(row=row_idx, column=4, value="-").alignment = align_center
    ws_sum.cell(row=row_idx, column=5, value="-").alignment = align_center
    ws_sum.cell(row=row_idx, column=6, value="-").alignment = align_center
    ws_sum.cell(row=row_idx, column=7, value="-").alignment = align_center
    for c in range(4, 8):
        ws_sum.cell(row=row_idx, column=c).border = thin_border

    cell_act = ws_sum.cell(row=row_idx, column=8, value=f"=SUM(H6:H{row_idx-1})")
    cell_act.font = font_bold
    cell_act.alignment = align_center
    cell_act.border = thin_border
    cell_act.fill = fill_light_cyan

    cell_tot = ws_sum.cell(row=row_idx, column=9, value=f"=SUM(I6:I{row_idx-1})")
    cell_tot.font = font_bold
    cell_tot.alignment = align_center
    cell_tot.border = thin_border
    cell_tot.fill = fill_light_blue

    ws_sum.row_dimensions[row_idx].height = 24

    # Explanation notes
    note_row = row_idx + 3
    notes = [
        ("NOTE 1 (NBA Criterion 4 Tagging):", "Every student displays their official classification: [N1] Regular 1st Year Admission, [N2] Lateral Entry Direct 2nd Year, [N5] Readmitted after Academic Break / Year-Back, [N6] Discontinued / Transferred out, and [YB] Year-Back / Detained."),
        ("NOTE 2 (Discontinued Students N6):", "Students who discontinued are included on their cohort rosters with explicit [N6] tags so that the sum of admitted students (N1 + N2 + N5 - N6) perfectly tallies with official Table 4.1. Discontinued rows are shaded in soft rose."),
        ("NOTE 3 (Year-Back Students YB):", "If a student suffers a Year-Back at Sem 4, they are listed as Detained [YB] with their base batch, do not repeat Year II (gap year off-campus), and join the junior batch at Sem 5 as Readmitted [N5]."),
        ("NOTE 4 (Lateral Entry N2):", "Lateral Entry students enter directly in Year II (Semester 3). Hence, they are present in Year II, III, and IV rosters, but absent from Year I."),
        ("NOTE 5 (Zero-Hypothetical Policy):", "Only verified records from the official COE marksheets, registration lists, and certificates are displayed.")
    ]
    for title, desc in notes:
        ws_sum.cell(row=note_row, column=2, value=title).font = font_bold
        ws_sum.cell(row=note_row, column=3, value=desc).font = font_regular
        ws_sum.merge_cells(start_row=note_row, start_column=3, end_row=note_row, end_column=9)
        note_row += 1

    ws_sum.column_dimensions['B'].width = 16
    ws_sum.column_dimensions['C'].width = 38
    ws_sum.column_dimensions['D'].width = 24
    ws_sum.column_dimensions['E'].width = 24
    ws_sum.column_dimensions['F'].width = 24
    ws_sum.column_dimensions['G'].width = 24
    ws_sum.column_dimensions['H'].width = 22
    ws_sum.column_dimensions['I'].width = 22

    # ================= 2. INDIVIDUAL ACADEMIC YEAR SHEETS =================
    year_titles = [
        (1, "I YEAR (First Year - Semesters 1 & 2)", "Freshman Cohort Admitted in Year I", "1F497D"),
        (2, "II YEAR (Second Year - Semesters 3 & 4)", "Sophomore Cohort (Regular + Lateral Entry + Readmitted)", "2E75B6"),
        (3, "III YEAR (Third Year - Semesters 5 & 6)", "Junior Cohort Continuing in Core Engineering", "385723"),
        (4, "IV YEAR (Fourth Year - Semesters 7 & 8)", "Senior Cohort / Final Year Graduating Class", "7030A0")
    ]

    for ay_label, ay_start, ay_status in academic_years:
        sheet_title = ay_label.replace("AY ", "AY_")
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        ws.cell(row=2, column=2, value=f"B.Tech Civil Engineering — Student Roster for {ay_label}").font = font_title
        ws.cell(row=3, column=2, value=f"Academic Classification: {ay_status} | Complete List Across I, II, III, and IV Years with NBA Categories").font = font_subtitle

        curr_row = 5

        for yr_lvl, yr_name, yr_sub, color_hex in year_titles:
            roster = get_ay_year_roster(students, ay_start, yr_lvl)
            cohort_b = str(ay_start - (yr_lvl - 1))

            act_cnt = sum(1 for s in roster if s.get('is_active', True))
            disc_cnt = sum(1 for s in roster if s.get('is_discontinued', False))
            yb_cnt = sum(1 for s in roster if s.get('is_off_campus', False))

            sec_title = f"{yr_name} — Batch {cohort_b} (Total Admitted: {len(roster)} | Active Enrolled: {act_cnt} | Discontinued [N6]: {disc_cnt} | Off-Campus [YB]: {yb_cnt})"
            ws.cell(row=curr_row, column=2, value=sec_title).font = font_sec_hdr
            ws.cell(row=curr_row, column=2).fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
            ws.cell(row=curr_row, column=2).alignment = align_left
            ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=12)
            ws.row_dimensions[curr_row].height = 26
            curr_row += 1

            if len(roster) == 0:
                ws.cell(row=curr_row, column=2, value=f"No student records in repository for Batch {cohort_b} (Pre-2020 cohort not covered in COE dataset).").font = font_italic
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=12)
                ws.row_dimensions[curr_row].height = 20
                curr_row += 2
                continue

            cols = [
                ("Sl", 6),
                ("Reg. No.", 14),
                ("Student Name", 28),
                ("Gender", 10),
                ("NBA Category Tag", 28),
                ("Category", 16),
                ("Batch", 14),
                ("Odd Semester", 26),
                ("Even Semester", 26),
                ("Academic Status", 24),
                ("Remarks", 32)
            ]

            for c_i, (col_name, col_w) in enumerate(cols, start=2):
                c_cell = ws.cell(row=curr_row, column=c_i, value=col_name)
                c_cell.font = font_tbl_hdr
                c_cell.fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                c_cell.alignment = align_center if c_i in [2, 3, 5, 6, 7, 8] else align_left
                c_cell.border = thin_border
            ws.row_dimensions[curr_row].height = 22
            curr_row += 1

            for s_idx, st in enumerate(roster, start=1):
                row_fill = fill_white
                if st.get('is_discontinued'):
                    row_fill = fill_disc_row
                elif st.get('is_off_campus'):
                    row_fill = fill_yb_row
                elif s_idx % 2 == 0:
                    row_fill = fill_alt_row

                # Sl No
                ws.cell(row=curr_row, column=2, value=s_idx).alignment = align_center
                # Reg No
                c_reg = ws.cell(row=curr_row, column=3, value=st['regNo'])
                c_reg.alignment = align_center
                c_reg.font = font_bold
                # Student Name
                c_name = ws.cell(row=curr_row, column=4, value=st['name'])
                c_name.alignment = align_left
                c_name.font = font_bold
                # Gender
                ws.cell(row=curr_row, column=5, value=st['gender']).alignment = align_center
                
                # NBA Category Tag
                c_tag = ws.cell(row=curr_row, column=6, value=st['nba_tag'])
                c_tag.alignment = align_center
                if "N1" in st['nba_tag']:
                    c_tag.font = font_tag_n1
                elif "N2" in st['nba_tag']:
                    c_tag.font = font_tag_n2
                elif "N5" in st['nba_tag']:
                    c_tag.font = font_tag_n5
                elif "N6" in st['nba_tag']:
                    c_tag.font = font_tag_n6
                elif "YB" in st['nba_tag']:
                    c_tag.font = font_tag_yb
                else:
                    c_tag.font = font_bold

                # Category
                ws.cell(row=curr_row, column=7, value=st['category']).alignment = align_center
                # Batch
                ws.cell(row=curr_row, column=8, value=st['batch']).alignment = align_center
                # Odd Sem
                ws.cell(row=curr_row, column=9, value=st['odd_sem']).alignment = align_left
                # Even Sem
                ws.cell(row=curr_row, column=10, value=st['even_sem']).alignment = align_left
                
                # Academic Status
                c_stat = ws.cell(row=curr_row, column=11, value=st['status'])
                c_stat.alignment = align_center
                if "Backlog" in st['status'] or "Discontinued" in st['status']:
                    c_stat.font = Font(name="Calibri", size=10, bold=True, color="C00000")
                elif "Pass" in st['status'] or "Graduated" in st['status'] or "Clear" in st['status']:
                    c_stat.font = Font(name="Calibri", size=10, bold=True, color="008000")
                elif "Year-Back" in st['status'] or "Off-Campus" in st['status']:
                    c_stat.font = Font(name="Calibri", size=10, bold=True, color="B45309")
                else:
                    c_stat.font = font_regular

                # Remarks
                ws.cell(row=curr_row, column=12, value=st['remarks']).alignment = align_left

                # Apply borders & fills to row
                for c_k in range(2, 13):
                    cell = ws.cell(row=curr_row, column=c_k)
                    cell.border = thin_border
                    cell.fill = row_fill
                    if st.get('is_discontinued') and c_k != 6:
                        cell.font = Font(name="Calibri", size=10, italic=True, color="475569")

                ws.row_dimensions[curr_row].height = 20
                curr_row += 1

            curr_row += 1

        # Column widths
        ws.column_dimensions['B'].width = 6
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 28
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 28
        ws.column_dimensions['G'].width = 16
        ws.column_dimensions['H'].width = 14
        ws.column_dimensions['I'].width = 26
        ws.column_dimensions['J'].width = 26
        ws.column_dimensions['K'].width = 24
        ws.column_dimensions['L'].width = 34

    wb.save(output_path)
    print(f"Successfully generated complete AY Excel workbook: {output_path}")

    # Mirror to public folder and artifact folder
    pub_path = os.path.join(cur_dir, 'public', 'Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx')
    if os.path.exists(os.path.dirname(pub_path)):
        try:
            wb.save(pub_path)
        except Exception:
            pass

    art_path = r"C:\Users\HP\.gemini\antigravity\brain\e8668f28-7b4e-4d60-b0d6-089987194df7\Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx"
    if os.path.exists(os.path.dirname(art_path)):
        try:
            wb.save(art_path)
        except Exception:
            pass

    return True

if __name__ == '__main__':
    generate_ay_excel()
