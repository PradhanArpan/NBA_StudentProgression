import json
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Load student database
json_path = r'C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard\student_data.json'
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

# Color styles
NAVY = "1F497D"
LIGHT_BLUE = "DCE6F1"
STEEL_BLUE = "2E75B6"
LIGHT_CYAN = "EDF2F8"
WHITE = "FFFFFF"
GRAY_HEADER = "595959"
LIGHT_GRAY = "F2F2F2"
BORDER_GRAY = "D9D9D9"

font_title = Font(name="Calibri", size=16, bold=True, color="1F497D")
font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
font_sec_hdr = Font(name="Calibri", size=12, bold=True, color=WHITE)
font_tbl_hdr = Font(name="Calibri", size=11, bold=True, color=WHITE)
font_bold = Font(name="Calibri", size=10, bold=True)
font_regular = Font(name="Calibri", size=10)
font_italic = Font(name="Calibri", size=9, italic=True, color="595959")

fill_navy = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
fill_light_blue = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
fill_light_gray = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
fill_light_cyan = PatternFill(start_color=LIGHT_CYAN, end_color=LIGHT_CYAN, fill_type="solid")

thin_border_side = Side(style='thin', color=BORDER_GRAY)
thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')

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

def get_ay_year_roster(ay_start, year_level):
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

        if year_level == 1 and cat == 'Lateral':
            continue
        if year_level == 1 and (prog == 'Readmitted' or 'Readmitted' in cat):
            continue

        # Discontinued: strictly exclude if beyond discontinuation semester
        if prog == 'Discontinued' and disc_sem:
            max_attended_year = math.ceil(disc_sem / 2)
            if year_level > max_attended_year:
                continue
            if disc_sem % 2 == 1 and (year_level * 2 - 1 > disc_sem):
                continue
        elif prog == 'Discontinued':
            last_sem = get_last_active_sem(s)
            if last_sem == 0 and year_level > 1:
                continue
            if last_sem > 0 and year_level > math.ceil((last_sem + 1) / 2):
                continue

        # Year-Back: student must NOT continue in the next year with original batch!
        if prog == 'Year-Back' or s.get('isYearBackExit'):
            detained_sem = yb_sem or 4
            detained_year = math.ceil(detained_sem / 2)
            if year_level > detained_year:
                continue

        sem_odd = s.get('semesters', {}).get(odd_key, {})
        sem_even = s.get('semesters', {}).get(even_key, {})

        odd_sgpa = sem_odd.get('sgpa')
        odd_pct = sem_odd.get('pct')
        odd_stat = sem_odd.get('status')

        even_sgpa = sem_even.get('sgpa')
        even_pct = sem_even.get('pct')
        even_stat = sem_even.get('status')

        if prog == 'Year-Back' and year_level == math.ceil((yb_sem or 4) / 2):
            status_desc = f"Year-Back (Detained at Sem {yb_sem or 4} ➔ Moves to Batch {s.get('yearBackJuniorBatch')})"
        elif prog == 'Discontinued' and disc_sem and year_level == math.ceil(disc_sem / 2):
            status_desc = f"Discontinued (Left from Sem {disc_sem})"
        elif odd_stat == 'Fail' or even_stat == 'Fail':
            status_desc = "With Backlog"
        elif odd_stat == 'Pass' and even_stat == 'Pass':
            status_desc = "Pass (All Cleared)"
        elif odd_stat == 'Pass' or even_stat == 'Pass':
            status_desc = "In Progress / Part Verified"
        elif 'Graduated (No Backlog)' in s.get('gradStatus', '') and year_level == 4:
            status_desc = "Graduated (Clear)"
        elif 'Graduated (With Backlog)' in s.get('gradStatus', '') and year_level == 4:
            status_desc = "Graduated (With Backlog)"
        else:
            status_desc = "Enrolled / Active"

        remarks = []
        if cat == 'Lateral':
            remarks.append("Lateral Entry in S3")
        if prog == 'Readmitted' or 'Readmitted' in cat:
            remarks.append("Readmitted in Year II after Year-Back")
        if prog == 'Year-Back':
            remarks.append(f"Year-Back from Sem {yb_sem or 4} ➔ Joins Batch {s.get('yearBackJuniorBatch')}")
        if prog == 'Discontinued':
            remarks.append(f"Discontinued from Sem {disc_sem or '—'}")
        if not remarks:
            remarks.append("Regular Admission")

        odd_str = f"SGPA {float(odd_sgpa):.2f} ({odd_pct}%)" if odd_sgpa else (odd_stat or "-")
        even_str = f"SGPA {float(even_sgpa):.2f} ({even_pct}%)" if even_sgpa else (even_stat or "-")

        roster.append({
            'regNo': s['regNo'],
            'name': s['name'],
            'gender': s.get('gender', '-'),
            'category': cat,
            'batch': f"Batch {base_batch}",
            'odd_sem': f"{odd_key.upper()}: {odd_str}",
            'even_sem': f"{even_key.upper()}: {even_str}",
            'status': status_desc,
            'remarks': "; ".join(remarks)
        })
        included_regnos.add(s['regNo'])

    # 2. Year-Back Students moving from Senior Batch into this Junior Batch
    # "If year back, their names should not continue in the next year with their batch but go to the junior batch, unless readmission record is there (for those already)"
    for s in students:
        prog = s.get('progressionStatus')
        target_batch = s.get('yearBackJuniorBatch')
        readmit_regno = s.get('readmissionRegNo')

        if (prog == 'Year-Back' or s.get('isYearBackExit')) and target_batch == base_batch:
            has_readmit_record = bool(readmit_regno and any(st['regNo'] == readmit_regno for st in students))
            if not has_readmit_record and s['regNo'] not in included_regnos:
                yb_sem = s.get('yearBackFromSem', 4)
                repeat_year = math.ceil(yb_sem / 2)
                if year_level >= repeat_year:
                    roster.append({
                        'regNo': s['regNo'],
                        'name': s['name'],
                        'gender': s.get('gender', '-'),
                        'category': 'Year-Back Repeat',
                        'batch': f"Batch {base_batch} (ex-Batch {s['batch']})",
                        'odd_sem': "Repeating Year",
                        'even_sem': "Repeating Year",
                        'status': f"Year-Back (Transferred from Batch {s['batch']} to Batch {base_batch})",
                        'remarks': f"Year-Back from Batch {s['batch']}; repeating in Batch {base_batch}"
                    })
                    included_regnos.add(s['regNo'])

    return roster

# -------------------------------------------------------------
# 1. SUMMARY SHEET
# -------------------------------------------------------------
ws_sum = wb.create_sheet(title="Executive Summary")
ws_sum.views.sheetView[0].showGridLines = True

ws_sum.cell(row=2, column=2, value="B.Tech Civil Engineering - Academic Year Cohort Distribution").font = font_title
ws_sum.cell(row=3, column=2, value="Official NBA Compliance & Department Student Progression Audit (AY 2020-21 to AY 2026-27)").font = font_subtitle

sum_headers = ["Academic Year", "NBA Status", "I Year (S1 & S2)", "II Year (S3 & S4)", "III Year (S5 & S6)", "IV Year (S7 & S8)", "Total AY Strength"]
for col_idx, h in enumerate(sum_headers, start=2):
    cell = ws_sum.cell(row=5, column=col_idx, value=h)
    cell.font = font_tbl_hdr
    cell.fill = fill_navy
    cell.alignment = align_center
    cell.border = thin_border
ws_sum.row_dimensions[5].height = 28

row_idx = 6
for ay_label, ay_start, ay_status in academic_years:
    r1 = get_ay_year_roster(ay_start, 1)
    r2 = get_ay_year_roster(ay_start, 2)
    r3 = get_ay_year_roster(ay_start, 3)
    r4 = get_ay_year_roster(ay_start, 4)
    tot = len(r1) + len(r2) + len(r3) + len(r4)

    ws_sum.cell(row=row_idx, column=2, value=ay_label).font = font_bold
    ws_sum.cell(row=row_idx, column=2).alignment = align_center
    ws_sum.cell(row=row_idx, column=2).border = thin_border

    ws_sum.cell(row=row_idx, column=3, value=ay_status).font = font_regular
    ws_sum.cell(row=row_idx, column=3).alignment = align_left
    ws_sum.cell(row=row_idx, column=3).border = thin_border

    for c_idx, val in enumerate([len(r1), len(r2), len(r3), len(r4), tot], start=4):
        cell = ws_sum.cell(row=row_idx, column=c_idx, value=val)
        cell.font = font_bold if c_idx == 8 else font_regular
        cell.alignment = align_center
        cell.border = thin_border
        if c_idx == 8:
            cell.fill = fill_light_cyan

    ws_sum.row_dimensions[row_idx].height = 22
    row_idx += 1

# Total Row
ws_sum.cell(row=row_idx, column=2, value="Grand Total Student-Years").font = font_bold
ws_sum.cell(row=row_idx, column=2).alignment = align_center
ws_sum.cell(row=row_idx, column=2).border = thin_border
ws_sum.cell(row=row_idx, column=3, value="Cumulative Over 7 AYs").font = font_italic
ws_sum.cell(row=row_idx, column=3).border = thin_border

for c_idx, col_letter in enumerate(['D', 'E', 'F', 'G', 'H'], start=4):
    cell = ws_sum.cell(row=row_idx, column=c_idx, value=f"=SUM({col_letter}6:{col_letter}{row_idx-1})")
    cell.font = font_bold
    cell.alignment = align_center
    cell.border = thin_border
    cell.fill = fill_light_blue
ws_sum.row_dimensions[row_idx].height = 24

# Explanation notes
note_row = row_idx + 3
notes = [
    ("NOTE 1 (Cohort Dynamics):", "In an academic year (e.g. AY 2023-24), students in Year I belong to Batch 2023, Year II to Batch 2022, Year III to Batch 2021, and Year IV to Batch 2020."),
    ("NOTE 2 (Discontinued Students):", "Students who discontinued / transferred out are strictly included ONLY during the specific academic years/semesters they were actively enrolled. Once discontinued from a semester, they are purged from all subsequent study years (e.g., Nandana J Nair is in Y1 & Y2, but excluded from Y3 & Y4)."),
    ("NOTE 3 (Year-Back Students):", "If a student suffers a year-back, their name does NOT continue in the next year with their senior batch. Instead, they move to the junior batch. If an official readmission record under a new registration number already exists, that record represents them without duplication."),
    ("NOTE 4 (Lateral Entry):", "Lateral Entry students enter the program directly in Year II (Semester 3). Hence, they are included in II, III, and IV Year rosters, but NOT in I Year."),
    ("NOTE 5 (Zero-Hypothetical Policy):", "Only verified students and exam results from the official COE XML spreadsheets and verified university marks are presented. Pre-2020 senior batches are left unstated.")
]
for title, desc in notes:
    ws_sum.cell(row=note_row, column=2, value=title).font = font_bold
    ws_sum.cell(row=note_row, column=3, value=desc).font = font_regular
    ws_sum.merge_cells(start_row=note_row, start_column=3, end_row=note_row, end_column=8)
    note_row += 1

# -------------------------------------------------------------
# 2. INDIVIDUAL ACADEMIC YEAR SHEETS
# -------------------------------------------------------------
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
    ws.cell(row=3, column=2, value=f"Academic Classification: {ay_status} | Complete List Across I, II, III, and IV Years").font = font_subtitle

    curr_row = 5

    for yr_lvl, yr_name, yr_sub, color_hex in year_titles:
        roster = get_ay_year_roster(ay_start, yr_lvl)
        cohort_b = str(ay_start - (yr_lvl - 1))

        ws.cell(row=curr_row, column=2, value=f"{yr_name} — Batch {cohort_b} (Total: {len(roster)} Students)").font = font_sec_hdr
        ws.cell(row=curr_row, column=2).fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        ws.cell(row=curr_row, column=2).alignment = align_left
        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=10)
        ws.row_dimensions[curr_row].height = 26
        curr_row += 1

        if len(roster) == 0:
            ws.cell(row=curr_row, column=2, value=f"No student records in repository for Batch {cohort_b} (Pre-2020 cohort not covered in 36 COE XML dataset).").font = font_italic
            ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=10)
            ws.row_dimensions[curr_row].height = 20
            curr_row += 2
            continue

        cols = ["Sl", "Reg. No.", "Student Name", "Gender", "Category", "Batch", "Odd Semester", "Even Semester", "Academic Status"]
        for c_i, col_name in enumerate(cols, start=2):
            c_cell = ws.cell(row=curr_row, column=c_i, value=col_name)
            c_cell.font = font_tbl_hdr
            c_cell.fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
            c_cell.alignment = align_center if c_i in [2, 3, 5, 6, 7] else align_left
            c_cell.border = thin_border
        ws.row_dimensions[curr_row].height = 22
        curr_row += 1

        for s_idx, st in enumerate(roster, start=1):
            ws.cell(row=curr_row, column=2, value=s_idx).alignment = align_center
            ws.cell(row=curr_row, column=3, value=st['regNo']).alignment = align_center
            ws.cell(row=curr_row, column=3).font = font_bold
            ws.cell(row=curr_row, column=4, value=st['name']).alignment = align_left
            ws.cell(row=curr_row, column=4).font = font_bold
            ws.cell(row=curr_row, column=5, value=st['gender']).alignment = align_center
            ws.cell(row=curr_row, column=6, value=st['category']).alignment = align_center
            ws.cell(row=curr_row, column=7, value=st['batch']).alignment = align_center
            ws.cell(row=curr_row, column=8, value=st['odd_sem']).alignment = align_left
            ws.cell(row=curr_row, column=9, value=st['even_sem']).alignment = align_left
            ws.cell(row=curr_row, column=10, value=st['status']).alignment = align_center

            status_cell = ws.cell(row=curr_row, column=10)
            if "Backlog" in st['status'] or "Discontinued" in st['status']:
                status_cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
            elif "Pass" in st['status'] or "Graduated" in st['status']:
                status_cell.font = Font(name="Calibri", size=10, bold=True, color="008000")
            elif "Year-Back" in st['status']:
                status_cell.font = Font(name="Calibri", size=10, bold=True, color="B25900")
            else:
                status_cell.font = font_regular

            for c_i in range(2, 11):
                ws.cell(row=curr_row, column=c_i).border = thin_border
                if s_idx % 2 == 0:
                    ws.cell(row=curr_row, column=c_i).fill = fill_light_gray

            ws.row_dimensions[curr_row].height = 20
            curr_row += 1

        curr_row += 2

for sheet in wb.worksheets:
    sheet.column_dimensions['A'].width = 3
    sheet.column_dimensions['B'].width = 8
    sheet.column_dimensions['C'].width = 16
    sheet.column_dimensions['D'].width = 34
    sheet.column_dimensions['E'].width = 10
    sheet.column_dimensions['F'].width = 24
    sheet.column_dimensions['G'].width = 14
    sheet.column_dimensions['H'].width = 26
    sheet.column_dimensions['I'].width = 26
    sheet.column_dimensions['J'].width = 36

out_path1 = r'C:\Users\HP\.gemini\antigravity\scratch\Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx'
out_path2 = r'C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard\Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx'
out_path3 = r'C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard\public\Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx'
out_path4 = r'C:\Users\HP\.gemini\antigravity\brain\e8668f28-7b4e-4d60-b0d6-089987194df7\Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx'

wb.save(out_path1)
wb.save(out_path2)
wb.save(out_path3)
wb.save(out_path4)
print("Updated all 4 copies of Excel workbook successfully!")
