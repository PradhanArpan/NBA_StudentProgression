import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

JSON_PATH = r"C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard\student_data.json"
EXCEL_PATH = r"C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard\Master_Student_Directory_Editable.xlsx"
PUBLIC_PATH = r"C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard\public\Master_Student_Directory_Editable.xlsx"
ARTIFACT_PATH = r"C:\Users\HP\.gemini\antigravity\brain\e8668f28-7b4e-4d60-b0d6-089987194df7\Master_Student_Directory_Editable.xlsx"

with open(JSON_PATH, 'r', encoding='utf-8') as f:
    db = json.load(f)

students = db['students']

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Master_Student_Directory"
ws.views.sheetView[0].showGridLines = True

# Title block
ws.merge_cells("A1:AJ1")
title_cell = ws["A1"]
title_cell.value = "DEPARTMENT OF CIVIL ENGINEERING — MASTER STUDENT DIRECTORY (EDITABLE)"
title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
title_cell.fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:AJ2")
sub_cell = ws["A2"]
sub_cell.value = "Instructions: Edit student progression details (Category, Progression Status, Discontinued Sem, Year-Back details) or Semester Marks. This sheet is the master truth that generates all Academic Year student lists."
sub_cell.font = Font(name="Calibri", size=10, italic=True, color="333333")
sub_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
sub_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 20

headers = [
    ("Batch", 10),
    ("Reg No", 14),
    ("Student Name", 28),
    ("Gender", 10),
    ("Domicile", 16),
    ("Category", 18),
    ("Progression Status", 20),
    ("Discontinued From Sem", 22),
    ("Year-Back Detained Sem", 22),
    ("Year-Back Junior Batch", 22),
    ("Readmission Reg No", 20),
    ("Overall CGPA", 14),
    ("Backlogs Count", 14),
    ("Grad Status", 22),
    ("S1 SGPA", 11), ("S1 %", 11), ("S1 Status", 11),
    ("S2 SGPA", 11), ("S2 %", 11), ("S2 Status", 11),
    ("S3 SGPA", 11), ("S3 %", 11), ("S3 Status", 11),
    ("S4 SGPA", 11), ("S4 %", 11), ("S4 Status", 11),
    ("S5 SGPA", 11), ("S5 %", 11), ("S5 Status", 11),
    ("S6 SGPA", 11), ("S6 %", 11), ("S6 Status", 11),
    ("S7 SGPA", 11), ("S7 %", 11), ("S7 Status", 11),
    ("S8 SGPA", 11), ("S8 %", 11), ("S8 Status", 11),
    ("Internship 1", 20),
    ("Internship 2", 20),
]

hdr_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

ws.row_dimensions[3].height = 26
for col_idx, (hdr_text, col_w) in enumerate(headers, start=1):
    c = ws.cell(row=3, column=col_idx, value=hdr_text)
    c.fill = hdr_fill
    c.font = hdr_font
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = thin_border
    col_letter = get_column_letter(col_idx)
    ws.column_dimensions[col_letter].width = col_w

sorted_students = sorted(students, key=lambda s: (int(s.get('batch', 0)), str(s.get('regNo', ''))))

row_num = 4
for s in sorted_students:
    ws.row_dimensions[row_num].height = 20
    
    prog = s.get('progressionStatus', 'Normal')
    disc_sem = s.get('discontinuedFromSem') or ''
    yb_sem = s.get('yearBackFromSem') or ''
    yb_batch = s.get('yearBackJuniorBatch') or ''
    readmit_reg = s.get('readmissionRegNo') or ''
    
    row_vals = [
        s.get('batch', ''),
        s.get('regNo', ''),
        s.get('name', ''),
        s.get('gender', ''),
        s.get('domicile', ''),
        s.get('category', 'Regular'),
        prog,
        disc_sem,
        yb_sem,
        yb_batch,
        readmit_reg,
        s.get('cgpa', ''),
        s.get('failCount', 0),
        s.get('gradStatus', ''),
    ]
    
    # S1-S8
    for i in range(1, 9):
        sm_key = f'sem{i}'
        sm_obj = s.get('semesters', {}).get(sm_key)
        if sm_obj and sm_obj.get('sgpa') is not None:
            row_vals.extend([
                sm_obj.get('sgpa', ''),
                sm_obj.get('pct', ''),
                sm_obj.get('status', 'Pass')
            ])
        else:
            row_vals.extend(['', '', '—'])
            
    row_vals.append(s.get('internship1', {}).get('company', '') if s.get('internship1') else '')
    row_vals.append(s.get('internship2', {}).get('company', '') if s.get('internship2') else '')
    
    # Alternating row fill
    row_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") if row_num % 2 == 0 else PatternFill(fill_type=None)
    
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws.cell(row=row_num, column=c_idx, value=val)
        cell.font = Font(name="Calibri", size=10)
        cell.border = thin_border
        cell.fill = row_fill
        
        # Alignments
        if c_idx in [1, 2, 4, 8, 9, 10, 11, 13, 14, 17, 20, 23, 26, 29, 32, 35]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx in [12, 15, 16, 18, 19, 21, 22, 24, 25, 27, 28, 30, 31, 33, 34]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
        # Highlight Discontinued / Year-Back rows gently
        if prog == 'Discontinued':
            cell.fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
        elif prog == 'Year-Back':
            cell.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
        elif prog == 'Readmitted':
            cell.fill = PatternFill(start_color="F5F3FF", end_color="F5F3FF", fill_type="solid")

    row_num += 1

# Auto-filter on header row
ws.auto_filter.ref = f"A3:AJ{row_num-1}"
ws.freeze_panes = "D4"

wb.save(EXCEL_PATH)
wb.save(PUBLIC_PATH)
try:
    wb.save(ARTIFACT_PATH)
except Exception as e:
    print(f"Artifact save note: {e}")

print(f"Successfully generated Master Student Directory with {len(students)} students at:")
print(f"  {EXCEL_PATH}")
