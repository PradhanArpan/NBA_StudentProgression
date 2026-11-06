import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(DIRECTORY, "student_data.json")
EXCEL_PATH = os.path.join(DIRECTORY, "Master_Student_Directory_Editable.xlsx")
PUBLIC_PATH = os.path.join(DIRECTORY, "public", "Master_Student_Directory_Editable.xlsx")
ARTIFACT_DIR = r"C:\Users\HP\.gemini\antigravity\brain\e8668f28-7b4e-4d60-b0d6-089987194df7"

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

def generate_master_excel(json_file=None, output_file=None):
    jpath = json_file or JSON_PATH
    out_path = output_file or EXCEL_PATH
    
    if not os.path.exists(jpath):
        print(f"Error: student_data.json not found at {jpath}")
        return False
        
    with open(jpath, 'r', encoding='utf-8') as f:
        db = json.load(f)

    students = db.get('students', [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master_Student_Directory"
    ws.views.sheetView[0].showGridLines = True

    # Title block
    ws.merge_cells("A1:AK1")
    title_cell = ws["A1"]
    title_cell.value = "DEPARTMENT OF CIVIL ENGINEERING — MASTER STUDENT DIRECTORY (EDITABLE)"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:AK2")
    sub_cell = ws["A2"]
    sub_cell.value = "Instructions: Edit student progression details (Category, Progression Status, Discontinued Sem, Year-Back details) or Semester Marks. This sheet is the master truth that generates all Academic Year student lists with official NBA tags (N1 to N6, YB)."
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="333333")
    sub_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    headers = [
        ("Batch", 10),
        ("Reg No", 14),
        ("Student Name", 28),
        ("Gender", 10),
        ("NBA Category Tag", 28),
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

    # NBA Tag Font Colors
    font_tag_n1 = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    font_tag_n2 = Font(name="Calibri", size=10, bold=True, color="065F46")
    font_tag_n5 = Font(name="Calibri", size=10, bold=True, color="581C87")
    font_tag_n6 = Font(name="Calibri", size=10, bold=True, color="991B1B")
    font_tag_yb = Font(name="Calibri", size=10, bold=True, color="92400E")

    fill_disc_row = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid")
    fill_yb_row = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    fill_alt_row = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    row_num = 4
    for s_idx, s in enumerate(sorted_students, start=1):
        ws.row_dimensions[row_num].height = 20
        
        prog = s.get('progressionStatus', 'Normal')
        disc_sem = s.get('discontinuedFromSem') or ''
        yb_sem = s.get('yearBackFromSem') or ''
        yb_batch = s.get('yearBackJuniorBatch') or ''
        readmit_reg = s.get('readmissionRegNo') or ''
        nba_tag = get_nba_tag(s)

        row_fill = fill_white
        if prog == 'Discontinued' or s.get('gradStatus') == 'Discontinued / Transferred':
            row_fill = fill_disc_row
        elif prog == 'Year-Back' or s.get('isYearBackExit'):
            row_fill = fill_yb_row
        elif s_idx % 2 == 0:
            row_fill = fill_alt_row
        
        row_vals = [
            s.get('batch', ''),
            s.get('regNo', ''),
            s.get('name', ''),
            s.get('gender', ''),
            nba_tag,
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

        # Add S1-S8 values
        for sm_idx in range(1, 9):
            sm_data = s.get('semesters', {}).get(f'sem{sm_idx}', {})
            sgpa = sm_data.get('sgpa')
            pct = sm_data.get('pct')
            stat = sm_data.get('status', '—')
            row_vals.extend([
                sgpa if sgpa is not None else '',
                pct if pct is not None else '',
                stat if stat else '—'
            ])

        # Internships
        row_vals.append(s.get('internship1', {}).get('company', '') if isinstance(s.get('internship1'), dict) else '')
        row_vals.append(s.get('internship2', {}).get('company', '') if isinstance(s.get('internship2'), dict) else '')

        for col_idx, val in enumerate(row_vals, start=1):
            c = ws.cell(row=row_num, column=col_idx, value=val)
            c.border = thin_border
            c.fill = row_fill
            if col_idx in [1, 2, 4, 8, 9, 10, 11, 12, 13, 14, 15] or (col_idx >= 16 and (col_idx - 16) % 3 != 1):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
            # Tag font styling
            if col_idx == 5:
                if "N1" in str(val):
                    c.font = font_tag_n1
                elif "N2" in str(val):
                    c.font = font_tag_n2
                elif "N5" in str(val):
                    c.font = font_tag_n5
                elif "N6" in str(val):
                    c.font = font_tag_n6
                elif "YB" in str(val):
                    c.font = font_tag_yb
            elif col_idx in [2, 3]:
                c.font = Font(name="Calibri", size=10, bold=True)
            elif prog == 'Discontinued':
                c.font = Font(name="Calibri", size=10, italic=True, color="475569")
            else:
                c.font = Font(name="Calibri", size=10)

        row_num += 1

    wb.save(out_path)
    print(f"Successfully generated Master Student Directory: {out_path}")

    # Mirror to public and artifact directories
    if os.path.exists(os.path.dirname(PUBLIC_PATH)):
        try:
            wb.save(PUBLIC_PATH)
        except Exception:
            pass

    art_file = os.path.join(ARTIFACT_DIR, "Master_Student_Directory_Editable.xlsx")
    if os.path.exists(ARTIFACT_DIR):
        try:
            wb.save(art_file)
        except Exception:
            pass

    return True

if __name__ == '__main__':
    generate_master_excel()
