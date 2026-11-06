import json
import os
import subprocess
import openpyxl

DIRECTORY = r"C:\Users\HP\.gemini\antigravity\scratch\nba-criteria-dashboard"
EXCEL_PATH = os.path.join(DIRECTORY, "Master_Student_Directory_Editable.xlsx")
JSON_PATH = os.path.join(DIRECTORY, "student_data.json")

def sync_excel_to_json():
    if not os.path.exists(EXCEL_PATH):
        print(f"Error: Excel file not found at {EXCEL_PATH}")
        return False
        
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        db = json.load(f)
        
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["Master_Student_Directory"]
    
    student_map = {s['regNo']: s for s in db['students']}
    
    updated_count = 0
    # Data starts on row 4
    for r in range(4, ws.max_row + 1):
        reg_no = str(ws.cell(row=r, column=2).value or '').strip()
        if not reg_no:
            continue
            
        s = student_map.get(reg_no)
        if not s:
            batch_val = str(ws.cell(row=r, column=1).value or '2026').strip()
            s = {
                "id": f"CIVIL_{batch_val}_{reg_no}",
                "regNo": reg_no,
                "batch": batch_val,
                "department": "Civil Engineering",
                "semesters": {},
                "placement": {"status": "Ongoing"},
                "internship1": {},
                "internship2": {}
            }
            db['students'].append(s)
            student_map[reg_no] = s
            
        s['batch'] = str(ws.cell(row=r, column=1).value or s.get('batch', '')).strip()
        s['name'] = str(ws.cell(row=r, column=3).value or s.get('name', '')).strip()
        s['gender'] = str(ws.cell(row=r, column=4).value or s.get('gender', '')).strip()
        s['domicile'] = str(ws.cell(row=r, column=5).value or s.get('domicile', '')).strip()
        s['category'] = str(ws.cell(row=r, column=6).value or 'Regular').strip()
        
        prog = str(ws.cell(row=r, column=7).value or 'Normal').strip()
        s['progressionStatus'] = prog
        
        disc_val = ws.cell(row=r, column=8).value
        s['discontinuedFromSem'] = int(disc_val) if disc_val and str(disc_val).strip() != '' else None
        
        yb_sem_val = ws.cell(row=r, column=9).value
        s['yearBackFromSem'] = int(yb_sem_val) if yb_sem_val and str(yb_sem_val).strip() != '' else None
        
        yb_batch_val = ws.cell(row=r, column=10).value
        s['yearBackJuniorBatch'] = str(yb_batch_val).strip() if yb_batch_val else None
        
        readmit_reg_val = ws.cell(row=r, column=11).value
        s['readmissionRegNo'] = str(readmit_reg_val).strip() if readmit_reg_val else None
        
        cgpa_val = ws.cell(row=r, column=12).value
        if cgpa_val is not None and str(cgpa_val).strip() != '':
            try:
                s['cgpa'] = round(float(cgpa_val), 2)
            except:
                pass
                
        fail_val = ws.cell(row=r, column=13).value
        if fail_val is not None and str(fail_val).strip() != '':
            try:
                s['failCount'] = int(fail_val)
            except:
                pass
                
        grad_status_val = ws.cell(row=r, column=14).value
        if grad_status_val:
            s['gradStatus'] = str(grad_status_val).strip()
            
        # S1 to S8
        col_offset = 15
        for sem_idx in range(1, 9):
            sm_key = f'sem{sem_idx}'
            sgpa_v = ws.cell(row=r, column=col_offset).value
            pct_v = ws.cell(row=r, column=col_offset + 1).value
            st_v = ws.cell(row=r, column=col_offset + 2).value
            
            if sgpa_v is not None and str(sgpa_v).strip() != '':
                try:
                    sgpa_f = round(float(sgpa_v), 2)
                    pct_f = round(float(pct_v), 2) if (pct_v is not None and str(pct_v).strip() != '') else round(sgpa_f * 9.5, 2)
                    st_str = str(st_v).strip() if st_v else 'Pass'
                    cls_str = 'Fail' if st_str == 'Fail' else ('FC-D' if sgpa_f >= 7.5 else 'FC')
                    
                    if sm_key not in s['semesters'] or not s['semesters'][sm_key]:
                        s['semesters'][sm_key] = {}
                        
                    s['semesters'][sm_key]['sgpa'] = sgpa_f
                    s['semesters'][sm_key]['pct'] = pct_f
                    s['semesters'][sm_key]['class'] = cls_str
                    s['semesters'][sm_key]['status'] = st_str
                except Exception as ex:
                    pass
            elif st_v == '—' or not sgpa_v:
                pass
                
            col_offset += 3
            
        updated_count += 1
        
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(db, f, indent=2, ensure_ascii=False)
        
    print(f"Successfully synced {updated_count} student records from Excel to student_data.json.")
    
    # Regenerate the AY Excel lists
    ay_script = os.path.join(DIRECTORY, "generate_excel_ay_rosters.py")
    if os.path.exists(ay_script):
        subprocess.run(["python", ay_script], check=True)
        print("Regenerated Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx.")
        
    return True

if __name__ == '__main__':
    sync_excel_to_json()
