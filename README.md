# NBA Criterion 4 & Student Progression Analytics Dashboard
### Department of Civil Engineering · CHRIST (Deemed to be University)

An interactive NBA Tier-I Pre-Qualifier & SAR analytics portal and Student Progression Information System for Criterion 4 (Students' Performance — 120 Marks), P4, and P5.

---

## ✨ Features
* **Criteria 4 Analytics**: Live computation for Sub-Criteria 4.1 to 4.7 (Enrolment Ratio, Success Rate, Academic Performance Index 1st/2nd/3rd Year, Placement & Higher Studies, and Professional Activities).
* **P-4 Student Progression**: Automated summary tables and complete student rosters for graduation without backlogs and with backlogs.
* **P-5 Division of Class**: 8-semester division of class performance matrix (Distinction, First Class, Second Class, Pass, Fail).
* **Student Directory & Manager**: Batchwise directory across 7 cohorts (2020 to 2026) with 8 semesters of SGPA/CGPA tracking, backlogs, domicile, and admission category.
* **Dual Internship & Placement File Attachments**:
  * Internship 1 (Site & Material QA/QC)
  * Internship 2 (Design & Capstone Analysis)
  * Placement & Higher Studies offer/admission letters
* **Automated Batchwise Document Organization**: Structured storage under `documents/Batch_{YYYY}/{Internship_1, Internship_2, Placement}/`.
* **Download Hub**: Clean Word/Excel templates and pre-filled completed Pre-Qualifier documents.

---

## 🚀 Quick Start (Local)

1. **Clone the repository**:
   ```bash
   git clone https://github.com/PradhanArpan/NBA_StudentProgression.git
   cd NBA_StudentProgression
   ```

2. **Run the server**:
   ```bash
   python server.py
   ```

3. **Open in Browser**:
   Open [http://localhost:5173](http://localhost:5173)

---

## ☁️ Deploy 24/7 on Render / Railway

* **Runtime**: Python 3
* **Start Command**: `python server.py`
* **Port**: Dynamic `$PORT` supported automatically
