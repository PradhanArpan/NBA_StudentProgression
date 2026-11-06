import http.server
import socketserver
import os
import json

PORT = int(os.environ.get('PORT', 5173))
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
NBA_DIR = r"C:\Users\HP\Downloads\NBA"
DATA_FILE = os.path.join(DIRECTORY, 'student_data.json')

def get_doc_path(filename):
    local_path = os.path.join(DIRECTORY, filename)
    if os.path.exists(local_path):
        return local_path
    alt_path = os.path.join(NBA_DIR, filename)
    if os.path.exists(alt_path):
        return alt_path
    return local_path

DOWNLOADS = {
    '/download-word': get_doc_path('Criterion-4_Completed_PreQualifier.docx'),
    '/download/criterion4-completed': get_doc_path('Criterion-4_Completed_PreQualifier.docx'),
    '/download/criterion4-template': get_doc_path('Criterion-4.docx'),
    '/download/p4-completed': os.path.join(DIRECTORY, 'P4_Student_Progression_Completed.xlsx'),
    '/download/p4-template': os.path.join(DIRECTORY, 'P4_Student_Progression_Template.xlsx'),
    '/download/p5-completed': os.path.join(DIRECTORY, 'P5_Result_Classification_Completed.xlsx'),
    '/download/p5-template': os.path.join(DIRECTORY, 'P5_Result_Classification_Template.xlsx'),
    '/download/master-excel': os.path.join(DIRECTORY, 'NBA_Civil_Criterion4_P4_P5_Master.xlsx'),
    '/download/master-student-directory': os.path.join(DIRECTORY, 'Master_Student_Directory_Editable.xlsx'),
    '/download/ay-student-rosters': os.path.join(DIRECTORY, 'Civil_Engineering_Academic_Year_Student_Lists_2020_to_2027.xlsx'),
    '/download/student-data-json': DATA_FILE,
}


class CustomHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_GET(self):
        clean_path = self.path.split('?')[0]
        # Never cache student_data.json
        if clean_path in ['/student_data.json', '/student_data.json/']:
            if os.path.exists(DATA_FILE):
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                self.send_header('Pragma', 'no-cache')
                self.send_header('Expires', '0')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Length', str(os.path.getsize(DATA_FILE)))
                self.end_headers()
                with open(DATA_FILE, 'rb') as f:
                    self.wfile.write(f.read())
                return

        if clean_path in DOWNLOADS:
            fpath = DOWNLOADS[clean_path]
            # Fallback to local directory if not found in NBA_DIR
            if not os.path.exists(fpath):
                local_fpath = os.path.join(DIRECTORY, os.path.basename(fpath))
                if os.path.exists(local_fpath):
                    fpath = local_fpath

            if os.path.exists(fpath):
                filename = os.path.basename(fpath)
                if filename.endswith('.docx'):
                    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                elif filename.endswith('.xlsx'):
                    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                elif filename.endswith('.json'):
                    mime = 'application/json'
                else:
                    mime = 'application/octet-stream'

                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(os.path.getsize(fpath)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                with open(fpath, 'rb') as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_error(404, f"File {os.path.basename(fpath)} not found")
                return

        # Serve files from documents folder
        if clean_path.startswith('/documents/'):
            doc_rel = clean_path[len('/documents/'):].replace('/', os.sep)
            fpath = os.path.join(DIRECTORY, 'documents', doc_rel)
            if not os.path.exists(fpath):
                # Check Downloads folder
                alt_fpath = os.path.join(NBA_DIR, 'documents', doc_rel)
                if os.path.exists(alt_fpath):
                    fpath = alt_fpath

            if os.path.exists(fpath):
                filename = os.path.basename(fpath)
                if filename.lower().endswith('.pdf'):
                    mime = 'application/pdf'
                elif filename.lower().endswith(('.jpg', '.jpeg')):
                    mime = 'image/jpeg'
                elif filename.lower().endswith('.png'):
                    mime = 'image/png'
                elif filename.lower().endswith('.docx'):
                    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                else:
                    mime = 'application/octet-stream'

                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'inline; filename="{filename}"')
                self.send_header('Content-Length', str(os.path.getsize(fpath)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                with open(fpath, 'rb') as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_error(404, f"Document {doc_rel} not found")
                return

        return super().do_GET()

    def do_POST(self):
        if self.path == '/save-data':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(content_length)
                data = json.loads(body.decode('utf-8'))
                
                # Automated rolling backup before saving
                backup_dir = os.path.join(DIRECTORY, 'backups')
                os.makedirs(backup_dir, exist_ok=True)
                import time, shutil
                backup_file = os.path.join(backup_dir, f'student_data_backup_{int(time.time())}.json')
                if os.path.exists(DATA_FILE):
                    shutil.copy2(DATA_FILE, backup_file)

                # Write to primary data file
                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                
                # Also mirror to public folder if exists
                pub_file = os.path.join(DIRECTORY, 'public', 'student_data.json')
                if os.path.exists(os.path.dirname(pub_file)):
                    with open(pub_file, 'w', encoding='utf-8') as f:
                        json.dump(data, f, indent=2, ensure_ascii=False)
                
                # Auto refresh both Master Directory & AY Roster Excel files
                try:
                    import subprocess
                    script1 = os.path.join(DIRECTORY, 'generate_excel_ay_rosters.py')
                    script2 = os.path.join(DIRECTORY, 'generate_master_student_directory.py')
                    if os.path.exists(script1):
                        subprocess.Popen(['python', script1])
                    if os.path.exists(script2):
                        subprocess.Popen(['python', script2])
                except Exception as ex:
                    print("Background excel refresh error:", ex)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(b'{"status":"ok"}')
            except Exception as e:
                self.send_error(500, str(e))
            return

        elif self.path == '/upload-document':
            import base64
            import re
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode('utf-8'))
                
                batch = str(payload.get('batch', '2022'))
                reg_no = str(payload.get('regNo', 'unknown'))
                student_name = payload.get('studentName', '')
                clean_name = re.sub(r'[^a-zA-Z0-9]', '_', student_name.strip()).strip('_')[:30]
                doc_type = payload.get('docType', 'Internship_1') # 'Internship_1', 'Internship_2', 'Placement'
                original_file_name = payload.get('fileName', 'document.pdf')
                ext = os.path.splitext(original_file_name)[1].lower() or '.pdf'
                
                base64_data = payload.get('fileBase64', '')
                if ',' in base64_data:
                    base64_data = base64_data.split(',', 1)[1]
                file_bytes = base64.b64decode(base64_data)

                # Target filename
                target_filename = f"{reg_no}_{clean_name}_{doc_type}{ext}"
                rel_url = f"/documents/Batch_{batch}/{doc_type}/{target_filename}"

                # Write to scratch documents
                target_dir = os.path.join(DIRECTORY, 'documents', f"Batch_{batch}", doc_type)
                os.makedirs(target_dir, exist_ok=True)
                full_scratch_path = os.path.join(target_dir, target_filename)
                with open(full_scratch_path, 'wb') as f:
                    f.write(file_bytes)

                # Also write to Downloads/NBA/documents
                nba_target_dir = os.path.join(NBA_DIR, 'documents', f"Batch_{batch}", doc_type)
                os.makedirs(nba_target_dir, exist_ok=True)
                full_nba_path = os.path.join(nba_target_dir, target_filename)
                with open(full_nba_path, 'wb') as f:
                    f.write(file_bytes)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                resp = json.dumps({"status": "ok", "filePath": rel_url, "fileName": target_filename})
                self.wfile.write(resp.encode('utf-8'))
            except Exception as e:
                self.send_error(500, str(e))
            return
        self.send_error(404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        print(f"  {self.address_string()} - {format % args}")


if __name__ == '__main__':
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.ThreadingTCPServer(("", PORT), CustomHandler) as httpd:
        print(f"\n  NBA Criterion 4, P4 & P5 Dashboard")
        print(f"  Running at http://localhost:{PORT}")
        print(f"  Template & Completed downloads configured.\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
