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
    '/download/student-data-json': DATA_FILE,
}


class CustomHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_GET(self):
        clean_path = self.path.split('?')[0]
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
                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
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
