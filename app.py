from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import os, time, subprocess, gc
from aggregate_engine import process_and_merge
from aggregate_engine_monthly import process_and_merge_monthly

app = Flask(__name__)
app.secret_key = 'super_secret_antigravity_key'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = '/tmp/uploads'
TEMPLATE_FILE = os.path.join(BASE_DIR, '00. 주간경영회의(20260410)_작성예정.docx')
OUTPUT_FILE = '/tmp/weekly_output.docx'

TEMPLATE_FILE_MONTHLY = os.path.join(BASE_DIR, '00.월간 업무 보고(대표이사님 보고용)_20260403 - 종합.docx')
OUTPUT_FILE_MONTHLY = '/tmp/monthly_output.docx'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def clean_uploads():
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], capture_output=True)
        time.sleep(0.5)
    except Exception:
        pass
    try:
        for f in os.listdir(UPLOAD_FOLDER):
            try: os.remove(os.path.join(UPLOAD_FOLDER, f))
            except: pass
    except Exception:
        pass


import traceback
from werkzeug.exceptions import HTTPException

@app.errorhandler(Exception)
def handle_exception(e):
    if isinstance(e, HTTPException):
        return e
    return f"<h1>System Error</h1><pre>{traceback.format_exc()}</pre>", 500

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        print("POST request received")
        clean_uploads()
        print("Uploads cleaned")
        
        files = request.files.getlist('branch_files')
        print(f"Number of files received: {len(files)}")
        uploaded_count = 0
        has_doc_error = False
        for file in files:
            if file.filename and not file.filename.startswith('~'):
                if file.filename.lower().endswith('.doc'):
                    flash(f"'{file.filename}' 파일은 구형(.doc) 포맷이라 취합할 수 없습니다. 워드에서 '다른 이름으로 저장'을 눌러 .docx 형식으로 변경 후 업로드해주세요!", "danger")
                    has_doc_error = True
                    continue
                clean_name = file.filename
                os.makedirs(UPLOAD_FOLDER, exist_ok=True)
                dest = os.path.join(UPLOAD_FOLDER, clean_name)
                file.save(dest)
                uploaded_count += 1
                
        if has_doc_error:
            return redirect(request.url)
                
        if uploaded_count == 0:
            flash("업로드된 지점 파일이 없습니다.", "warning")
            return redirect(request.url)
        
        # Force GC and wait for file handles to release
        gc.collect()
        time.sleep(1)
            
        report_type = request.form.get('report_type', 'weekly')
        
        try:
            expected = {f"{i:02d}" for i in range(1, 18)}
            uploaded = set(f[:2] for f in os.listdir(UPLOAD_FOLDER) if not f.startswith('~'))
            missing = expected - uploaded
            
            if report_type == 'monthly':
                process_and_merge_monthly(TEMPLATE_FILE_MONTHLY, UPLOAD_FOLDER, OUTPUT_FILE_MONTHLY)
                return send_file(OUTPUT_FILE_MONTHLY, as_attachment=True, download_name="00. 월간보고서_통합_작성완료.docx")
            else:
                process_and_merge(TEMPLATE_FILE, UPLOAD_FOLDER, OUTPUT_FILE)
                return send_file(OUTPUT_FILE, as_attachment=True, download_name="00. 주간경영회의_통합_작성완료.docx")
        except Exception as e:
            flash(f"처리 중 오류 발생: {str(e)}", "danger")
            return redirect(request.url)

    return render_template('index.html', download_ready=False, active_tab=None)

@app.route('/download')
def download():
    rtype = request.args.get('type', 'weekly')
    if rtype == 'monthly':
        response = send_file(OUTPUT_FILE_MONTHLY, as_attachment=True, download_name="00. 월간경영회의_최종_작성완료.docx")
    else:
        response = send_file(OUTPUT_FILE, as_attachment=True, download_name="00. 주간경영회의_최종_작성완료.docx")
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

if __name__ == '__main__':
    # Try to use waitress for a more stable experience, fallback to flask dev server
    try:
        from waitress import serve
        import socket
        
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        
        print(f"\n" + "="*50)
        print(f"서버가 성공적으로 시작되었습니다!")
        print(f"1. 내 PC에서 접속: http://localhost:3000")
        print(f"2. 같은 네트워크의 다른 PC에서 접속: http://{local_ip}:3000")
        print("="*50 + "\n")
        
        serve(app, host='0.0.0.0', port=3000)
    except ImportError:
        print("\nWaitress가 설치되지 않아 Flask 개발 서버로 실행합니다.")
        app.run(host='0.0.0.0', port=3000, debug=True)
