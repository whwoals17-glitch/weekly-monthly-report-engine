from flask import Flask, request, render_template, send_file, flash, redirect, url_for, jsonify
import os, time, subprocess, gc, zipfile, io
from aggregate_engine import process_and_merge
from aggregate_engine_monthly import process_and_merge_monthly
from filename_normalizer import normalize_filename, batch_rename_in_folder, convert_doc_to_docx, BRANCH_DEFINITIONS

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
        
        report_type = request.form.get('report_type', 'weekly')
        files = request.files.getlist('branch_files')
        print(f"Number of files received: {len(files)}")
        uploaded_count = 0
        has_doc_error = False
        unrecognized_files = []
        
        for file in files:
            if file.filename and not file.filename.startswith('~'):
                os.makedirs(UPLOAD_FOLDER, exist_ok=True)
                temp_dest = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(temp_dest)
                
                # 삼화타워, 판교 등 구형(.doc) 포맷이면 자동으로 .docx로 변환!
                if file.filename.lower().endswith('.doc'):
                    try:
                        temp_dest = convert_doc_to_docx(temp_dest, remove_original=True)
                        print(f"[자동 포맷 변환] {file.filename} (.doc -> .docx 완료)")
                    except Exception as e:
                        flash(f"'{file.filename}' 파일의 .docx 자동 변환에 실패했습니다 ({e}). 워드에서 '다른 이름으로 저장'으로 .docx 변경 후 업로드해주세요.", "danger")
                        has_doc_error = True
                        continue
                
                # 파일명이 제각각이거나 번호가 틀려도 자동으로 표준 파일명으로 변환
                b_info, std_name, _ = normalize_filename(os.path.basename(temp_dest), temp_dest, force_report_type=report_type, convert_doc=True)
                if b_info:
                    final_dest = os.path.join(UPLOAD_FOLDER, std_name)
                    if os.path.abspath(temp_dest) != os.path.abspath(final_dest):
                        if os.path.exists(final_dest):
                            try: os.remove(final_dest)
                            except: pass
                        os.rename(temp_dest, final_dest)
                    print(f"[자동 파일명 변환] {file.filename}  -->  {std_name}")
                else:
                    unrecognized_files.append(file.filename)
                    
                uploaded_count += 1
                
        if has_doc_error:
            return redirect(request.url)
                
        if uploaded_count == 0:
            flash("업로드된 지점 파일이 없습니다.", "warning")
            return redirect(request.url)
        
        # Force GC and wait for file handles to release
        gc.collect()
        time.sleep(1)
            
        try:
            expected = {f"{i:02d}" for i in range(1, 17)}
            uploaded = set(f[:2] for f in os.listdir(UPLOAD_FOLDER) if not f.startswith('~') and f[:2].isdigit())
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

    return render_template('index.html', download_ready=False, active_tab=None, branches=BRANCH_DEFINITIONS)


@app.route('/api/preview_rename', methods=['POST'])
def api_preview_rename():
    """파일명들을 받아 표준 파일명 변환 결과를 미리보기로 반환합니다."""
    data = request.get_json(force=True, silent=True) or {}
    filenames = data.get('filenames', [])
    report_type = data.get('report_type', None)
    target_date = data.get('target_date', None)
    
    # 다수결 날짜 계산
    from filename_normalizer import extract_date_from_filename_or_doc, detect_report_type
    detected_dates = [extract_date_from_filename_or_doc(fn) for fn in filenames if fn]
    common_date = None
    if target_date and len(target_date.strip()) == 8 and target_date.strip().isdigit():
        common_date = target_date.strip()
    elif detected_dates:
        counts = {}
        for d in detected_dates:
            counts[d] = counts.get(d, 0) + 1
        common_date = sorted(counts.keys(), key=lambda x: (counts[x], x), reverse=True)[0]

    results = []
    for fn in filenames:
        b_info, new_name, dt = normalize_filename(fn, force_report_type=report_type, force_date=common_date)
        results.append({
            "old_name": fn,
            "new_name": new_name,
            "branch_num": b_info["num"] if b_info else "??",
            "branch_name": b_info["std_name_weekly"] if b_info else "문서 내부 확인 필요",
            "recognized": b_info is not None,
            "changed": fn != new_name,
        })
    return jsonify({"common_date": common_date, "results": results})


@app.route('/api/rename_local_folder', methods=['POST'])
def api_rename_local_folder():
    """서버(내 PC)의 특정 폴더 경로에 있는 파일들의 이름을 직접 일괄 변경합니다."""
    data = request.get_json(force=True, silent=True) or {}
    folder_path = data.get('folder_path', '').strip()
    target_date = data.get('target_date', '').strip() or None
    dry_run = bool(data.get('dry_run', False))
    
    if not folder_path or not os.path.isdir(folder_path):
        return jsonify({"ok": False, "error": f"유효한 폴더 경로가 아닙니다: {folder_path}"}), 400
        
    try:
        results = batch_rename_in_folder(folder_path, unify_date=True, target_date=target_date, dry_run=dry_run)
        return jsonify({"ok": True, "folder_path": folder_path, "results": results})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route('/api/browse_folder', methods=['POST'])
def api_browse_folder():
    """윈도우 폴더 선택 대화상자를 띄워 선택된 경로를 반환합니다."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder_selected = filedialog.askdirectory(title="파일명을 일괄 변경할 폴더를 선택하세요")
        root.destroy()
        return jsonify({"ok": True, "folder_path": folder_selected})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route('/rename_zip', methods=['POST'])
def rename_zip():
    """업로드된 파일들의 이름을 표준 규격으로 일괄 변경하여 ZIP 파일로 다운로드합니다."""
    clean_uploads()
    files = request.files.getlist('branch_files')
    target_date = request.form.get('target_date', '').strip() or None
    report_type = request.form.get('report_type_filter', '').strip() or None

    saved_paths = []
    for file in files:
        if file.filename and not file.filename.startswith('~'):
            dest = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(dest)
            saved_paths.append(dest)

    if not saved_paths:
        flash("변경할 파일이 선택되지 않았습니다.", "warning")
        return redirect(url_for('index'))

    batch_rename_in_folder(UPLOAD_FOLDER, unify_date=True, target_date=target_date, dry_run=False)

    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname in sorted(os.listdir(UPLOAD_FOLDER)):
            if fname.startswith('~'):
                continue
            fpath = os.path.join(UPLOAD_FOLDER, fname)
            if os.path.isfile(fpath):
                zf.write(fpath, arcname=fname)
    mem_zip.seek(0)
    return send_file(mem_zip, mimetype='application/zip', as_attachment=True, download_name="표준파일명_일괄변경완료.zip")


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
