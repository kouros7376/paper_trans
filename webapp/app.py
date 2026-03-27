# -*- coding: utf-8 -*-
"""
app.py
======
다우오피스 전자결재 양식 변환기 - 웹 애플리케이션

관리부 직원이 브라우저에서 파일을 업로드하면
자동으로 다우오피스 전자결재용 HTML로 변환해주는 사내 웹 서비스입니다.

사용법:
  python app.py              → http://localhost:5000 에서 접속
  python app.py --port 8080  → 포트 변경
  python app.py --host 0.0.0.0  → 외부 접속 허용 (사내 배포용)

지원 파일: HWP, DOCX, XLSX, PDF
"""

import os
import sys
import uuid
import logging
from pathlib import Path
from datetime import datetime

from flask import (
    Flask, render_template, request, redirect,
    url_for, send_file, flash, jsonify
)
from werkzeug.utils import secure_filename

# ── 상위 폴더의 unified_converter 모듈 임포트 ──
sys.path.insert(0, str(Path(__file__).parent.parent))
from unified_converter import convert

# ─────────────────────────────────────────────
# Flask 앱 설정
# ─────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.urandom(24)  # 세션/플래시 메시지용 비밀키

# 파일 업로드 설정
BASE_DIR = Path(__file__).parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
CONVERTED_FOLDER = BASE_DIR / "converted"
UPLOAD_FOLDER.mkdir(exist_ok=True)
CONVERTED_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 최대 50MB
app.config['TEMPLATES_AUTO_RELOAD'] = True  # 템플릿 변경 즉시 반영

# 지원하는 파일 확장자
ALLOWED_EXTENSIONS = {'hwp', 'docx', 'xlsx', 'xls', 'pdf'}

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def allowed_file(filename: str) -> bool:
    """파일 확장자가 지원 목록에 있는지 확인합니다."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ─────────────────────────────────────────────
# 라우트 (페이지) 정의
# ─────────────────────────────────────────────

@app.route('/')
def index():
    """메인 페이지 - 파일 업로드 화면"""
    # 변환 이력 조회 (최근 20건)
    history = _get_conversion_history()
    return render_template('index.html', history=history)


@app.route('/convert', methods=['POST'])
def convert_file():
    """파일 업로드 및 변환 처리"""

    # ── 파일 유효성 검사 ──
    if 'file' not in request.files:
        flash('파일이 선택되지 않았습니다.', 'error')
        return redirect(url_for('index'))

    file = request.files['file']
    original_name = file.filename or ''
    logger.info("업로드 요청 수신 - filename: '%s', content_type: '%s'",
                original_name, file.content_type)

    if original_name == '':
        flash('파일이 선택되지 않았습니다.', 'error')
        return redirect(url_for('index'))

    if not allowed_file(original_name):
        flash(f'지원하지 않는 파일 형식입니다. (지원: {", ".join(ALLOWED_EXTENSIONS)})', 'error')
        return redirect(url_for('index'))

    # ── 파일 저장 ──
    # 고유 ID를 붙여서 파일명 충돌 방지
    unique_id = uuid.uuid4().hex[:8]
    # 확장자 추출 (한글 파일명에서도 안전하게)
    ext = Path(original_name).suffix.lower()
    safe_name = f"{unique_id}{ext}"
    upload_path = UPLOAD_FOLDER / safe_name
    file.save(str(upload_path))

    logger.info("파일 업로드 완료: %s (%s)", original_name, _format_size(upload_path.stat().st_size))

    # ── 변환 실행 ──
    output_name = f"{unique_id}_{Path(original_name).stem}.html"
    output_path = CONVERTED_FOLDER / output_name

    try:
        convert(str(upload_path), str(output_path))

        # 제목이 UUID(파일명)인 경우 원본 파일명으로 교체
        html_content = output_path.read_text(encoding='utf-8')
        original_stem = Path(original_name).stem
        if unique_id in html_content:
            html_content = html_content.replace(unique_id, original_stem)
            output_path.write_text(html_content, encoding='utf-8')

        logger.info("변환 성공: %s → %s", original_name, output_name)

        # 변환 이력 저장
        _save_history(unique_id, original_name, output_name)

        flash(f'변환 완료! [{original_name}] → HTML 양식이 생성되었습니다.', 'success')
        return redirect(url_for('preview', file_id=unique_id))

    except Exception as e:
        logger.error("변환 실패: %s - %s", original_name, str(e))
        flash(f'변환 중 오류가 발생했습니다: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/clear-history', methods=['POST'])
def clear_history():
    """변환 이력 및 파일 전체 삭제"""
    import shutil
    # 이력 파일 삭제
    if HISTORY_FILE.exists():
        HISTORY_FILE.unlink()
    # 업로드/변환 폴더 비우기
    for folder in [UPLOAD_FOLDER, CONVERTED_FOLDER]:
        for f in folder.iterdir():
            if f.is_file():
                f.unlink()
    logger.info("변환 이력 및 파일 전체 삭제 완료")
    return '', 204


@app.route('/preview/<file_id>')
def preview(file_id):
    """변환 결과 미리보기 페이지"""
    history = _get_conversion_history()
    entry = None
    for h in history:
        if h['id'] == file_id:
            entry = h
            break

    if not entry:
        flash('해당 변환 결과를 찾을 수 없습니다.', 'error')
        return redirect(url_for('index'))

    # 변환된 HTML 내용 읽기
    output_path = CONVERTED_FOLDER / entry['output_name']
    if not output_path.exists():
        flash('변환 파일이 삭제되었습니다.', 'error')
        return redirect(url_for('index'))

    html_content = output_path.read_text(encoding='utf-8')
    file_size = _format_size(output_path.stat().st_size)

    return render_template(
        'preview.html',
        entry=entry,
        html_content=html_content,
        file_size=file_size,
        history=history,
    )


@app.route('/download/<file_id>')
def download(file_id):
    """변환된 HTML 파일 다운로드"""
    history = _get_conversion_history()
    for h in history:
        if h['id'] == file_id:
            output_path = CONVERTED_FOLDER / h['output_name']
            if output_path.exists():
                # 다운로드 파일명을 원본 이름 기반으로 설정
                download_name = Path(h['original_name']).stem + '_다우오피스양식.html'
                return send_file(
                    str(output_path),
                    as_attachment=True,
                    download_name=download_name,
                )
    flash('파일을 찾을 수 없습니다.', 'error')
    return redirect(url_for('index'))


@app.route('/raw/<file_id>')
def raw_html(file_id):
    """변환된 HTML 원본 내용 반환 (iframe용)"""
    history = _get_conversion_history()
    for h in history:
        if h['id'] == file_id:
            output_path = CONVERTED_FOLDER / h['output_name']
            if output_path.exists():
                return output_path.read_text(encoding='utf-8')
    return '<p>파일을 찾을 수 없습니다.</p>', 404


# ─────────────────────────────────────────────
# 변환 이력 관리 (간단한 텍스트 파일 기반)
# ─────────────────────────────────────────────
HISTORY_FILE = BASE_DIR / "conversion_history.txt"


def _save_history(file_id: str, original_name: str, output_name: str):
    """변환 이력을 파일에 저장합니다."""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ext = Path(original_name).suffix.lower()
    line = f"{file_id}|{original_name}|{output_name}|{timestamp}|{ext}\n"
    with open(HISTORY_FILE, 'a', encoding='utf-8') as f:
        f.write(line)


def _get_conversion_history(limit: int = 20) -> list:
    """최근 변환 이력을 가져옵니다."""
    if not HISTORY_FILE.exists():
        return []
    lines = HISTORY_FILE.read_text(encoding='utf-8').strip().split('\n')
    history = []
    for line in reversed(lines[-limit:]):
        parts = line.strip().split('|')
        if len(parts) >= 5:
            history.append({
                'id': parts[0],
                'original_name': parts[1],
                'output_name': parts[2],
                'timestamp': parts[3],
                'file_type': parts[4].upper(),
            })
    return history


def _format_size(size_bytes: int) -> str:
    """바이트 수를 읽기 쉬운 형태로 변환합니다."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


# ─────────────────────────────────────────────
# 서버 실행
# ─────────────────────────────────────────────
if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='다우오피스 전자결재 양식 변환기 웹 서버')
    parser.add_argument('--host', default='0.0.0.0', help='서버 주소 (기본: 0.0.0.0 = 외부 접속 허용)')
    parser.add_argument('--port', type=int, default=5000, help='포트 번호 (기본: 5000)')
    parser.add_argument('--debug', action='store_true', help='디버그 모드')
    args = parser.parse_args()

    print("=" * 55)
    print("  다우오피스 전자결재 양식 변환기")
    print(f"  http://localhost:{args.port} 에서 접속하세요")
    if args.host == '0.0.0.0':
        import socket
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        print(f"  사내 접속: http://{local_ip}:{args.port}")
    print("=" * 55)

    app.run(host=args.host, port=args.port, debug=args.debug)
