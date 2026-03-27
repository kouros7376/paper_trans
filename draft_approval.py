"""
draft_approval.py
=================
다우오피스 전자결재 기안 자동화 스크립트 (방안 A: HTML 폼 연동형)

동작 순서:
  1. .env 파일에서 API 인증 정보(clientId, clientSecret) 로드
  2. 테스트용 더미 데이터 딕셔너리를 HTML 템플릿(template.html)에 병합
  3. multipart/form-data 형식으로 기안 API에 POST 요청
  4. 응답 결과(성공 URL 또는 오류 내용)를 콘솔에 출력

참고 엔드포인트:
  POST https://api.daouoffice.com/public/v4/approval/document/popup
"""

import os
import sys
import logging
from datetime import date
from pathlib import Path

import requests
from dotenv import load_dotenv


# ─────────────────────────────────────────────
# 0. 로깅 설정
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# 1. 환경 변수 로드 (.env 파일)
# ─────────────────────────────────────────────
# 스크립트와 같은 폴더의 .env 파일을 자동으로 읽어옵니다.
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

CLIENT_ID     = os.getenv("DAOU_CLIENT_ID")
CLIENT_SECRET = os.getenv("DAOU_CLIENT_SECRET")

# ─── 인증 정보 누락 시 즉시 종료 ───
if not CLIENT_ID or not CLIENT_SECRET:
    logger.error(
        ".env 파일에 DAOU_CLIENT_ID 또는 DAOU_CLIENT_SECRET 값이 없습니다.\n"
        "  → .env.example 파일을 복사하여 .env 파일을 만든 뒤 실제 값을 입력해 주세요."
    )
    sys.exit(1)


# ─────────────────────────────────────────────
# 2. API 기본 설정
# ─────────────────────────────────────────────
BASE_URL    = "https://api.daouoffice.com"
API_PATH    = "/public/v4/approval/document/popup"
FULL_URL    = BASE_URL + API_PATH

# 요청 타임아웃 (초) — 네트워크 환경에 맞게 조정하세요.
REQUEST_TIMEOUT = 30


# ─────────────────────────────────────────────
# 3. 테스트용 더미 데이터 정의
#    ↑ 실제 사용 시 이 딕셔너리의 값만 교체하면 됩니다.
# ─────────────────────────────────────────────
today = date.today()

# --- 인적사항 ---
DRAFT_DATA = {
    # 인계자 정보
    "hand_over_dept":     "정보시스템팀",
    "hand_over_position": "팀장",
    "hand_over_name":     "홍길동",

    # 인수자 정보
    "hand_in_dept":       "정보시스템팀",
    "hand_in_position":   "과장",
    "hand_in_name":       "김철수",

    # 부서장
    "dept_manager":       "이부장",

    # 인수인계 프로젝트
    "project": (
        "1. 사내 그룹웨어(다우오피스) 유지보수 및 신규 연동 개발\n"
        "2. ERP 시스템 연간 유지보수 계약 관리\n"
        "3. 재해복구(DR) 시스템 운영"
    ),

    # 중요 문제점 및 개선사항 (최대 4개)
    "issue_1": "ERP ↔ 다우오피스 전자결재 연동 테스트 미완료 (담당자 인수 필요)",
    "issue_2": "NAS 스토리지 용량 80% 초과 — 아카이빙 정책 수립 필요",
    "issue_3": "방화벽 정책 연간 검토 일정: 2026년 6월 예정",
    "issue_4": "",

    # 진행 및 미결사항
    "progress_status": (
        "ERP 전자결재 연동 개발: 70% 완료, 테스트 단계 진행 중\n"
        "NAS 증설 발주: 견적 수령 완료, 구매 품의 기안 예정"
    ),
    "pending_status": (
        "ERP 연동 테스트 완료 후 운영 환경 배포 (2026-04-15 목표)\n"
        "NAS 증설 최종 승인 대기"
    ),

    # 서류 및 기타
    "documents": (
        "서버 관리 대장 (최신본: \\\\NAS\\IT_Docs\\서버관리대장_2026.xlsx)\n"
        "네트워크 구성도 (최신본: \\\\NAS\\IT_Docs\\네트워크구성도_v3.vsd)"
    ),
    "others": (
        "공유 폴더 경로: \\\\NAS\\IT_Docs (읽기/쓰기 권한 신청 필요)\n"
        "GitHub 저장소: github.com/company/it-automation (접근 권한 이관 완료)"
    ),

    # PMS 업무일지 최종입력 일자
    "pms_last_date": str(today.year),
    "pms_month":     str(today.month),
    "pms_day":       str(today.day),

    # 제출일
    "submit_year":  str(today.year),
    "submit_month": str(today.month),
    "submit_day":   str(today.day),
}

# --- 전자결재 기안 설정 ---
APPROVAL_CONFIG = {
    # 기안자 사원번호 (다우오피스에 등록된 사번과 정확히 일치해야 합니다)
    "draftEmpNo": "220125",

    # 다우오피스 Site Admin > 결재 양식에서 부여한 고유 코드
    "formCode": "HandOverForm",

    # 결재 문서 제목 (동적으로 생성)
    "title": f"업무인수인계서_{DRAFT_DATA['hand_over_name']}→{DRAFT_DATA['hand_in_name']}_{today.strftime('%Y%m%d')}",

    # (선택) 결재 처리 결과를 수신할 콜백 URL — 사내 서버 주소로 교체하세요.
    "callbackUrl": "",

    # (선택) 외부 시스템에서 관리하는 문서 고유 ID
    "partnerDocId": f"HANDOVER-{today.strftime('%Y%m%d')}-001",

    # (선택) 제품명/버전 — 연동 시스템 식별용
    "productName": "IT자동화시스템",
    "productVersion": "1.0",
}


# ─────────────────────────────────────────────
# 4. HTML 템플릿 로드 및 데이터 병합
# ─────────────────────────────────────────────
def load_template(template_path: Path, data: dict) -> str:
    """
    HTML 템플릿 파일을 읽어서 {변수명} 자리에 실제 데이터를 채워 반환합니다.

    Args:
        template_path: template.html 파일 경로
        data: 치환할 데이터 딕셔너리

    Returns:
        데이터가 채워진 HTML 문자열
    """
    if not template_path.exists():
        raise FileNotFoundError(
            f"HTML 템플릿 파일을 찾을 수 없습니다: {template_path}\n"
            "  → 이 스크립트와 같은 폴더에 template.html이 있어야 합니다."
        )

    # UTF-8로 템플릿 파일 읽기
    raw_html = template_path.read_text(encoding="utf-8")

    # 줄바꿈 문자(\n)가 HTML 텍스트에 있으면 <br>로 변환
    # (textarea 안의 내용은 그대로, input value 안의 줄바꿈만 처리)
    sanitized_data = {}
    for key, value in data.items():
        # HTML 특수문자 이스케이프 (XSS 방지 기본 처리)
        safe_value = (
            str(value)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )
        sanitized_data[key] = safe_value

    # str.format_map(): 딕셔너리 키를 {키} 자리에 치환
    # 누락된 키가 있어도 빈 문자열로 처리하는 DefaultDict 활용
    class _DefaultDict(dict):
        def __missing__(self, key):
            logger.warning("템플릿 변수 '{%s}'에 매핑된 데이터가 없어 빈 값으로 처리합니다.", key)
            return ""

    filled_html = raw_html.format_map(_DefaultDict(sanitized_data))
    return filled_html


# ─────────────────────────────────────────────
# 5. 전자결재 기안 API 호출
# ─────────────────────────────────────────────
def submit_draft(content_html: str, config: dict) -> None:
    """
    다우오피스 전자결재 기안 API를 호출합니다.

    Args:
        content_html: 결재 본문 HTML 문자열
        config: 기안 설정 딕셔너리 (APPROVAL_CONFIG)
    """
    # multipart/form-data 파라미터 구성
    # ※ requests 라이브러리에서 files= 매개변수를 쓰면 자동으로 multipart 처리됩니다.
    form_data = {
        # 필수 인증 파라미터
        "clientId":     (None, CLIENT_ID),
        "clientSecret": (None, CLIENT_SECRET),

        # 필수 기안 파라미터
        "formCode":    (None, config["formCode"]),
        "title":       (None, config["title"]),
        "draftEmpNo":  (None, config["draftEmpNo"]),
        "content":     (None, content_html, "text/html; charset=utf-8"),
    }

    # 선택 파라미터 — 값이 있을 때만 추가
    optional_fields = ["callbackUrl", "partnerDocId", "productName", "productVersion"]
    for field in optional_fields:
        value = config.get(field, "")
        if value:
            form_data[field] = (None, value)

    logger.info("=" * 55)
    logger.info("다우오피스 전자결재 기안 API 호출 시작")
    logger.info("  URL        : %s", FULL_URL)
    logger.info("  문서 제목  : %s", config["title"])
    logger.info("  기안자 사번: %s", config["draftEmpNo"])
    logger.info("  양식 코드  : %s", config["formCode"])
    logger.info("=" * 55)

    try:
        # POST 요청 전송 (requests가 Content-Type: multipart/form-data 자동 설정)
        response = requests.post(
            url=FULL_URL,
            files=form_data,
            timeout=REQUEST_TIMEOUT,
        )

        # ─── 성공 응답 처리 ───
        if response.status_code == 200:
            logger.info("[성공] HTTP 200 응답 수신")

            # 다우오피스는 기안 팝업 HTML을 직접 반환하거나 JSON으로 응답할 수 있음
            content_type = response.headers.get("Content-Type", "")
            if "application/json" in content_type:
                try:
                    result = response.json()
                    logger.info("  응답 JSON : %s", result)

                    # 일반적으로 code 필드로 성공/실패 판단
                    code = str(result.get("code", ""))
                    if code == "200" or code.upper() == "OK":
                        logger.info("  → 기안 성공! 결재 문서가 생성되었습니다.")
                        doc_url = result.get("data", {}).get("docUrl") or result.get("docUrl")
                        if doc_url:
                            logger.info("  결재 문서 URL: %s", doc_url)
                    else:
                        logger.warning("  → API 처리 중 오류: code=%s, message=%s",
                                       code, result.get("message", "메시지 없음"))
                except ValueError:
                    logger.warning("  JSON 파싱 실패. 원본 응답을 확인하세요.")
                    logger.info("  응답 내용 (앞 500자): %s", response.text[:500])

            elif "text/html" in content_type:
                # 기안 팝업 HTML이 직접 반환된 경우 — 브라우저에서 열어야 함
                logger.info("  → 기안 팝업 HTML 반환됨. 아래 내용을 브라우저에서 렌더링하세요.")
                # 반환된 HTML을 파일로 저장 (확인용)
                popup_path = Path(__file__).parent / "draft_popup_result.html"
                popup_path.write_text(response.text, encoding="utf-8")
                logger.info("  결과 HTML 저장 완료: %s", popup_path)
            else:
                logger.info("  응답 Content-Type: %s", content_type)
                logger.info("  응답 내용 (앞 500자): %s", response.text[:500])

        # ─── 클라이언트 오류 (4xx) ───
        elif 400 <= response.status_code < 500:
            logger.error("[실패] HTTP %d 오류", response.status_code)
            _log_error_details(response)

        # ─── 서버 오류 (5xx) ───
        elif response.status_code >= 500:
            logger.error("[실패] HTTP %d 서버 내부 오류 — 다우기술 지원팀에 문의하세요.", response.status_code)
            _log_error_details(response)

        else:
            logger.warning("[예상치 못한 응답] HTTP %d", response.status_code)
            _log_error_details(response)

    except requests.exceptions.SSLError as e:
        logger.error("[SSL 오류] HTTPS 인증서 문제가 발생했습니다: %s", e)
        logger.error("  → 사내 방화벽/프록시가 SSL을 차단하는지 확인하세요.")

    except requests.exceptions.ConnectionError as e:
        logger.error("[연결 오류] API 서버에 접속할 수 없습니다: %s", e)
        logger.error("  → 네트워크 연결 상태 및 방화벽 설정을 확인하세요.")
        logger.error("  → 허용 도메인: api.daouoffice.com, doas.daouoffice.com")

    except requests.exceptions.Timeout:
        logger.error("[타임아웃] %d초 이내에 응답이 없었습니다.", REQUEST_TIMEOUT)
        logger.error("  → 네트워크 상태를 확인하거나 REQUEST_TIMEOUT 값을 늘려보세요.")

    except requests.exceptions.RequestException as e:
        logger.error("[요청 오류] 예상치 못한 오류가 발생했습니다: %s", e)


def _log_error_details(response: requests.Response) -> None:
    """오류 응답의 상태 코드와 내용을 상세히 로그로 출력합니다."""
    logger.error("  상태 코드 : %d", response.status_code)
    logger.error("  응답 내용 : %s", response.text[:1000])

    # JSON 오류 메시지가 있으면 추가 파싱
    try:
        error_json = response.json()
        logger.error("  오류 코드 : %s", error_json.get("code", "N/A"))
        logger.error("  오류 메시지: %s", error_json.get("message", "N/A"))
    except ValueError:
        pass  # JSON이 아닌 경우 위에서 이미 텍스트 출력함


# ─────────────────────────────────────────────
# 6. 메인 실행부
# ─────────────────────────────────────────────
def main():
    """메인 실행 함수: 템플릿 로드 → 데이터 병합 → API 호출"""

    # HTML 템플릿 파일 경로 (이 스크립트와 같은 폴더)
    template_path = Path(__file__).parent / "template.html"

    try:
        logger.info("HTML 템플릿 로드 중: %s", template_path)
        filled_html = load_template(template_path, DRAFT_DATA)
        logger.info("  템플릿 로드 완료 (크기: %d bytes)", len(filled_html.encode("utf-8")))

    except FileNotFoundError as e:
        logger.error("%s", e)
        sys.exit(1)

    # 기안 API 호출
    submit_draft(filled_html, APPROVAL_CONFIG)


if __name__ == "__main__":
    main()
