# =============================================================================
# config.py - 다우오피스 조직도 동기화 설정 파일
# =============================================================================
# 이 파일에 인증키와 각종 설정값을 입력합니다.
# 보안상 이 파일은 외부에 공유하지 마세요!
# =============================================================================

# ─────────────────────────────────────────────
# [필수] 다우오피스 OpenAPI 인증 정보
# 위치: 통합설정 > 시스템연동 > 연동관리 > OpenAPI
# ─────────────────────────────────────────────
DAOU_CLIENT_ID     = "여기에_클라이언트ID_입력"       # 예: f1c2f0b4beeaaffe
DAOU_CLIENT_SECRET = "여기에_클라이언트_비밀번호_입력"  # 예: xxxxxxxxxxxxxx

# ─────────────────────────────────────────────
# [필수 아님] 다우오피스 API 서버 주소 (변경 불필요)
# ─────────────────────────────────────────────
DAOU_API_BASE = "https://api.daouoffice.com"

# ─────────────────────────────────────────────
# [선택] 퇴사자 판정 기준
# expiredDate(퇴사일)가 오늘 이전이거나 status가 STOP이면 퇴사자로 처리
# ─────────────────────────────────────────────
RESIGN_STATUS = ["STOP"]          # 이 상태값이면 퇴사/중지로 분류
DORMANT_STATUS = ["DORMANT"]      # 메일 휴면 상태 (별도 처리 가능)

# ─────────────────────────────────────────────
# [선택] 보고서 및 데이터 저장 경로
# ─────────────────────────────────────────────
import os
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))  # 스크립트 폴더
DATA_DIR    = os.path.join(BASE_DIR, "data")              # 이전 데이터 저장
LOG_DIR     = os.path.join(BASE_DIR, "logs")              # 실행 로그
REPORT_DIR  = os.path.join(BASE_DIR, "reports")           # 엑셀 보고서

# 이전 스냅샷 파일명 (비교 기준 데이터)
SNAPSHOT_FILE = os.path.join(DATA_DIR, "last_snapshot.json")

# ─────────────────────────────────────────────
# [선택] AD(Active Directory) 연동 설정
# 실제 환경에 맞게 수정하세요
# ─────────────────────────────────────────────
AD_ENABLED      = False                    # True로 변경하면 AD 연동 활성화
AD_SERVER       = "ldap://192.168.1.10"   # AD 서버 IP
AD_DOMAIN       = "company.local"          # AD 도메인
AD_ADMIN_USER   = "administrator"          # AD 관리자 계정
AD_ADMIN_PASS   = "패스워드입력"            # AD 관리자 비밀번호
AD_BASE_DN      = "DC=company,DC=local"   # AD 기본 경로
AD_USER_OU      = "OU=Users,DC=company,DC=local"  # 사용자 OU 경로

# ─────────────────────────────────────────────
# [선택] 이메일 알림 설정 (보고서 자동 발송용)
# ─────────────────────────────────────────────
EMAIL_ENABLED   = False
EMAIL_HOST      = "smtp.company.com"
EMAIL_PORT      = 587
EMAIL_USER      = "it@company.com"
EMAIL_PASS      = "이메일패스워드"
EMAIL_RECEIVERS = ["manager@company.com", "hr@company.com"]  # 수신자 목록
