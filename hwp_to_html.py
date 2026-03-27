# -*- coding: utf-8 -*-
"""
hwp_to_html.py
==============
HWP 파일을 HTML 템플릿으로 변환하는 스크립트

동작 순서:
  1. HWP 파일(OLE 컴파운드 문서)을 열고 BodyText 스트림 추출
  2. zlib 압축 해제 후 HWP 바이너리 레코드 파싱
  3. 텍스트 및 테이블 구조를 분석하여 HTML로 변환
  4. 다우오피스 전자결재 양식에 맞는 HTML 템플릿 생성

필요 패키지: pip install olefile
"""

import olefile
import zlib
import struct
import sys
from pathlib import Path


# ─────────────────────────────────────────────
# 1. HWP 본문 스트림 추출 (압축 해제 포함)
# ─────────────────────────────────────────────
def extract_body_sections(hwp_path: str) -> list:
    """
    HWP 파일에서 BodyText 섹션들을 추출합니다.

    Args:
        hwp_path: HWP 파일 경로

    Returns:
        각 섹션의 압축 해제된 바이너리 데이터 리스트
    """
    ole = olefile.OleFileIO(hwp_path)

    # FileHeader에서 압축 여부 확인
    header = ole.openstream('FileHeader').read()
    flags = int.from_bytes(header[36:40], 'little')
    is_compressed = bool(flags & 0x01)
    is_encrypted = bool(flags & 0x02)

    if is_encrypted:
        ole.close()
        raise ValueError("암호화된 HWP 파일은 지원하지 않습니다.")

    # BodyText 섹션들 수집
    sections = []
    for stream in ole.listdir():
        path = '/'.join(stream)
        if path.startswith('BodyText/Section'):
            raw = ole.openstream(path).read()
            if is_compressed:
                # HWP는 raw deflate 사용 (-15 = windowBits)
                decompressed = zlib.decompress(raw, -15)
                sections.append(decompressed)
            else:
                sections.append(raw)

    ole.close()
    return sections


# ─────────────────────────────────────────────
# 2. HWP 레코드에서 텍스트 추출
# ─────────────────────────────────────────────
def parse_hwp_text(section_data: bytes) -> list:
    """
    HWP 바이너리 섹션 데이터에서 텍스트를 추출합니다.

    HWP 레코드 구조:
      - 4바이트 헤더: 태그ID(10bit) + 레벨(10bit) + 크기(12bit)
      - 크기가 0xFFF이면 다음 4바이트가 실제 크기
      - HWPTAG_PARA_TEXT(태그 67): UTF-16LE 인코딩된 문단 텍스트

    Args:
        section_data: 압축 해제된 섹션 바이너리 데이터

    Returns:
        추출된 텍스트 문단 리스트
    """
    pos = 0
    paragraphs = []

    while pos < len(section_data):
        if pos + 4 > len(section_data):
            break

        # 레코드 헤더 읽기
        header = struct.unpack_from('<I', section_data, pos)[0]
        tag_id = header & 0x3FF          # 하위 10비트: 태그 ID
        level = (header >> 10) & 0x3FF   # 중간 10비트: 레벨
        size = (header >> 20) & 0xFFF    # 상위 12비트: 데이터 크기
        pos += 4

        # 크기가 0xFFF이면 확장 크기 사용
        if size == 0xFFF:
            if pos + 4 > len(section_data):
                break
            size = struct.unpack_from('<I', section_data, pos)[0]
            pos += 4

        # HWPTAG_PARA_TEXT = 67 (HWPTAG_BEGIN(66) + 1)
        if tag_id == 67:
            text = _decode_para_text(section_data[pos:pos + size])
            paragraphs.append(text)

        pos += size

    return paragraphs


def _decode_para_text(text_data: bytes) -> str:
    """
    HWP 문단 텍스트 데이터를 UTF-16LE에서 파이썬 문자열로 디코딩합니다.

    HWP 특수 제어 문자(코드 0~31)를 적절히 처리합니다:
      - 코드 0: 무시
      - 코드 1~3, 11, 15~18, 21~24: 확장 제어 (뒤에 14바이트 추가 데이터)
      - 코드 9: 탭 (확장 제어 + 탭 문자)
      - 코드 10: 줄바꿈
      - 코드 13: 문단 끝
    """
    text = ''
    i = 0

    # 확장 제어 문자 목록 (뒤에 14바이트 추가 데이터가 따라옴)
    # HWP 문서에서 코드 1~9, 11, 15~18, 21~24는 확장 제어 문자
    EXTENDED_CHARS = {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 15, 16, 17, 18, 21, 22, 23, 24}

    while i < len(text_data) - 1:
        char_code = struct.unpack_from('<H', text_data, i)[0]
        i += 2

        if char_code < 32:
            # 확장 제어 문자는 추가 14바이트 건너뛰기
            if char_code in EXTENDED_CHARS:
                if char_code == 9:
                    text += '\t'
                elif char_code == 11:
                    text += '[표]'  # 표/그리기 개체 자리
                i += 14
            elif char_code == 10:
                text += '\n'
            elif char_code == 13:
                pass  # 문단 끝 (별도 처리 불필요)
            # 그 외 제어 문자는 무시
        else:
            text += chr(char_code)

    return text


# ─────────────────────────────────────────────
# 3. 추출된 텍스트를 HTML로 변환
# ─────────────────────────────────────────────
def texts_to_html(paragraphs: list, title: str = "") -> str:
    """
    추출된 문단 텍스트들을 다우오피스 전자결재용 HTML로 변환합니다.

    라벨-값 패턴을 자동 감지하여 테이블 구조의 양식 HTML을 생성합니다.

    Args:
        paragraphs: 문단 텍스트 리스트
        title: 문서 제목 (없으면 첫 번째 비어있지 않은 문단 사용)

    Returns:
        HTML 문자열
    """
    # 비어있는 문단과 [표] 마커, 워터마크 텍스트 제거
    SKIP_KEYWORDS = {'[표]', '문서서식포탈비', '문서서식포탈비즈폼', '폼', '비즈폼'}
    cleaned = []
    for p in paragraphs:
        stripped = p.strip()
        if stripped and stripped not in SKIP_KEYWORDS:
            cleaned.append(stripped)

    # 제목 자동 감지 (한자 공백이 들어간 "사  고  경  위  서" 같은 패턴)
    doc_title = title
    for i, text in enumerate(cleaned):
        # 공백이 2개 이상 포함되고, 서/서류/서약 등으로 끝나는 문서 제목 패턴
        if '  ' in text and len(text) < 30:
            doc_title = text
            cleaned.pop(i)
            break

    # 라벨-값 쌍 감지
    # HWP 양식에서 라벨과 값은 연속된 문단으로 나타남
    # 라벨 패턴: 한글 + 공백 (예: "성    명", "사고일시", "부    서")
    LABEL_PATTERNS = [
        '결', '재', '담', '이', '본 부 장',  # 결재란 (건너뛰기)
    ]
    FORM_LABELS = [
        '사 고 자', '부서', '성명', '직위', '연 락 처', '연락처',
        '사고일시', '사고장소', '사고차량', '사고원인', '사고내용',
        '사진첨부', '작 성 자', '작성자',
    ]

    # 결재란 텍스트 제거
    skip_labels = {'결', '재', '담   당', '이   사', '본 부 장', '담당', '이사'}

    # 라벨-값 쌍을 구조화
    form_fields = []
    i = 0
    while i < len(cleaned):
        text = cleaned[i]

        # 결재란 관련 텍스트 건너뛰기
        if text in skip_labels:
            i += 1
            continue

        # 라벨인지 확인 (공백이 많은 짧은 텍스트)
        is_label = False
        normalized = text.replace(' ', '')
        for label in FORM_LABELS:
            label_norm = label.replace(' ', '')
            if normalized == label_norm or normalized.startswith(label_norm):
                is_label = True
                break

        # "사고원인 및" + "사고내용" 패턴 처리 (라벨이 여러 줄에 걸침)
        if normalized in ('사고원인및', '사고원인'):
            # 다음 문단이 "사고내용"이면 합치기
            if i + 1 < len(cleaned) and cleaned[i + 1].replace(' ', '') == '사고내용':
                label = '사고원인 및 사고내용'
                # 그 다음이 실제 값
                value = cleaned[i + 2] if i + 2 < len(cleaned) else ''
                form_fields.append((label, value))
                i += 3
                continue

        if is_label:
            # 다음 문단이 값
            value = cleaned[i + 1] if i + 1 < len(cleaned) else ''
            # 값이 또 다른 라벨인지 확인
            val_norm = value.replace(' ', '')
            is_next_label = any(val_norm == l.replace(' ', '') or val_norm.startswith(l.replace(' ', ''))
                               for l in FORM_LABELS)
            if is_next_label:
                form_fields.append((text, ''))  # 값 없는 라벨
                i += 1
            else:
                form_fields.append((text, value))
                i += 2
        else:
            # 라벨이 아닌 단독 텍스트 (제출 문구 등)
            form_fields.append((None, text))
            i += 1

    # HTML 생성
    html_parts = []
    html_parts.append('<!-- HWP에서 자동 변환된 HTML 양식 -->')
    html_parts.append('<div data-id="appContent">')
    html_parts.append('')
    html_parts.append('<div style="font-family: \'Malgun Gothic\', dotum, Arial, sans-serif; '
                      'font-size: 10pt; line-height: 1.6; margin: 0 auto; max-width: 800px;">')
    html_parts.append('')

    # 문서 제목
    html_parts.append('  <!-- 문서 제목 -->')
    html_parts.append('  <div data-id="appTitle" style="text-align: center; margin-bottom: 20px;">')
    html_parts.append(f'    <h2 style="font-size: 16pt; font-weight: bold; letter-spacing: 4px; '
                      f'border-bottom: 2px solid #333; padding-bottom: 8px;">')
    html_parts.append(f'      {_html_escape(doc_title)}')
    html_parts.append('    </h2>')
    html_parts.append('  </div>')
    html_parts.append('')

    # 양식 테이블 생성
    html_parts.append('  <!-- 양식 내용 -->')
    html_parts.append('  <table border="1" cellpadding="6" cellspacing="0"')
    html_parts.append('         style="width: 100%; border-collapse: collapse; margin-bottom: 16px;">')

    param_counter = 0
    for label, value in form_fields:
        param_counter += 1
        param_id = f'apprPostParam{param_counter}'
        escaped_value = _html_escape(value)

        if label is None:
            # 단독 텍스트 (제출 문구 등) — 전체 너비 셀
            html_parts.append(f'    <tr>')
            html_parts.append(f'      <td colspan="2" style="text-align: center; padding: 10px;">')
            html_parts.append(f'        {escaped_value}')
            html_parts.append(f'      </td>')
            html_parts.append(f'    </tr>')
        elif '\n' in value:
            # 여러 줄 값 → textarea
            html_parts.append(f'    <tr>')
            html_parts.append(f'      <td style="background-color: #f5f6f8; font-weight: bold; '
                            f'text-align: center; width: 25%; vertical-align: top;">')
            html_parts.append(f'        {_html_escape(label)}')
            html_parts.append(f'      </td>')
            html_parts.append(f'      <td>')
            html_parts.append(f'        <textarea name="field_{param_counter}" data-id="{param_id}" '
                            f'style="width: 98%; min-height: 80px; border: none; '
                            f'font-family: inherit; font-size: 10pt; resize: vertical;"'
                            f'>{escaped_value}</textarea>')
            html_parts.append(f'      </td>')
            html_parts.append(f'    </tr>')
        else:
            # 일반 라벨-값 쌍 → input 필드
            html_parts.append(f'    <tr>')
            html_parts.append(f'      <td style="background-color: #f5f6f8; font-weight: bold; '
                            f'text-align: center; width: 25%;">')
            html_parts.append(f'        {_html_escape(label)}')
            html_parts.append(f'      </td>')
            html_parts.append(f'      <td>')
            html_parts.append(f'        <input type="text" name="field_{param_counter}" '
                            f'data-id="{param_id}" value="{escaped_value}" '
                            f'style="width: 95%; border: none; font-family: inherit; font-size: 10pt;" />')
            html_parts.append(f'      </td>')
            html_parts.append(f'    </tr>')

    html_parts.append('  </table>')
    html_parts.append('')
    html_parts.append('</div><!-- /font-family div -->')
    html_parts.append('</div><!-- /data-id="appContent" -->')

    return '\n'.join(html_parts)


def _html_escape(text: str) -> str:
    """HTML 특수문자 이스케이프 처리"""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


# ─────────────────────────────────────────────
# 4. 메인 실행부
# ─────────────────────────────────────────────
def convert_hwp_to_html(hwp_path: str, output_path: str = None) -> str:
    """
    HWP 파일을 HTML로 변환하는 메인 함수입니다.

    Args:
        hwp_path: 입력 HWP 파일 경로
        output_path: 출력 HTML 파일 경로 (없으면 자동 생성)

    Returns:
        생성된 HTML 파일 경로
    """
    hwp_path = Path(hwp_path)

    if not hwp_path.exists():
        raise FileNotFoundError(f"HWP 파일을 찾을 수 없습니다: {hwp_path}")

    # 출력 경로 자동 설정
    if not output_path:
        output_path = hwp_path.with_suffix('.html')
    output_path = Path(output_path)

    print(f"[1/3] HWP 파일 읽는 중: {hwp_path}")
    sections = extract_body_sections(str(hwp_path))
    print(f"       → {len(sections)}개 섹션 추출 완료")

    print(f"[2/3] 텍스트 파싱 중...")
    all_paragraphs = []
    for section in sections:
        paragraphs = parse_hwp_text(section)
        all_paragraphs.extend(paragraphs)
    print(f"       → {len(all_paragraphs)}개 문단 추출 완료")

    print(f"[3/3] HTML 변환 중...")
    html = texts_to_html(all_paragraphs, title=hwp_path.stem)

    # 파일 저장
    output_path.write_text(html, encoding='utf-8')
    print(f"       → 저장 완료: {output_path}")
    print(f"       → 크기: {len(html.encode('utf-8')):,} bytes")

    return str(output_path)


if __name__ == "__main__":
    # 명령줄 인자 또는 기본 파일 경로
    if len(sys.argv) > 1:
        hwp_file = sys.argv[1]
    else:
        hwp_file = "(양식)차량사고경위서.hwp"

    try:
        result_path = convert_hwp_to_html(hwp_file)
        print(f"\n변환 완료! 결과 파일: {result_path}")
    except Exception as e:
        print(f"\n오류 발생: {e}", file=sys.stderr)
        sys.exit(1)
