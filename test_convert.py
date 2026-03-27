# -*- coding: utf-8 -*-
"""
테스트 스크립트: 가상 휴가신청서를 unified_converter로 변환하고 결과를 저장합니다.
"""
from pathlib import Path
from unified_converter import convert

# 입력/출력 경로
input_file = str(Path(__file__).parent / "(테스트)휴가신청서.docx")
output_file = str(Path(__file__).parent / "test_output" / "휴가신청서_변환결과.html")

# 출력 폴더 생성
Path(output_file).parent.mkdir(exist_ok=True)

# 변환 실행
print(f"입력: {input_file}")
print(f"출력: {output_file}")
convert(input_file, output_file)
print("변환 완료!")

# 변환 결과 읽어서 출력
result = Path(output_file).read_text(encoding='utf-8')
print(f"\n변환 결과 크기: {len(result)} bytes")
print("=" * 60)
print(result)
print("=" * 60)
