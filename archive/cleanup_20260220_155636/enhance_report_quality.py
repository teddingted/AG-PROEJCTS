# -*- coding: utf-8 -*-
"""
Professional Report Quality Enhancement Script
전문 보고서 품질 향상 스크립트
"""
import sys
import os

# UTF-8 Fix
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

import re

def enhance_report_quality(input_path, output_path):
    """
    Enhance report to professional commissioning standard
    전문 시운전 보고서 표준으로 향상
    """
    
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 용어 표준화 매핑
    terminology_map = {
        # 기술 용어 → 공식 용어
        r'\bTAG\b': '계측기 번호',
        r'\bTag\b': '계측기 번호',
        r'\btag\b': '계측기',
        r'\bsignal\b': '신호',
        r'\bSignal\b': '신호',
        r'\bSensor_Tag\b': '계측기 식별번호',
        r'\bComponent\b': '구성품',
        
        # 축약어 확장
        r'\bPV\b(?!C)': '공정 변수(PV)',  # Process Variable
        r'\bSP\b(?!\s)': '설정값(SP)',  # Set Point
        r'\bOP\b': '출력값(OP)',  # Output
        r'\bCV\b': '제어 밸브',
        r'\bHX\b': '열교환기',
        
        # 센서/장비 명칭 개선
        r'ERS_(\w+)_Y': r'\1 계측기',
        r'ERS_(\w+)_OUTPOS': r'\1 밸브 위치',
        r'ERS_(\w+)_CTRNOUT': r'\1 제어기 출력',
        
        # 전문 용어 사용
        r'data point': '측정값',
        r'dataset': '데이터셋',
        r'outlier': '이상치',
        r'anomaly': '이상 현상',
        
        # 수치 표현 개선
        r'(\d+)개': r'\1개소',
        r'(\d+)회': r'\1회',
    }
    
    enhanced = content
    
    # 용어 치환
    for pattern, replacement in terminology_map.items():
        enhanced = re.sub(pattern, replacement, enhanced)
    
    # 섹션 제목 표준화
    section_improvements = {
        '## 시험 목적': '## 2. 시험 목적 및 범위',
        '## 시험 구성': '## 3. 시험 구성 및 방법',
        '## 시험 결과': '## 4. 시험 결과 및 분석',
        '## 센서 성능 검증': '## 5. 계측기 성능 검증',
        '## 성능 검증': '## 6. 시스템 성능 적합성 검증',
        '## 발견사항 및 권고사항': '## 7. 발견사항 및 개선 권고',
        '## 결론': '## 8. 종합 평가 및 결론',
    }
    
    for old, new in section_improvements.items():
        enhanced = enhanced.replace(old, new)
    
    # 전문적 표현 개선
    professional_phrases = {
        '확인됨': '확인되었음',
        '검증됨': '검증되었음',
        '성공': '정상적으로 완료',
        '실패': '미달',
        '문제': '개선 필요 사항',
        '버그': '오류',
        '튜닝': '조정',
        '체크': '확인',
        
        # 판정 표현
        '합격': '적합',
        '불합격': '부적합',
        'PASS': '적합',
        'FAIL': '부적합',
        
        # 수치 표현
        '약 ': '약 ',  # 전각 공백 제거
        ' %': '%',  # 공백 제거
    }
    
    for old, new in professional_phrases.items():
        enhanced = enhanced.replace(old, new)
    
    # 문장 끝 표준화 (평서문)
    enhanced = re.sub(r'입니다\.', '임.', enhanced)
    enhanced = re.sub(r'습니다\.', '음.', enhanced)
    
    # 계측기 번호 형식 개선 (ERS_81PIT0001_Y → 압력계 81-PT-0001)
    def format_instrument_tag(match):
        tag = match.group(1)
        
        # Parse instrument type
        if 'PIT' in tag or 'PI' in tag:
            return f'압력계 {tag}'
        elif 'TIT' in tag or 'TI' in tag or 'TE' in tag:
            return f'온도계 {tag}'
        elif 'FIT' in tag or 'FI' in tag:
            return f'유량계 {tag}'
        elif 'LIT' in tag or 'LI' in tag:
            return f'레벨계 {tag}'
        elif 'PCV' in tag or 'TCV' in tag or 'FCV' in tag:
            return f'제어밸브 {tag}'
        elif 'PIC' in tag or 'TIC' in tag or 'FIC' in tag:
            return f'제어기 {tag}'
        else:
            return f'계측기 {tag}'
    
    # enhanced = re.sub(r'ERS_(\w+)_\w+', format_instrument_tag, enhanced)
    
    # Save enhanced version
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(enhanced)
    
    print(f"✓ 보고서 품질 향상 완료")
    print(f"  입력: {input_path}")
    print(f"  출력: {output_path}")
    print(f"  길이: {len(content):,} → {len(enhanced):,} 자")
    
    return enhanced

if __name__ == "__main__":
    input_file = r"c:\Users\Admin\Desktop\AG-BEGINNING\시운전_시험성적서_8196_FULL.md"
    output_file = r"c:\Users\Admin\Desktop\AG-BEGINNING\시운전_시험성적서_8196_PROFESSIONAL.md"
    
    if os.path.exists(input_file):
        enhance_report_quality(input_file, output_file)
        print("\n다음 단계: PDF 재생성")
        print("  python convert_to_pdf_enhanced.py")
    else:
        print(f"Error: {input_file} not found")
