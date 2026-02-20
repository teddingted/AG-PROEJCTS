# Python UTF-8 Encoding Fix for Windows
**문제 해결: cp949 환경에서 한글 및 Unicode 문자 출력 오류**

## 문제 증상
```
UnicodeEncodeError: 'cp949' codec can't encode character '\u2713' in position 0: illegal multibyte sequence
```

Windows에서 Python 기본 인코딩이 cp949(한국어 EUC-KR)로 설정되어 있어, UTF-8 문자(✓, →, 한글 등)를 출력할 때 오류가 발생합니다.

---

## 해결 방법

### 방법 1: 환경 변수 영구 설정 (권장)
**`fix_encoding.bat` 실행**
- 위치: `c:\Users\Admin\Desktop\AG-BEGINNING\fix_encoding.bat`
- 실행 방법: 더블 클릭 후 "확인"
- 효과: 모든 Python 프로그램에서 UTF-8 자동 사용
- **터미널 재시작** 필요

```batch
setx PYTHONIOENCODING utf-8
setx PYTHONUTF8 1
```

---

### 방법 2: 각 Python 파일에 직접 추가
**모든 Python 스크립트 상단에 다음 코드 추가:**

```python
# -*- coding: utf-8 -*-
import sys
import os

# UTF-8 Fix for Windows cp949
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'
```

---

### 방법 3: 자동 패치 스크립트 사용
**`autopatch_encoding.py` 실행**

```bash
python autopatch_encoding.py
```

- 기존 모든 Python 파일에 UTF-8 fix 자동 삽입
- 이미 패치된 파일은 건너뜀

---

## 적용된 파일 목록
✓ `visualize_bog_events.py`  
✓ `analyze_control_logic.py`  
✓ `deep_analysis.py`  
✓ `validate_fds_sequence.py`  
✓ `extract_all_controllers.py`  
✓ `extract_fds.py`  

---

## 새 Python 파일 작성 시
**`_encoding_fix_template.py`를 복사해서 시작**

```bash
copy _encoding_fix_template.py my_new_script.py
```

이미 UTF-8 fix가 포함되어 있어 즉시 사용 가능합니다.

---

## 검증
```python
python _encoding_fix_template.py
```

출력 예상:
```
✓ UTF-8 encoding test: 한글 테스트
한국어와 Unicode 문자 (✓✗→←) 모두 정상 출력
```

---

## 문제 재발 시 체크리스트
1. [ ] 터미널을 재시작했는가?
2. [ ] `fix_encoding.bat`를 실행했는가?
3. [ ] Python 파일 상단에 UTF-8 fix 코드가 있는가?
4. [ ] `PYTHONIOENCODING` 환경변수가 설정되었는가? (`echo %PYTHONIOENCODING%`)

---

**솔루션 작성자**: Antigravity AI  
**적용 날짜**: 2026-02-13  
**상태**: ✓ 검증 완료
