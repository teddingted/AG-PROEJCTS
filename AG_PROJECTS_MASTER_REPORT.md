# AG-PROJECTS & AG-PROJECTS-mac- 종합 및 비전 보고서
**작성일**: 2026-02-23  
**문서 목적**: `github/teddingted/AG-projects` (Windows 환경) 및 `AG-projects-mac-` (Mac 환경) 저장소의 성과를 집대성하고, 향후 프로젝트 통합 및 비전을 수립함.

---

## 1. 프로젝트 통합 및 환경 표준화 성과 (Executive Summary)

현재 AG-PROJECTS는 산발적으로 진행되던 여러 아이디어를 규격화된 디렉토리 구조(`Project 1~5`)로 편입시키며 **본격적인 모듈화 궤도**에 올랐습니다. 특히 핵심 성과는 다음과 같습니다.

1. **저장소 이원화 및 규칙 확립**: Windows 개발 환경(`AG-PROJECTS`)과 Mac 개발 환경(`AG-PROJECTS-mac-`) 간의 개발 격차를 줄이기 위해 코어 모듈을 `AG-REPOSITORY` 체제로 통합하고, `git_rules.md`를 기반으로 한 일관된 Git 워크플로우를 정립했습니다.
2. **다양한 형태의 릴리즈 성공**: 순수 Python 스크립트뿐만 아니라 HTA 기반 애플리케이션, Windows 포팅(설치 무관 구동) 등 산업 현장(폐쇄망)에서도 즉각 활용할 수 있는 뛰어난 이식성을 달성했습니다.

---

## 2. 프로젝트별 발자취 및 주요 성과 (Key Achievements)

### [Project 1] Auto Plot Digitizer
* **개요**: 그래프/플롯 이미지에서 수치 데이터를 역엔지니어링하여 디지털화하는 도구.
* **주요 성과**:
  * 툴팁 UI 및 X/Y 좌표, 변화율 표시 기능 고도화로 데이터 추출 직관성 극대화.
  * Windows 환경 완벽 포팅 완료(`AutoPlotDigitizerV2_Windows_Port`).

### [Project 2] Hi-Way Calculator
* **개요**: 업무 효율을 극대화하기 위한 맞춤형 빠른 계산기 앱.
* **주요 성과**:
  * 파이썬 스크립트 종속성을 탈피하여 **HTA(HTML Application)** 기반 단독 실행 체계 확립.
  * 패키징 스크립트(`organize_hiway.bat`)를 통한 일괄 구조화 완료.

### [Project 3] HYSYS Automation (가장 발전된 플래그십 모델)
* **개요**: Python-HYSYS COM Interface 연동을 통한 공정 시뮬레이션 무인 최적화기.
* **주요 성과**:
  * 수작업 기준 1케이스당 3분 소요되던 작업을 **초단위(무인) 자동화** 완료 (일 2,000+ 케이스 처리).
  * **하이브리드 탐색 알고리즘**: 온도/압력 극저온 공정에서의 Solver 발산을 잡기 위해 'Grid Scan'과 'Smart Anchor(Warm Reset)' 로직을 융합 결합.
  * **Advanced Data Mining**: 24시간 분량의 운전 데이터 분석을 통해 밸브 피로도(DPCV111 순환 문제), 열교환기 Reverse Flow 감지, 12시간 연속 Alarm 오작동 등 **5가지 숨겨진 현장 엔지니어링 인사이트** 도출 성공.

---

## 3. 향후 프로젝트 디벨롭 방향 및 지향점 (Future Directions & Vision)

지금까지의 과정이 **"분산된 아이디어의 구현과 통합(PoC)"**이었다면, 향후 방향성은 **"엔터프라이즈급 안정성 확보와 AI 융합"**을 지향해야 합니다.

### 🎯 [방향 1] 기술적 고도화 및 CI/CD 파이프라인 구축
* **Cross-Platform 동기화 완성**: `AG-projects`와 `AG-projects-mac-` 브랜치 간의 완전한 CI/CD(GitHub Actions 등) 구축. 커밋 푸시 시 Windows 빌드와 Mac 구동 테스트가 자동으로 이루어지는 파이프라인 고려.
* **플랫폼 독립형 웹 기반 전환**: HTA나 특정 OS 독립 실행 파일도 좋지만, 장기적으로는 Node.js나 Python 웹서버(FastAPI)+React/Vue 기반의 로컬 웹앱 구조로 전환하여 Mac과 Windows에서 동일한 브라우저 UI로 접근하도록 고도화.

### 🎯 [방향 2] AI / Data Driven 아키텍처로의 진화 (HYSYS Automation)
* **머신러닝 기반 공정 이상 탐지 (Anomaly Detection)**: Advanced Insights Report에서 확인한 수동 데이터 패턴 마이닝을 자동화 모듈로 격상. HYSYS 구동 중 실시간으로 수집되는 값 위에서 AI가 스스로 병목과 밸브 피로도를 에측.
* **Autonomous Engineering**: 사용자가 Flow만 입력하면, 스크립트가 스스로 가장 비용(Power)이 적고 안정성(MA)이 높은 지점을 탐색한 뒤 Report를 PDF로 렌더링하여 이메일로 발송하는 완전 무인화 시스템 구축.

### 🎯 [방향 3] "AG Engineering Master Suite" 생태계 조성
* 개별 프로젝트(Digitizer, Calculator, Hysys Optimizer)를 묶는 **그랜드 대시보드(Master Launcher)** 개발.
* 프로젝트 간의 데이터 호환성 확보. 예를 들어, Auto Plot Digitizer로 추출한 현장 커브 데이터를 Hysys Automation의 기초 컴프레서 커브 곡선으로 직접 Feed-in 할 수 있는 유기적 연결고리(Data Pipeline) 개발.

---

## 4. 맺음말 (Conclusion)

안전과 신뢰성이 최우선인 산업 공학 도구에, 최신 스크립팅과 데이터 마이닝을 접목한 지금까지의 시도는 매우 창의적이고 실용적인 가치를 지닙니다. 향후 Mac과 PC의 완벽한 듀얼 체제 위에서 프로젝트들이 상호 침투하고 확장되는 **'연결된 엔지니어링 생태계'**를 만드는 것이 다음 마일스톤이 될 것입니다.
