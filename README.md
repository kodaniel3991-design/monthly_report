# 월차보고 시스템 — 설치 및 실행 가이드

## 1단계: 필수 소프트웨어 설치

### Python 3.11+
```
https://www.python.org/downloads/
```
설치 시 **"Add Python to PATH"** 반드시 체크

### Visual Studio Code
```
https://code.visualstudio.com/
```

### Claude Code (VS Code 확장)
VS Code 실행 → 확장(Extensions, Ctrl+Shift+X) → **"Claude"** 검색 → 설치  
또는 터미널에서:
```bash
code --install-extension Anthropic.claude-code
```

### Git
```
https://git-scm.com/downloads
```

---

## 2단계: 프로젝트 세팅

```bash
# 1. 프로젝트 폴더로 이동
cd monthly_report_system

# 2. 가상환경 생성
python -m venv .venv

# 3. 가상환경 활성화
# Windows
.venv\Scripts\activate
# Mac/Linux
source .venv/bin/activate

# 4. 패키지 설치
pip install -r requirements.txt

# 5. VS Code 워크스페이스 열기
code monthly_report.code-workspace
```

---

## 3단계: 실행

```bash
# 가상환경 활성화 상태에서
streamlit run app.py
```

브라우저에서 `http://localhost:8501` 자동 열림

---

## 4단계: Claude Code와 바이브 코딩 시작

### Claude Code가 자동으로 읽는 파일
```
CLAUDE.md  ← 세션 시작 시 자동 인식
```
이 파일에 프로젝트 전체 컨텍스트가 담겨 있어서  
새 세션을 열어도 설명 없이 바로 코딩 가능.

### Claude Code 실행
- VS Code 좌측 사이드바 Claude 아이콘 클릭
- 또는 단축키: `Ctrl+Shift+C` (Mac: `Cmd+Shift+C`)

### 첫 메시지 예시
```
CLAUDE.md 읽어줘. 
오늘은 사업계획 입력 UI (pages/06_사업계획_입력.py)를 만들 거야.
annual_plan 테이블에 월별 계획치를 저장하고,
01_손익실적_입력.py에서 계획 대비 실적을 비교할 수 있게 해줘.
```

---

## 월별 작업 흐름

```
매월 마감 후 (약 30분 소요 목표)

1. 브라우저에서 http://localhost:8501 접속
2. ① 손익실적 입력 → 공장별 실적 입력 → 저장
3. ② 인원·노무비 입력 → 인원, 상여, 퇴직 → 저장
4. ③ 업계동향 입력 → 뉴스 텍스트 → 저장 (선택)
5. ④ 대시보드 → 자동 계산 결과 확인
6. ⑤ 보고서 다운로드 → Excel 저장
```

---

## 파일 구조

```
monthly_report_system/
├── CLAUDE.md          ★ Claude Code 컨텍스트 (핵심)
├── app.py             Streamlit 메인
├── requirements.txt
├── monthly_report.code-workspace
├── data/
│   └── monthly_report.db   (자동 생성)
├── core/
│   ├── calculator.py  계산 엔진
│   ├── database.py    DB CRUD
│   └── excel_generator.py  Excel 생성
└── pages/
    ├── 01_손익실적_입력.py
    ├── 02_인원_노무비_입력.py
    ├── 03_업계동향_입력.py
    ├── 04_대시보드.py
    └── 05_보고서_다운로드.py
```

---

## 자주 있는 문제

**Q: `streamlit: command not found`**
```bash
# 가상환경이 활성화됐는지 확인
.venv\Scripts\activate   # Windows
source .venv/bin/activate  # Mac
```

**Q: DB 초기화 오류**
```bash
# data 폴더가 없을 경우
mkdir data
python core/database.py  # DB 재생성
```

**Q: 포트 충돌**
```bash
streamlit run app.py --server.port 8502
```
