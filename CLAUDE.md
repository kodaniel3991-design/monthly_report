# 월차보고 시스템 — Claude Code 컨텍스트

## 프로젝트 한 줄 요약
진양오토모티브 김해 공장의 **월별 경영실적을 자동 집계·계산·보고서 생성**하는 내부 웹 시스템.  
기존에 5개 Excel 파일을 수동으로 작성하던 월차보고 업무를 자동화한다.

---

## 기술 스택

| 역할 | 선택 | 이유 |
|------|------|------|
| Web UI | **Streamlit** | 1인 사용, Python 단일 파일 |
| DB | **SQLite** (`data/monthly_report.db`) | 설치 불필요, 단순 |
| Excel 생성 | **openpyxl** | 기존 양식 재현 |
| 계산 | **Python dataclass** (`core/calculator.py`) | 순수 함수, 테스트 용이 |
| 차트 | **Plotly** | Streamlit 네이티브 |

```bash
# 실행
streamlit run app.py

# 패키지 설치
pip install -r requirements.txt
```

---

## 폴더 구조

```
monthly_report_system/
├── CLAUDE.md                  ← 이 파일 (Claude Code 자동 인식)
├── app.py                     ← Streamlit 진입점 + 사이드바
├── requirements.txt
├── data/
│   └── monthly_report.db      ← SQLite (자동 생성됨)
├── core/
│   ├── database.py            ← DB 스키마 + CRUD 함수
│   ├── calculator.py          ← 모든 계산 로직 (수식 역공학 완료)
│   └── excel_generator.py     ← Excel 보고서 생성 (openpyxl)
└── pages/                     ← Streamlit 멀티페이지
    ├── 01_손익실적_입력.py
    ├── 02_인원_노무비_입력.py
    ├── 03_업계동향_입력.py
    ├── 04_대시보드.py
    └── 05_보고서_다운로드.py
```

---

## 핵심 비즈니스 규칙 (절대 변경 금지)

### 사업부 구분
- **RKM** = 김해공장 + 부산공장 (르노코리아 납품)
- **HKMC** = 울산공장 + 김해2공장 (현대기아 납품)

### 부가가치 공식 ★★★
```
附加價値 = 賣出額 - 変動費 + 変動複利厚生費
         = 한계이익 + 복리후생비(변동)
```
> **근거**: 기존 Excel 파일 역공학 검증 완료  
> RKM: 3,503,150 - 2,608,461 + 9,905 = 904,594 ✓  
> HKMC: 1,345,027 - 1,094,678 + 1,305 = 251,654 ✓  
> 전체: 4,848,177 - 3,703,139 + 11,210 = 1,156,248 ✓

### 6대 노동생산성 지표
```python
부가가치율     = 附加價値 / 賣出額
노동생산성     = 附加價値 / 상시종업원수           # 단위: 천원/인
근로소득배분율  = 勞務費 / 附加價値
인건비율       = 人件費 / 賣出額
1인당임금수준   = (勞務費 + 退職金) / 종업원수      # 단위: 천원/인/월
시간당임금     = 勞務費 / 실작업시간              # 단위: 천원/h (생산직)
```

### 노무비 RKM/HKMC 배분 방식
```
기본급 배분 = 총노무비 × 근무시간비율
상여금 배분 = 사업부별 실제 지급액 (별도 입력)
최종 RKM 노무비 = 기본급_RKM + 상여금_RKM
```

### 데이터 입력 순서 (월 마감 후)
```
① 손익실적 입력 (공장별 4개)
② 인원·노무비 입력 (RAW 데이터)
③ 업계동향 입력 (선택)
④ 대시보드 → 자동 계산 확인
⑤ 보고서 다운로드 → Excel 생성
```

---

## DB 스키마 요약

```sql
monthly_pl      -- 손익실적 (공장별 4개 × 40개 항목, UNIQUE: year+month)
monthly_labor   -- 인원·근무시간·상여·퇴직급여 (UNIQUE: year+month)
annual_plan     -- 연간 사업계획 (UNIQUE: year+month+item_code)
industry_news   -- 업계동향 뉴스
monthly_top_models    -- 국내 판매 TOP10
monthly_market_share  -- 시장점유율
```

---

## 기존 Excel 파일과의 관계 (원본 참조용)

| 원본 파일 | 시스템 대응 | 상태 |
|-----------|------------|------|
| `3월_손익_실적.xlsx` | `monthly_pl` 테이블 + `01_손익실적_입력.py` | ✅ 완료 |
| `2_1_노동생산성.xlsx` - 노무비(인원,근무시간) | `monthly_labor` 테이블 | ✅ 완료 |
| `2_1_노동생산성.xlsx` - 노무비(25-05) | `monthly_labor` 테이블 | ✅ 완료 |
| `2_2_사업계획.xlsx` - 사업계획(1118)최종 | `annual_plan` 테이블 | ⚠️ 입력 UI 미완 |
| `2_2_사업계획.xlsx` - 회계팀자료 3월 | `monthly_pl` + 자동 계산 | ✅ 자동화됨 |
| `1_1_업계동향.xlsx` | `industry_news` 테이블 | ✅ 완료 |
| `1_2_운영실적.xls` | 미구현 (서술형, 낮은 우선순위) | ❌ 미착수 |

---

## 현재 완료 / 남은 작업

### ✅ 완료
- DB 스키마 6개 테이블
- 계산 엔진 (`core/calculator.py`) — 부가가치 공식 검증 완료
- Streamlit 5개 화면 뼈대
- Excel 생성 기본 구조 (`core/excel_generator.py`)
- 테스트 데이터 DB 저장 완료 (2026년 3월 기준)

### 🔴 우선순위 높음 (다음에 할 것)
1. **사업계획 입력 UI** — `pages/06_사업계획_입력.py` 신규 작성  
   → `annual_plan` 테이블에 월별 계획치 저장  
   → 손익실적 화면에서 계획 대비 실적 컬럼 채우기

2. **손익실적 Excel 출력 정확도 개선**  
   → 기존 `3월_손익_실적.xlsx`와 숫자·레이아웃 1:1 비교  
   → 퍼센트(%) 컬럼 정확히 채우기

3. **누계 계산**  
   → `load_all_months(year)` 함수 이미 있음  
   → 대시보드에 1~N월 누계 섹션 추가

### 🟡 우선순위 중간
4. 전월 대비 자동 비교 (대시보드)
5. 노동생산성 Excel 출력 정확도 개선
6. 월별 추이 차트 (꺾은선)

### 🟢 낮은 우선순위
7. 운영실적(서술형) 입력 및 Excel 출력
8. 사용자 인증 (현재 1인 사용이므로 불필요)
9. 배포 (로컬 실행으로 충분)

---

## 자주 쓰는 패턴

### DB에서 데이터 불러와서 계산하기
```python
from core.database import load_monthly_pl, load_monthly_labor
from core.calculator import (
    build_factory_pl_from_db, build_labor_input_from_db,
    _sum_factories, calc_value_added,
    calc_labor_productivity_total, calc_labor_productivity_by_division
)

pl_data    = load_monthly_pl(year, month)
labor_data = load_monthly_labor(year, month)

gimhae  = build_factory_pl_from_db(pl_data, "gimhae")
busan   = build_factory_pl_from_db(pl_data, "busan")
ulsan   = build_factory_pl_from_db(pl_data, "ulsan")
gimhae2 = build_factory_pl_from_db(pl_data, "gimhae2")

rkm   = _sum_factories("RKM",  gimhae, busan)
hkmc  = _sum_factories("HKMC", ulsan,  gimhae2)
total = _sum_factories("계",   gimhae, busan, ulsan, gimhae2)
labor = build_labor_input_from_db(labor_data)
```

### 새 Streamlit 페이지 기본 구조
```python
import streamlit as st
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

st.set_page_config(page_title="페이지명", layout="wide")
st.title("제목")

year  = st.selectbox("연도", range(2024, 2028), index=2)
month = st.selectbox("월",   range(1, 13),       index=2)
```

### DB 저장 패턴
```python
# INSERT OR UPDATE (upsert)
save_monthly_pl(year, month, data_dict)
save_monthly_labor(year, month, data_dict)
```

---

## 코딩 규칙

- 단위는 항상 **천원(KRW)** — 주석이나 라벨에 명시
- 공장 코드: `gimhae` / `busan` / `ulsan` / `gimhae2`
- 사업부 코드: `rkm` / `hkmc`
- DB 컬럼명: `{항목코드}_{공장코드}` (예: `material_gimhae`)
- Streamlit `number_input` key는 짧게 (`f"mat_{fcode}"`)
- 0 또는 None 처리: `data.get(key, 0) or 0`

---

## 검증 기준값 (2026년 3월 실적)

새 계산 로직을 추가할 때 아래 값으로 검증할 것.

| 항목 | RKM | HKMC | 전체 |
|------|-----|------|------|
| 賣出額 | 3,503,150 | 1,345,027 | 4,848,177 |
| 附加價値 | 904,594 | 251,654 | 1,156,248 |
| 附加價値率 | 25.82% | 18.71% | 23.85% |
| 勞動生産性 | 16,229 | 5,440 | 11,336 |
| 종업원수 | 55.7명 | 46.3명 | 102명 |
| 실작업시간 | 8,772h | 6,643h | 15,415.5h |
