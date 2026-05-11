# AJW 발주계획 시스템 — COWork 인수인계 문서

> 작성일: 2026-05-11 | 담당자: 박정원 (SCM팀) | Git: keiruun-sgv  
> 마지막 커밋: `5921406` — STEP 1~3 파싱 오류 수정 및 판매분석(STEP 2) 기능 추가

---

## 1. 프로젝트 개요

(주)에이제이월드 SCM팀이 사용하는 **생산자재 발주계획 자동화 웹앱**.  
맥산텔레콤 ERP에서 추출한 Excel 파일을 업로드하면 연간 발주계획 Excel을 자동 생성한다.

- **실행**: `streamlit run web_app.py --server.address=0.0.0.0 --server.port=8501`
- **venv 경로**: `AI/연간 발주 계획 새우기/.venv/Scripts/streamlit.exe`
- **접속 URL**: http://localhost:8501
- **GitHub**: https://github.com/keiruun-svg/my-first (master 브랜치)

---

## 2. 파일 구조

```
AI/연간 발주 계획 새우기/
├── web_app.py          ★ Streamlit 웹앱 진입점 (6탭 UI + 비즈니스 로직)
├── step1_core.py       ★ STEP 1 코어 — ERP 파일 → 생산자재 사용내역 Excel
├── step2_core.py       ★ STEP 3 코어 — 가공파일 → 연간발주계획 Excel
├── convert_core.py     ★ ERP 원본 파싱 — 구매조회/구매현황 → 연도별 시트
├── ojc_filter.py         판매량 파일에서 OJC 완제품만 추출 (STEP 2에서 import)
├── metadata.json         품번·품명·구매처·리드타임 영구 저장
├── settings.json         리드타임 기본값·색상 설정 영구 저장
├── inventory.json        케이블·하우징 현재고·기발주 영구 저장
└── sales_analysis.json   STEP 2 판매 분석 결과 영구 저장 (STEP 3에서 자동 참조)
```

---

## 3. 3-STEP 워크플로우

```
[ERP 원본 파일 — 구매현황 or 구매조회]
      │
      ▼
┌─────────────┐
│   STEP 1   │──► 다운로드 ①: 생산자재_사용내역.xlsx  (참고용, STEP 3 입력 아님)
│  ERP 가공  │──► 다운로드 ②: 가공파일(연도별시트).xlsx  ← STEP 3 입력
└─────────────┘

[전체 판매량.xlsx] + [구매관리(맥산).xlsx, 선택]
      │
      ▼
┌─────────────┐
│   STEP 2   │──► sales_analysis.json 자동 저장 (STEP 3 자동 참조)
│  판매 분석  │──► OJC_판매량_정리.xlsx 다운로드
└─────────────┘

[가공파일(연도별시트).xlsx]  +  sales_analysis.json (자동 참조)
      │
      ▼
┌─────────────┐
│   STEP 3   │──► 연간발주계획.xlsx
│  발주계획   │
└─────────────┘
```

### 중요: STEP 3 입력 파일
- ✅ **올바른 입력**: `가공파일(연도별시트).xlsx` (STEP 1 다운로드 ②)
- ✅ **직접 입력 가능**: ERP 원본 파일 (자동 변환)
- ❌ **잘못된 입력**: `생산자재_사용내역.xlsx` (STEP 1 다운로드 ①) → 명확한 오류 발생

---

## 4. 탭 구성 (web_app.py)

| 탭 | 변수 | 기능 |
|----|------|------|
| 📤 STEP 1 — ERP 파일 가공 | tab1 | 구매조회/구매현황 → 사용내역 + 가공파일 생성 |
| 📈 STEP 2 — 판매 분석 | tab2 | 전체 판매량 OJC 필터링 + 생산비중 계산 |
| 📊 STEP 3 — 발주계획 생성 | tab3 | 가공파일 → 연간발주계획 Excel |
| 📋 품번 관리 | tab4 | 케이블·하우징·페롤 품번 CRUD |
| 📦 재고 현황 | tab5 | 현재고·기발주 입력 |
| ⚙️ 파라미터 & 양식 설정 | tab6 | 리드타임 기본값·색상 설정 |

---

## 5. 입력 파일 형식 (convert_core.py 기준)

### 5-1. 맥산 납품 형식 (구매현황 시트, 자동 감지)

헤더 Row2의 첫 컬럼에 '구매처', '거래처', '공급처', '납품처' 중 하나 포함 시 자동 분기.

| cells 인덱스 | 컬럼명 | 예시 |
|-------------|--------|------|
| [0] | 구매처명 (무시) | 맥산텔레콤 |
| [1] | 납고번호 (무시) | 24/09/09-1 |
| [2] | **납고일자** | 2024/09/05 -3 |
| [3] | **품목코드** | 14-C-565 |
| [4] | **품목명** | DROP-CABLE(LC/PC-FC/PC) |
| [5] | 창고명 (무시) | 021_의왕-물류 |
| [6] | **규격명** | 15M (없으면 None) |
| [7] | **수량** | 10 |

- Row 1: 회사명/기간 (무시), Row 2: 헤더, Row 3~: 데이터 (`min_row=3`)

### 5-2. 기존 구매현황 형식

| cells 인덱스 | 컬럼명 |
|-------------|--------|
| [0] | **날짜** (YYYY/MM/DD -N) |
| [4] | **품목코드** |
| [5] | **품목명[규격명]** |
| [6] | **수량** |

### 5-3. 날짜 파싱 (_parse_buy_date)

| 형식 | 예시 | 처리 |
|------|------|------|
| datetime 객체 | openpyxl data_only=True 반환 | `.year`, `.month` 사용 |
| YYYY/MM/DD | 2024/09/05 -3 | regex |
| YYYY-MM-DD | 2025-01-09 00:00:00 | regex |
| YYYYMMDD | 20250109-2 | regex |
| YY/MM/DD | 24/09/09-1 | '20'+YY 변환 |

---

## 6. 영구 저장 JSON 구조

### metadata.json
```json
{
  "cable": {
    "2.0mm|A1-SP": {"품번": "P14-RM-4188", "품명": "...", "구매처": "HUAMAI", "리드타임": 60}
  },
  "housing": {
    "2.0mm|LC/PC 청색": [
      {"품번": "P14-LS-4228", "품명": "...", "구매처": "FIBERCAN", "리드타임": 60}
    ]
  },
  "ferrule": {
    "LC/PC": {"품번": "...", "품명": "FERRULE (W/ FLANGE,LC/PC TYPE)", "구매처": "", "리드타임": null}
  }
}
```

### sales_analysis.json (STEP 2 출력 → STEP 3 자동 참조)
```json
{
  "14-C-565": {
    "품목명": "DROP-CABLE(LC/PC-FC/PC)",
    "23": {"sales": 1200, "production": 960, "ratio": 0.8},
    "24": {"sales": 1500, "production": 1200, "ratio": 0.8},
    "25": {"sales": 870,  "production": 700,  "ratio": 0.805}
  }
}
```

### inventory.json
```json
{
  "cable":   {"2.0mm|A1-SP": {"현재고": 500, "기발주": 200}},
  "housing": {"2.0mm|LC/PC 청색": [{"현재고": 1000, "기발주": 0}]}
}
```

---

## 7. 핵심 함수 시그니처

### step1_core.run()
```python
result_bytes, logs, cable_keys, housing_keys, years, converted_bytes = step1_core.run(
    row_path,        # ERP 원본 또는 가공파일 경로
    cable_meta,      # {(파이, 케이블종류): {품번, 품명, 구매처, 리드타임}}
    housing_meta_in, # {(파이, 하우징타입): [{품번, 품명, ...}]}
    ojc_ref_path,    # 보통 None
    settings,        # {"lead_time_default": 60, "colors": {...}}
    progress_cb      # callable(pct: int, msg: str) or None
)
# converted_bytes: 연도별 시트 가공파일 bytes → STEP 3 입력용
```

### step2_core.run() (STEP 3 발주계획)
```python
xlsx_bytes, logs = step2_core.run(
    row_path,          # 가공파일(연도별시트).xlsx 경로
    usage_path=None,   # 구버전 호환, 보통 None
    ojc_ref_path=None,
    cmp_path=None,     # 구버전 호환, 보통 None
    settings={},
    ferrule_meta={},   # {커넥터타입: {품번, 품명, ...}}
    progress_cb=None,
    cable_meta_in={},  # metadata.json["cable"]
    housing_meta_in={},# metadata.json["housing"]
    inventory={},      # inventory.json
    sales_data=None    # sales_analysis.json (STEP 2 결과)
)
```

### web_app.build_sales_analysis() (STEP 2 판매분석)
```python
analysis, summary, errors, ojc_xl_bytes = build_sales_analysis(
    sales_bytes,     # 전체 판매량.xlsx bytes
    purchase_bytes,  # 구매관리(맥산).xlsx bytes — None이면 판매만 집계
    progress_cb=None
)
```

---

## 8. Streamlit session_state 키

| 키 | 내용 | 다운로드 키 | 초기화 버튼 |
|----|------|------------|------------|
| `s1_result` | STEP 1 결과 | `s1_dl_result`, `s1_dl_conv` | `s1_clear` |
| `s2_result` | STEP 2 판매분석 결과 | `s2_dl_ojc` | `s2_clear` |
| `s3_result` | STEP 3 발주계획 결과 | `s3_dl_result` | `s3_clear` |

---

## 9. 이번 세션 주요 수정 내역

### convert_core.py
| 항목 | 내용 |
|------|------|
| `_detect_col_format()` 신규 | 헤더 Row2 자동 감지 — 구매처/거래처 계열 포함 시 ojc 파서 분기 |
| `_parse_buy_date()` 신규 | datetime·YY/MM/DD·YYYY/MM/DD·YYYYMMDD 모두 처리 |
| ojc 컬럼 인덱스 교정 | cells[0]=거래처→무시, cells[2]=날짜, cells[3]=품목코드, cells[7]=수량 |
| `logs` 전파 수정 | `preprocess()` → `_parse_rows_from_sheet()` logs 전달 |
| 진단 메시지 강화 | 파싱 0건 시 샘플 행·스킵 건수를 오류 메시지에 포함 |

### step2_core.py
| 항목 | 내용 |
|------|------|
| `YR_PALETTE` 추가 | 모듈 상단에 누락된 컬러 팔레트 상수 정의 |
| 잘못된 파일 감지 | `케이블 사용내역` 시트 감지 시 명확한 `ValueError` 발생 |
| `sales_data` 연동 | STEP 2 판매분석 결과로 수요 기반 분석 자동 수행 |

### step1_core.py
| 항목 | 내용 |
|------|------|
| `converted_bytes` 반환 | ERP 변환 시 가공파일 bytes 별도 반환 |

### web_app.py
| 항목 | 내용 |
|------|------|
| 탭 6개로 확장 | STEP 2 판매분석 탭 추가 |
| `build_sales_analysis()` 신규 | ojc_filter 연동 + 맥산 생산비중 계산 |
| `parse_production_data()` 신규 | 구매관리(맥산) Excel → 연도별 품목코드별 수량 |
| STEP 3 `ready2` 조건 수정 | `usage_file` 의존성 제거 → `row_file2 is not None` 만 확인 |
| session_state 키 분리 | STEP 2=`s2_result`, STEP 3=`s3_result` (기존 충돌 해결) |
| `purchase_bytes=None` 처리 | 구매관리 미업로드 시 판매 데이터만 분석 (오류 없음) |

---

## 10. 실제 테스트 파일

| 파일 | 경로 | 용도 |
|------|------|------|
| 의왕 생산 내역 25년.xlsx | `C:\Users\AJWorld\Desktop\AI 자료\` | STEP 1 — 맥산 납품 형식 (23~26년 9,485행) |
| 판매량.xlsx | `C:\Users\AJWorld\Desktop\AI 자료\` | STEP 2 — OJC 전체 판매량 |

---

## 11. 디버그 결과 (2026-05-11)

```
STEP1 ✅  연도:['23','24','25','26']  케이블:32종  하우징:30종
STEP2 ✅  분석:1,862품목  OJC행:26,105건  오류:없음
STEP3 ✅  가공파일 입력 → 30,169bytes 발주계획 생성
STEP3 ✅  사용내역.xlsx 잘못 업로드 시 명확한 오류 출력
```

---

## 12. 남은 TODO

- [ ] STEP 2 구매관리(맥산) 파일 업로드 후 생산비중(ratio) 검증
- [ ] 품번 관리 탭 데이터가 STEP 1 재실행 시 정상 반영 확인
- [ ] 26년 데이터 포함 시 STEP 3 발주계획 컬럼 레이아웃 검증
- [ ] `sales_analysis.json` / `inventory.json` .gitignore 추가 권장 (런타임 데이터)

---

## 13. 서버 재시작

```bash
# Windows — venv 직접 실행
"C:\Users\AJWorld\todo-app\AI\연간 발주 계획 새우기\.venv\Scripts\streamlit.exe" ^
  run web_app.py --server.address=0.0.0.0 --server.port=8501
```
