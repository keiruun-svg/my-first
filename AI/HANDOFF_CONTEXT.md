# AJW 프로젝트 인수인계 문서
# 새 대화 시작 시 이 문서를 붙여넣으세요

## 나는 누구인가
나는 (주)에이제이월드(AJW) IT팀의 박정원입니다.
Claude Cowork를 사용해 구매 데이터 가공 자동화 시스템을 개발했습니다.
이 문서는 이전 작업의 전체 컨텍스트를 다음 Claude 세션에 전달하기 위한 것입니다.

---

## 1. 프로젝트 개요

### 목적
구매조회 xlsx → 케이블/하우징 분석용 가공파일(xlsx) 자동 변환

### 산출물 (바탕화면 AI 폴더)
| 파일명 | 용도 |
|--------|------|
| `가공파일_2026_v양식.xlsx` | 2026년 Q1 (2시트) |
| `가공파일_통합_v양식.xlsx` | 2023~2025년 (6시트) |
| `구매조회_가공파일_변환_프롬프트.md` | 재실행용 지시문 |
| `ajw-purchase-convert.skill` | Cowork 스킬 |
| `converter_app.py` | GUI 실행 파일 소스 |
| `build_exe.bat` | .exe 빌드 스크립트 |
| `구매조회_가공자동화_작업보고서.docx` | 작업 보고서 (AJW 양식) |

---

## 2. 입력 파일 규격

### 구매현황 파일 (Q1 등 단일 기간)
- 시트명: `구매현황`
- 헤더: 2행, 데이터: 3행~
- col1=일자(YYYY/MM/DD -N), col5=품목코드, col6=품목명[규격명], col7=수량

### 통합 파일 (다년도)
- 시트명: 임의 (첫 번째 시트 사용)
- 헤더: 1행, 데이터: 2행~
- col2=날짜(YYYYMMDD-N 또는 YY/MM/DD), col4=품목코드, col5=품목명, col7=규격명(분리), col8=수량

---

## 3. OJC 제품 지식

### 커넥터 (3종) × 페롤 (2종) = 6가지 타입
`LC/PC`, `LC/APC`, `SC/PC`, `SC/APC`, `FC/PC`, `FC/APC`
- PC = 청색 부트, APC = 녹색 부트

### 품명 규칙
| 유형 | 패턴 | 타입 추출 방식 |
|------|------|-------------|
| LG DOJC/SOJC | `DOJC-SM-LC/PC-SC/APC-5M` | 커넥터+페롤 이미 결합 |
| LG MOJC | `MOJC-SM-12C-SC/PC-SC/PC-3M` | 결합, 코어수 별도 파싱 |
| KT OJC-A1/C2 | `OJC-A1-SC/LC-SM-3-APC/PC-SP` | 커넥터(pos3)+페롤(pos6) 분리 조합 |
| DROP 단일 | `DROP-CABLE(LC/PC)` | 양끝 동일 |
| DROP 혼합 | `DROP-CABLE(LC/PC-SC/APC)` | 괄호 내 두 타입 |
| DROP 2심 | `DROP-CABLE(LC/PC)-2C` | 양끝 동일, 코어=2 |
| PIGTAIL | `PIGTAIL-LC/APC-SM-12C` | 타입1만 (타입2 없음) |

---

## 4. 케이블 종류 분류 규칙

```python
def derive_kind(품목명, 규격명, core):
    PIGTAIL + MM/OM3  → 'om1-pigtail'
    PIGTAIL           → 'pigtail'
    DROP              → 'drop'
    Optical Cable Parts → 'a2'
    MOJC              → f'b3-{core}c'  # 예: b3-12c
    OJC-C2 + 4코어   → 'b3-4c'
    OJC-C2 그 외     → f'a1-{core}c'  # 예: a1-8c
    규격명에 B3       → 'b3'
    MM(OM3)/OM3       → 'om3'
    -MM-/-MM          → 'om1'
    규격명에 청/적/녹/자 → 'a1-청/적/녹/자'
    나머지            → 'a1'
```

---

## 5. 하우징 수량 계산 규칙 ★ 핵심 ★

### 원칙
```
cps = 코어수  (모든 케이블 동일, 예외 없음)
SIDE A (타입1) → 주색 컬럼 × cps × qty
SIDE B (타입2):
  - SP (코어<2): 주색 컬럼 × cps × qty
  - DP (코어≥2) + 일반 케이블: 적색 컬럼 × cps × qty
  - DP (코어≥2) + 다심(b3-Nc, a1-Nc): 주색 컬럼 × cps × qty  ← 적색 없음!
```

### 색상 결정
```python
MM_KINDS = {'om1', 'om1-pigtail', 'om3'}  # ← OM3도 베이지!

LC/PC + MM_KINDS + 2.0mm → BEIGE (col 35)
LC/PC + 그 외             → 청색
LC/APC                    → 녹색
SC/PC                     → 청색
SC/APC                    → 녹색
FC/PC                     → 흑색
FC/APC                    → 녹색
DP SIDE B                 → 적색 (단, 다심 제외)
```

### 다심 여부 판별
```python
def is_multicore(kind):
    return bool(re.match(r'^(b3|a1)-\d+c$', kind))
# True이면 SIDE B도 주색, 적색 없음
```

### 검증 공식 (AV열)
```
F=1: V×2
F=2: V×F×2
F>2: F×2×V
```

### 컬럼 매핑 (하우징 탭)
```
col 23: 2.0MM LC/PC 청색   col 24: 2.0MM LC/PC 적색
col 25: 2.0MM LC/APC 녹색  col 26: 2.0MM LC/APC 적색
col 27: 2.0MM SC/PC 청색   col 28: 2.0MM SC/PC 적색
col 29: 2.0MM SC/APC 녹색  col 30: 2.0MM SC/APC 적색
col 31: 2.0MM FC/PC 흑색   col 32: 2.0MM FC/PC 적색
col 33: 2.0MM FC/APC 녹색  col 34: 2.0MM FC/APC 적색
col 35: 2.0MM BEIGE (OM1·OM3 LC/PC 전용)
col 36~47: 3.0MM (위와 동일 순서)
col 48: 검증, col 49: 계산수량
```

### 검증 예시
| 품목 | qty | cps | 결과 |
|------|-----|-----|------|
| DOJC-SM-LC/PC-LC/PC | 10 | 2 | 청색20 + 적색20 = 40 |
| DOJC-MM-LC/PC-LC/PC (OM1) | 11 | 2 | BEIGE22 + 적색22 = 44 |
| DOJC-MM(OM3)-LC/PC-LC/PC | 7 | 2 | BEIGE14 + 적색14 = 28 |
| MOJC-SM-12C-SC/PC-SC/PC | 24 | 12 | 청색288 + 청색288 = 576 |
| OJC-C2-8C-SC/PC-SC/PC | 10 | 8 | 청색80 + 청색80 = 160 |
| DROP-CABLE(LC/PC)-2C | 20 | 2 | 청색40 + 적색40 = 80 |

---

## 6. 출력 파일 구조

### 케이블 탭 (24열)
A=품목코드, B=품목명, C=규격명, D=케이블종류, E=파이, F=코어수, G=길이,
H=타입1(정적값), I=타입2(정적값), J~U=월별수량, V=사용량, W=최고제작량, X=소요량

### 하우징 탭 (49열)
A~I=기본정보, J~U=월별수량, V=합계, W~AI=2.0mm하우징, AJ~AU=3.0mm하우징, AV=검증, AW=계산수량

> **중요**: 타입1/2 및 하우징 수량은 반드시 Python 정적값으로 저장 (Excel 수식 금지)

---

## 7. 최종 convert.py 핵심 코드

```python
# MM_KINDS (베이지 적용 대상)
MM_KINDS = {'om1', 'om1-pigtail', 'om3'}

# 다심 여부 (적색 미사용)
def is_multicore(kind):
    return bool(re.match(r'^(b3|a1)-\d+c$', kind))

# 하우징 계산
def calc_housing(row):
    cps = core  # 항상 코어수
    use_red = (core >= 2) and not is_multicore(kind)
    # SIDE A → add_primary(t1, cps*qty)
    # SIDE B → 적색(use_red) 또는 주색(아니면)

# 코어수 파싱 (주요 규칙)
MOJC-SM-NC   → N (후행 '-' 없어도 파싱)
OJC-C2-...-NC → N (뒤에 추가 문자 있어도 파싱)
Optical Cable Parts,...NCCore → N (품목명에서 직접)
DOJC → 2, SOJC → 1
```

---

## 8. 향후 작업 / 미완성 항목

- [ ] `.exe` 파일 빌드 미완성 (build_exe.bat 실행 필요 - Windows에서)
- [ ] GUI 앱 테스트 미완료 (converter_app.py 실행 필요)
- [ ] 2026년 Q2 이후 데이터 추가 예정
- [ ] 새 품목명 패턴 발생 시 extract_types() 업데이트 필요

---

## 9. 다음 작업 시 참고사항

1. 변경사항 발생 시 항상 **프롬프트 문서 + 스킬 파일 동시 업데이트**
2. 스킬 업데이트: `convert.py` 수정 → zip 패키징 → `.skill` 저장 → Cowork 재설치
3. 새 파일 변환 시: `converter_app.py` 실행 또는 스킬 사용
4. 통합 파일은 다년도 자동 감지 → 연도별 시트 생성

---

*이 문서를 새 Claude 세션에 붙여넣으면 이전 작업 맥락을 이어받을 수 있습니다.*
*작성일: 2026-04-24 / 작성: 박정원(IT팀) + Claude*
