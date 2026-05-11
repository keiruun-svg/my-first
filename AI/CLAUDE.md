# AJW 구매조회 데이터 가공 자동화

## 프로젝트 요약
(주)에이제이월드(AJW) 구매조회 xlsx → 케이블/하우징 분석용 가공파일 자동 변환 시스템.
담당자: 박정원 (IT팀) | 최종 수정: 2026-04-24

## 파일 구조
```
AI/  (바탕화면)
├── converter_app.py                    # GUI 실행 파일 (tkinter + 변환 로직 통합)
├── build_exe.bat                       # Windows .exe 빌드 스크립트
├── 가공파일_2026_v양식.xlsx              # 2026 Q1 결과 (2시트)
├── 가공파일_통합_v양식.xlsx              # 2023-2025 결과 (6시트)
├── 구매조회_가공파일_변환_프롬프트.md     # Claude 재실행용 지시문
├── ajw-purchase-convert.skill          # Cowork 스킬
├── 구매조회_가공자동화_작업보고서.docx    # AJW 양식 보고서
├── HANDOFF_CONTEXT.md                 # 이전 세션 인수인계 문서
└── CLAUDE.md                          # 이 파일 (Claude Code용)
```

## 입력 파일 두 종류

### 구매현황 파일
- 시트: `구매현황`, 헤더: 2행, 데이터: 3행~
- col1=날짜(`YYYY/MM/DD -N`), col5=품목코드, col6=`품목명[규격명]`, col7=수량

### 통합 파일 (다년도)
- 시트: 첫 번째 시트, 헤더: 1행, 데이터: 2행~
- col2=날짜(`YYYYMMDD-N` 또는 `YY/MM/DD`), col4=품목코드, col5=품목명, col7=규격명(분리), col8=수량

## 출력 구조 (연도별 2시트)
- `YY년_케이블` (24열): 품목코드~타입2, 월별수량×12, 사용량·최고제작량·소요량
- `YY년 하우징` (49열): 기본정보, 월별수량×12, 합계, 하우징키트×26, 검증·계산수량

---

## OJC 도메인 지식

### 커넥터 × 페롤 = 6가지 타입
`LC/PC`(청) `LC/APC`(녹) `SC/PC`(청) `SC/APC`(녹) `FC/PC`(흑) `FC/APC`(녹)

### 품명 → 타입1/타입2 추출 규칙
| 품명 패턴 | 추출 방법 |
|----------|---------|
| `DOJC-SM-LC/PC-SC/APC-5M` | 3·4번째 구분자 사이 값 (커넥터+페롤 결합) |
| `MOJC-SM-12C-SC/PC-SC/PC` | 동일 (코어수 NC 별도 파싱) |
| `OJC-A1-SC/LC-SM-3-APC/PC-SP` | pos3(커넥터)+pos6(페롤) 분리 조합 |
| `DROP-CABLE(LC/PC)` | 괄호 내, 양끝 동일 |
| `DROP-CABLE(LC/PC-SC/APC)` | 괄호 내 두 타입 |
| `PIGTAIL-LC/APC-SM-12C` | 타입1만 (타입2 없음) |

### 케이블 종류 분류
```python
PIGTAIL+MM/OM3 → 'om1-pigtail'
PIGTAIL        → 'pigtail'
DROP           → 'drop'
Optical Cable Parts → 'a2'
MOJC           → f'b3-{core}c'      # 예: b3-12c
OJC-C2 core=4  → 'b3-4c'
OJC-C2 기타    → f'a1-{core}c'     # 예: a1-8c
규격명에 B3    → 'b3'
MM(OM3)/OM3    → 'om3'
-MM-/-MM       → 'om1'
규격명 색상글자 → 'a1-청/적/녹/자'
나머지         → 'a1'
```

---

## 하우징 계산 핵심 규칙

```python
MM_KINDS = {'om1', 'om1-pigtail', 'om3'}  # 베이지(BEIGE) 적용 대상

def is_multicore(kind):
    """MOJC/OJC-C2: b3-Nc / a1-Nc 패턴 → 적색 미사용"""
    return bool(re.match(r'^(b3|a1)-\d+c$', kind))

# 계산 원칙
cps = core  # 항상 코어수 (예외 없음)

# SIDE A (타입1) → 주색 × cps × qty
#   LC/PC + MM_KINDS + 2.0mm → BEIGE (col35)
#   LC/PC 그 외 → 청색
#   LC/APC → 녹색, SC/PC → 청색, SC/APC → 녹색, FC/PC → 흑색, FC/APC → 녹색

# SIDE B (타입2)
#   core < 2 (SP)        → 주색 (SIDE A와 동일 로직)
#   core >= 2 + 일반      → 적색 (DOJC, OJC-A1-DP, DROP-2C)
#   core >= 2 + 다심      → 주색 (MOJC, OJC-C2: 적색 없음!)
```

### 하우징 컬럼 매핑
```
col23/36: LC/PC 청색(주)  col24/37: LC/PC 적색(부)
col25/38: LC/APC 녹색     col26/39: LC/APC 적색
col27/40: SC/PC 청색      col28/41: SC/PC 적색
col29/42: SC/APC 녹색     col30/43: SC/APC 적색
col31/44: FC/PC 흑색      col32/45: FC/PC 적색
col33/46: FC/APC 녹색     col34/47: FC/APC 적색
col35: BEIGE (2.0mm OM1·OM3 LC/PC 전용)
col48: 검증  col49: 계산수량
```

### 검증 수식 (AV열)
`=IF(F=1,V*2,IF(F=2,V*F*2,IF(F>2,F*2*V)))`

### 케이스별 정답
| 품목 | qty | cps | 계산수량 |
|------|-----|-----|---------|
| DOJC-SM-LC/PC-LC/PC | 10 | 2 | 청색20 + 적색20 = **40** |
| DOJC-MM-LC/PC (OM1) | 11 | 2 | BEIGE22 + 적색22 = **44** |
| DOJC-MM(OM3)-LC/PC  | 7  | 2 | BEIGE14 + 적색14 = **28** |
| MOJC-12C-SC/PC      | 24 | 12 | 청색288 + 청색288 = **576** |
| OJC-C2-8C-SC/PC     | 10 | 8  | 청색80 + 청색80 = **160** |
| DROP-2C-LC/PC       | 20 | 2  | 청색40 + 적색40 = **80** |

---

## 코어수 파싱 주의사항
```python
MOJC-SM-NC    → re.search(r'MOJC-(?:SM|MM)-(\d+)C', p)   # 후행 '-' 불필요
OJC-C2-...-NC → re.search(r'OJC-C2-.*-(\d+)C', p)        # 뒤에 추가 문자 허용
Optical Parts → re.search(r'(\d+)Core', p)                # 1Core, 2Core 파싱
DOJC=2, SOJC=1
```

---

## 중요 원칙 (반드시 지킬 것)

1. **타입1/2 및 하우징 수량 → Python 정적값 저장** (Excel 수식 금지)
2. **변경 발생 시 3가지 동시 업데이트:**
   - `구매조회_가공파일_변환_프롬프트.md`
   - `scripts/convert.py`
   - `.skill` 재패키징
3. **OM3도 BEIGE** (`MM_KINDS = {om1, om1-pigtail, om3}`)
4. **다심(MOJC, OJC-C2) SIDE B → 주색** (적색 없음)

## TODO
- [ ] `build_exe.bat` 실행하여 .exe 빌드 (Windows에서 직접 실행)
- [ ] `converter_app.py` GUI 테스트

---

## OJC 판매내역 필터링

### 스크립트 위치
`AI/연간 발주 계획 새우기/ojc_filter.py`

### 용도
전체 판매량 Excel에서 OJC 완제품만 추출 → 정리된 Excel 생성

### 실행
```bash
python ojc_filter.py <입력파일.xlsx> [출력파일.xlsx]
# 출력 생략 시 → OJC_판매량_정리.xlsx
```

### 입력 파일 컬럼 규격
| 컬럼명 | 내용 |
|--------|------|
| 년/월/일 | 날짜 |
| 품목코드 | 제품 코드 |
| 품목명 | 제품명 (prefix로 OJC 여부 판별) |
| 규격명 | 규격 |
| 수량 | 판매 수량 (음수=취소, 자동 제거) |
| 거래처 | 고객사 |
| 창고 | 출고 창고 |

### OJC 분류 기준 (prefix 매칭)
```python
OJC_PREFIXES = {
    'KT OJC'            : ('OJC-A1-', 'OJC-C2-'),
    'LG OJC'            : ('SOJC-', 'DOJC-', 'MOJC-'),
    'DROP'              : ('DROP-CABLE',),
    'PIGTAIL'           : ('PIGTAIL-',),
    'Optical Cable Parts': ('Optical Cable Parts',),
    'DX-MM'             : ('DX-MM',),
}
# Distribution 케이블 (별도 탭): ('Distribution-CABLE', 'DISTRIBUTION CABLE')
```

### 출력 구조
- **OJC 판매량** 시트: 년/월/일, OJC종류, 품목코드, 품목명, 규격명, 수량, 거래처, 창고
  - OJC종류별 배경색 구분 (KT=파랑, LG=녹색, DROP=노랑, PIGTAIL=주황)
- **Distribution 케이블** 시트: 동일 구조, 별도 관리

### 처리 로직
1. 전체 DataFrame 로드
2. 품목명 prefix 매칭으로 OJC종류 분류
3. 음수/0 수량(취소건) 자동 제거
4. 품명 형식 검증 (KT: 7개 이상 `-` 구분, LG: 4개 이상)
5. Excel 저장 + 스타일 적용

### 현재 데이터 파일
| 파일 | 내용 |
|------|------|
| `판매량_정리.xlsx` | 23~26년 OJC 완제품 정리본 (15개 시트) |
| `B1ILLHXVWW3VU1C.xlsx` | 2025.01~2026.05 구매조회 원본 (1,755행) |
| `가공파일_통합_v양식.xlsx` | 2023~2025 케이블/하우징 가공파일 (6시트) |

### ojc_filter.py를 호출하는 방법 (Python 내부)
```python
from ojc_filter import run
ojc_df, dist_df = run('전체판매내역.xlsx', 'OJC만_정리.xlsx')
```
