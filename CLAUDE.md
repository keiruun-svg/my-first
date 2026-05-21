# AJW 생산자재 발주계획 시스템 — CLAUDE.md

AJW (주)에이제이월드 SCM팀 — 운영 도우미 웹 개발 1단계 - 생산 자재 발주 계획 자동 생성

---

## 기술 스택

| 항목 | 버전 |
|---|---|
| React | 19 |
| TypeScript | 6 |
| Vite | 8 |
| Tailwind CSS | 4 |
| Excel 출력 | ExcelJS 4 |
| Excel 파싱 | SheetJS (xlsx) 0.18 |
| 백엔드 | Supabase (PostgreSQL) + localStorage fallback |

---

## 개발 명령어

```bash
# 작업 디렉토리: react-app/
npm run dev      # Vite 개발 서버 (localhost:5173)
npm run build    # tsc + vite build
npx tsc --noEmit # 타입 체크만
```

---

## 디렉토리 구조

```
react-app/src/
├── App.tsx                     # 탭 라우팅, 초기 Supabase 로드
├── components/
│   ├── Step1.tsx               # STEP 1: ERP 파일 가공
│   ├── Step2.tsx               # STEP 2: 판매 분석
│   ├── Step3.tsx               # STEP 3: 발주계획 생성
│   ├── Dashboard.tsx           # 홈 대시보드
│   ├── SalesAnalysisTab.tsx    # 판매 현황 분석
│   ├── PartNumberGenerator.tsx # OJC 품번 생성기
│   ├── MaterialManager.tsx     # 자재 관리 (관리자 전용)
│   ├── InventoryReconciliationTab.tsx  # 재고 대사 (EMP ↔ 이카운트)
│   ├── ImportTab.tsx           # 수입 관리 (2단계 구현 예정)
│   └── Settings.tsx            # 파라미터 설정
└── lib/
    ├── types.ts                # 공용 타입: AppSettings, Metadata, Inventory
    ├── supabase.ts             # Supabase 클라이언트 + CRUD
    ├── step1Core.ts / step2Core.ts / step3Core.ts
    ├── ojcFilter.ts / ojcAutoDetect.ts
    ├── aggregate/
    │   ├── salesAgg.ts         # 타입별 판매·생산 집계, CAGR
    │   └── pivot.ts            # ERP → 연도별 피벗
    ├── parse/
    │   ├── parseERP.ts         # ERP 구매현황 파싱
    │   ├── parseSales.ts / parseDetailedSales.ts
    │   └── classify.ts         # deriveKind, derivePai, deriveCore, deriveLength
    └── output/
        ├── writeGaong.ts       # 가공파일.xlsx (STEP 1)
        ├── writeSalesAgg.ts    # 판매분석.xlsx (STEP 2)
        └── writeStep3Excel.ts  # 발주계획.xlsx (STEP 3)
```

---

## 탭 구조 (App.tsx ALL_TABS)

| id | 설명 |
|---|---|
| home | 홈 대시보드 |
| step1 | ERP 파일 가공 → 가공파일.xlsx |
| step2 | 판매 분석 → 판매분석.xlsx |
| step3 | 발주계획 생성 → 발주계획.xlsx |
| sales | 판매 현황 분석 |
| partnum | OJC 품번 생성기 |
| material | 자재 관리 (관리자 전용, 비밀번호 `VITE_ADMIN_PASS`) |
| recon | 재고 관리 — EMP ↔ 이카운트 재고 대사 |
| import | 수입 관리 (EMP 코드 생성 + 발주·입고) — 구현 예정 |
| settings | 파라미터 & 양식 설정 |

---

## 3-STEP 데이터 흐름

```
[STEP 1] ERP 파일 (.xls/.xlsx)
  → parseERP → buildPivot → classify (케이블 종류·파이·코어·길이)
  → 가공파일.xlsx 출력 (집계 시트 + 코드매핑 시트)
  → metadata(Cable/Housing/Ferrule) Supabase 저장

[STEP 2] 판매파일 + 생산파일
  → parseSales + aggregateSales → CAGR 계산
  → 판매분석.xlsx 출력 / SalesAggResult Supabase 저장

[STEP 3] 가공파일 + 판매분석 + metadata + inventory
  → step3Core 발주계획 계산:
      안전재고 = 3개년 피크평균 × (LT/30) × safetyK
      발주필요량 = max(0, 예측연간 + 안전재고 - 현재고 - 기발주)
  → 발주계획.xlsx 출력 (케이블사용내역, 하우징사용내역, 품번별발주집계, 월별발주계획)
```

---

## Supabase

**테이블: `app_data`** (단일 키-값)

| id | 내용 |
|---|---|
| settings | AppSettings (k=1.5, lead_time=60, colors) |
| metadata | Metadata { cable, housing, ferrule } |
| inventory | Inventory { cable, housing, ferrule } |
| sales | SalesAnalysis |
| sales_agg | SalesAggResult |
| ojc_products | OjcProduct[] |

- `CAN_WRITE = import.meta.env.DEV` — **프로덕션은 읽기 전용**
- Supabase 없을 때 localStorage 자동 폴백
- `.env`: `VITE_SUPABASE_URL`, `VITE_SUPABASE_KEY`, `VITE_ADMIN_PASS`

---

## 핵심 타입 (lib/types.ts)

```typescript
interface Metadata {
  cable:   Record<string, CableMeta>                   // key: "파이|종류"
  housing: Record<string, HousingComp | HousingComp[]> // key: "파이|타입" (단일 or 다중 부품 배열)
  ferrule: Record<string, FerruleMeta>                 // key: "LC/PC" 등
}
interface Inventory {
  cable:   Record<string, { 현재고: number; 기발주: number }>
  housing: Record<string, { 현재고: number; 기발주: number } | { 현재고: number; 기발주: number }[]>
  ferrule: Record<string, { 현재고: number; 기발주: number }>
}
// Step3Row — step3Core.ts에 정의 (types.ts 아님)
interface Step3Row {
  type:           'cable' | 'housing'
  key:            string
  label:          string
  pai:            string
  unit:           string
  byYear:         Record<string, { annual: number; peak: number; monthly: number[] }>
  years:          string[]
  latestAnnual:   number
  latestPeak:     number
  forecastAnnual: number  // latestAnnual × (1 + CAGR); CAGR 없으면 latestAnnual
  appliedCagr:    number  // CAGR 미적용 시 0
  품번:            string
  품명:            string
  구매처:          string
  리드타임:        number
  안전재고:        number
  현재고:          number
  기발주:          number
  발주필요량:      number
  isSubRow:       boolean  // 하우징 다중 부품의 2번째+ 행
}
```

---

## 케이블 분류 로직 (classify.ts)

`deriveKind(품명, 규격)` → 케이블 타입 키:
`pigtail` / `om1-pigtail` / `drop` / `b3` (MOJC) / `om3` / `om1` / `a1` / `a1-청` / `a1-적` / `a1-녹` / `a1-자` / `a2`

`derivePai()` → `'2.0mm'` / `'3.0mm'` / `'0.9mm'` (mm 포함 문자열)
`deriveCore()` → 1, 2, 4, 6, 8, 12, 24 등
`deriveLength()` → 길이(m) 정수

---

## OJC 도메인 지식

**OJC 품번 형식:** `{OJC타입}{코어타입}{코어수}{재질}{커넥터}{마킹}-{길이코드}`
예: `C-ALNN-3` = SM 1코어 PVC SC/PC 마킹없음 3m

**EMP 로트 코드 형식:** `{기본코드} ({업체코드}-{YYMM}-{발주차수}_{입고회차})`
예: `14-K-107 (FLC-2601-01_02)`

이카운트 코드 Set 기준 역방향 매칭으로 기본코드 추출 → `resolveBaseCode()` in InventoryReconciliationTab.tsx

---

## 재고 대사 (InventoryReconciliationTab)

입력 4종:
1. EMP 재고현황 (`.xls`) — 헤더: 상품코드|상품명|창고|다중로케이션|재고수량
2. 이카운트 수불부 (`.xlsx` 복수) — 마지막 `합계` 행 마지막 컬럼 = 현재고
3. 미출하 현황 (`.xlsx`) — 시트 `미출하현황`, col[6]=코드, col[10]=잔량
4. 일일 입출고 확인서 (`.xlsx`) — 날짜별 시트, 행7=헤더, col[1]='출고' 행만

공식: `EMP조정 = EMP재고 − 출고합산` / `EC조정 = EC재고 + 미출하` / `차이 = EMP조정 − EC조정`

ZZ-ZZ 로케이션 자동배정: `assignLocation(품명)` — 15가지 규칙 (D1, D2, M1, EX 등)

---

## 코딩 컨벤션

- 한국어 변수명 사용: `품번`, `품명`, `구매처`, `리드타임`, `현재고`, `기발주`, `안전재고`
- Excel 출력: ExcelJS — `ALL_BORDERS`, `fill()`, `font()` 공통 헬퍼 재사용
- Excel 파싱: SheetJS `read()` → `utils.sheet_to_json(ws, { header: 1 })` → `unknown[][]`
- Supabase 저장: `setState(x)` 먼저, `saveToSupabase(x)` 비동기 후행

### Excel 다운로드 파일명 규칙

데이터를 다루는 탭은 기본으로 Excel 다운로드 기능을 포함한다.

파일명 형식: **`기능명_${today()}.xlsx`** (날짜+시간 필수)

```typescript
import { downloadXlsx, today } from '../lib/download'
// today() → "YYYYMMDD_HHmm" 형식

downloadXlsx(buffer, `발주계획_${today()}.xlsx`)   // 예: 발주계획_20260521_1430.xlsx
```

예외: `재고대사`는 기준일자 + 작성일 조합 → `재고대사_${date}_${today()}.xlsx`

---

## 현재 구현 상태 (2026-05-20)

| 탭 | 상태 |
|---|---|
| STEP 1/2/3 | 🔄 핵심 기능 구현 → 피드백 및 수정 중 |
| 판매 현황 분석 | 🔄 핵심 기능 구현 → 피드백 및 수정 중 |
| 품번 생성기 | ✅ 완료 |
| 자재 관리 | ✅ 완료 |
| 재고 관리 (대사) | 🔄 핵심 기능 구현 → 피드백 및 수정 중 (고도화 2단계 예정) |
| 수입 관리 | 🔲 플레이스홀더 (2단계 구현 예정) |

## 운영 도우미 웹 개발 2단계 (미구현)

1. **수입 관리**: EMP 코드 생성기 + 업체코드 관리 (Supabase `vendors` 테이블) + 발주 현황 & 입고 관리 + 리드타임 분석
2. **재고 관리 고도화**: Supabase `recon_history` 누적 저장 + 이력 차트
3. **앱 리브랜딩**: "AJW SCM 어시스턴트" + STEP 1/2/3 드롭다운 네비게이션

## 운영 도우미 웹 개발 3단계 (미구현)

완제품 수입 어시스턴트 — 수입 관리 탭 내 서브섹션으로 구현

1. **수요 패턴 분석**: 3년치 월별 판매 경향 → 정기/계절/간헐 수요 분류, 발주 타이밍 권고
2. **수익성 분석**: 수입가 · 생산가 · 판매가 비교 → 수입 / 생산 / 철수 의사결정 지원

데이터 소스:
- 판매 데이터: `SalesAggResult` (STEP 2 저장값)
- 현재고: 이카운트 수불부 또는 ERP 파일 업로드 (양쪽 호환)
- 단가: 원가 파일 업로드 (파서 별도 개발 필요)

---

## 문서 유지보수 가이드

### 어떤 작업을 했을 때 무엇을 업데이트하나

| 작업 | 업데이트 대상 |
|------|-------------|
| 새 탭 추가 | CLAUDE.md 탭 구조 + 디렉토리 구조 + 구현 상태 |
| 새 컴포넌트/파일 추가 | CLAUDE.md 디렉토리 구조 |
| 타입 변경 (types.ts, step3Core.ts 등) | CLAUDE.md 핵심 타입 섹션 (반드시 파일 직접 읽어서 복사) |
| 공식·계산 로직 변경 | CLAUDE.md 해당 섹션 + README.md |
| 단계 완료 (2단계, 3단계 등) | CLAUDE.md 구현 상태 + memory/project.md 구현 현황 |
| 파싱 규칙 변경 (입력 파일 컬럼 등) | CLAUDE.md 재고 대사 섹션 |
| 환경변수 추가/변경 | CLAUDE.md Supabase 섹션 + .env.example + README.md |
| Supabase 스키마 변경 | supabase/schema.sql + README.md |

### Claude에게 요청하는 방법

작업이 끝난 후 다음 중 하나를 말하면 됩니다:

- **"문서 업데이트해줘"** — 이번 세션에서 변경된 내용을 CLAUDE.md와 memory/project.md에 반영
- **"메모리 최신화해줘"** — memory/project.md의 구현 현황·작업 이력 갱신
- **"CLAUDE.md 검토해줘"** — 실제 코드와 대조해서 오래된 내용 찾기

### 파일별 역할 요약

| 파일 | 역할 | 업데이트 빈도 |
|------|------|-------------|
| `CLAUDE.md` | Claude 세션용 프로젝트 레퍼런스 (코드 구조·타입·로직) | 기능 추가·변경 시 |
| `README.md` | 새 환경 세팅 가이드 (빠른 시작·탭 목록·워크플로우) | 큰 변경 시 |
| `supabase/schema.sql` | DB 초기화 SQL | 스키마 변경 시 |
| `memory/project.md` | 세션 간 컨텍스트 전달 (구현 현황·계획) | 단계 완료·계획 변경 시 |
