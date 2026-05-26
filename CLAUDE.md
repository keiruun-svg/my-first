# AJW SCM 어시스턴트 — CLAUDE.md

AJW (주)에이제이월드 SCM팀 — 운영 도우미 웹 (2단계 완료)

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
├── App.tsx                     # 탭 라우팅, NAV_TABS / STEP_TABS
├── components/
│   ├── Dashboard.tsx           # 홈 대시보드
│   ├── Step1.tsx               # STEP 1: ERP 파일 가공
│   ├── Step2.tsx               # STEP 2: 판매 분석
│   ├── Step3.tsx               # STEP 3: 발주계획 생성
│   ├── SalesAnalysisTab.tsx    # 판매 현황 분석
│   ├── PartNumberGenerator.tsx # OJC 품번 생성기 (서브탭: auto/manual/rules/emp/location)
│   ├── EmpCodeGenerator.tsx    # EMP 로트코드 생성기 (품번 생성기 서브탭)
│   ├── EmpLocationGenerator.tsx # EMP 창고 로케이션 생성기 (품번 생성기 서브탭)
│   ├── MaterialManager.tsx     # 자재 관리 (관리자 전용)
│   ├── InventoryReconciliationTab.tsx  # 재고 대사 (EMP ↔ 이카운트)
│   ├── ImportTab.tsx           # 수입 관리 (서브탭: 발주계획/간헐적수요/발주현황)
│   ├── ImportOrdersView.tsx    # 발주 현황 (업체+차수 단위 등록·조회)
│   └── Settings.tsx            # 파라미터 설정
└── lib/
    ├── types.ts                # 공용 타입: AppSettings, Metadata, Inventory
    ├── supabase.ts             # Supabase 클라이언트 + CRUD + localStorage 헬퍼
    ├── download.ts             # downloadXlsx, today(), pickSaveFile
    ├── assignLocation.ts       # assignLocation(품명) → 창고 코드 (15규칙)
    ├── ojcFilter.ts            # classifyOjc, classifyOjcDetailed, COLOR_MAP_OJC
    ├── ojcAutoDetect.ts
    ├── step1Core.ts / step2Core.ts / step3Core.ts
    ├── aggregate/
    │   ├── salesAgg.ts         # 타입별 판매·생산 집계, CAGR
    │   └── pivot.ts            # ERP → 연도별 피벗
    ├── parse/
    │   ├── parseERP.ts         # ERP 구매현황 파싱
    │   ├── parseSales.ts / parseDetailedSales.ts
    │   ├── parseItemCodes.ts   # 이카운트/EMP 품목 리스트 파싱 → ItemCode[]
    │   ├── parseStockFile.ts   # 이카운트 수불부/EMP 재고현황 자동감지 → Map<code, qty>
    │   └── classify.ts         # deriveKind, derivePai, deriveCore, deriveLength
    └── output/
        ├── writeGaong.ts       # 가공파일.xlsx (STEP 1)
        ├── writeSalesAgg.ts    # 판매분석.xlsx (STEP 2)
        └── writeStep3Excel.ts  # 발주계획.xlsx (STEP 3)
```

---

## 탭 구조 (App.tsx)

### NAV_TABS (상단 네비게이션)

| id | 설명 |
|---|---|
| home | 홈 대시보드 |
| steps | 📋 생산자재 발주계획 (드롭다운 — STEP_TABS 하위) |
| sales | 판매 현황 분석 |
| partnum | OJC 품번 생성기 |
| material | 자재 관리 (관리자 전용, 비밀번호 `VITE_ADMIN_PASS`) |
| recon | 재고 관리 — EMP ↔ 이카운트 재고 대사 |
| import | 수입 관리 (완제품 수입 발주계획 + 발주 현황) |
| settings | 파라미터 & 양식 설정 |

### STEP_TABS (생산자재 발주계획 하위)

| id | 설명 |
|---|---|
| step1 | ERP 파일 가공 → 가공파일.xlsx |
| step2 | 판매 분석 → 판매분석.xlsx |
| step3 | 발주계획 생성 → 발주계획.xlsx |

```typescript
type StepId = 'step1' | 'step2' | 'step3'
type NavId  = typeof NAV_TABS[number]['id']
type TabId  = Exclude<NavId, 'steps'> | StepId
function isStepId(id: string): id is StepId
```

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

## 수입 관리 탭 (ImportTab)

완제품 OJC 수입 관련 3개 서브탭.

### 서브탭 1: 발주계획

- 입력: 재고파일 (`parseStockFile` — 이카운트/EMP 자동감지) + 판매현황 파일
- OJC 품목 필터 (`classifyOjc`) → 월간 집계 → 월간최고 기반 커버리지 계산
- 커버리지 = 현재고 ÷ 월간최고 / 발주필요량 = (월간최고 × 목표개월) − 현재고
- 간헐적 수요 필터: 판매 발생 월수 < 기준(기본 3개월) → 발주 대상 제외

### 서브탭 2: 간헐적 수요

- 간헐적 수요 품목만 별도 표시
- 연도별 (판매횟수 / 판매총량 / 최고단월) 요약

### 서브탭 3: 발주 현황 (ImportOrdersView)

- 발주건 = 업체 + 연도(YY) + 차수 (예: `FLC 26-05차`)
- 차수 자동 제안 (같은 업체·연도 마지막 차수 +1)
- 입고 처리 → 실제입고일 기입 → 리드타임(일) 자동 계산
- D-day 카운트다운 / 지연 시 빨강 표시

**`ImportOrder` 타입**
```typescript
interface ImportOrder {
  id:           string   // 고유 ID
  vendorCode:   string   // 'FLC'
  vendorName:   string   // 'FIBERCAN'
  year:         string   // '26'
  seq:          number   // 5
  orderDate:    string   // 'YYYY-MM-DD'
  expectedDate: string   // 'YYYY-MM-DD'
  actualDate?:  string   // 입고 완료 시
}
// localStorage key: 'ajw_import_orders'
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
| item_codes | ItemCode[] (이카운트/EMP 품목 리스트) |

**테이블: `vendors`** (업체코드 관리)
```
code  text  PRIMARY KEY
name  text
```

**localStorage 전용** (Supabase 미연동)
```
ajw_detailed_sales   DetailedSalesRow[]   # 판매현황 분석 → 수입 발주계획 공유
ajw_import_orders    ImportOrder[]        # 완제품 수입 발주 현황
```

- `CAN_WRITE = true` — 관리자 게이트는 비밀번호(`VITE_ADMIN_PASS`)로 접근 제어
- Supabase 연동 토글: `getSyncEnabled()` / `setSyncEnabled()` (localStorage)
- `.env`: `VITE_SUPABASE_URL`, `VITE_SUPABASE_KEY`, `VITE_ADMIN_PASS`

---

## 핵심 타입 (lib/types.ts)

```typescript
interface Metadata {
  cable:   Record<string, CableMeta>
  housing: Record<string, HousingComp | HousingComp[]>
  ferrule: Record<string, FerruleMeta>
}
interface Inventory {
  cable:   Record<string, { 현재고: number; 기발주: number }>
  housing: Record<string, { 현재고: number; 기발주: number } | { 현재고: number; 기발주: number }[]>
  ferrule: Record<string, { 현재고: number; 기발주: number }>
}
```

**Step3Row** (step3Core.ts)
```typescript
interface Step3Row {
  type: 'cable' | 'housing'; key: string; label: string; pai: string; unit: string
  byYear: Record<string, { annual: number; peak: number; monthly: number[] }>
  years: string[]; latestAnnual: number; latestPeak: number
  forecastAnnual: number; appliedCagr: number
  품번: string; 품명: string; 구매처: string; 리드타임: number
  안전재고: number; 현재고: number; 기발주: number; 발주필요량: number
  isSubRow: boolean
}
```

**ItemCode** (lib/supabase.ts, lib/parse/parseItemCodes.ts)
```typescript
interface ItemCode { code: string; name: string; spec: string }
```

---

## 케이블 분류 로직 (classify.ts)

`deriveKind(품명, 규격)` → 케이블 타입 키:
`pigtail` / `om1-pigtail` / `drop` / `b3` (MOJC) / `om3` / `om1` / `a1` / `a1-청` / `a1-적` / `a1-녹` / `a1-자` / `a2`

`derivePai()` → `'2.0mm'` / `'3.0mm'` / `'0.9mm'`
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

ZZ-ZZ 로케이션 자동배정: `assignLocation(품명)` in `lib/assignLocation.ts` — 15가지 규칙

---

## 코딩 컨벤션

- 한국어 변수명 사용: `품번`, `품명`, `구매처`, `리드타임`, `현재고`, `기발주`, `안전재고`
- Excel 출력: ExcelJS — `ALL_BORDERS`, `fill()`, `font()` 공통 헬퍼 재사용
- Excel 파싱: SheetJS `read()` → `utils.sheet_to_json(ws, { header: 1 })` → `unknown[][]`
- Supabase 저장: `setState(x)` 먼저, `saveToSupabase(x)` 비동기 후행

### Excel 다운로드 파일명 규칙

```typescript
import { downloadXlsx, today } from '../lib/download'
// today() → "YYYYMMDD_HHmm"

downloadXlsx(buffer, `발주계획_${today()}.xlsx`)
// 예외: 재고대사 → `재고대사_${date}_${today()}.xlsx`
```

---

## 현재 구현 상태 (2026-05-26)

| 탭 | 상태 |
|---|---|
| STEP 1/2/3 (생산자재 발주계획) | 🔄 핵심 기능 구현 → 피드백 및 수정 중 |
| 판매 현황 분석 | 🔄 핵심 기능 구현 → 피드백 및 수정 중 |
| 품번 생성기 | ✅ 완료 (서브탭: OJC 자동/수동/규칙편집/EMP코드/EMP로케이션) |
| 자재 관리 | ✅ 완료 |
| 재고 관리 (대사) | 🔄 핵심 기능 구현 → 피드백 수정 중 (이력 누적 고도화 예정) |
| 수입 관리 | ✅ 완료 (발주계획/간헐적수요/발주현황 서브탭) |

## 미구현 항목

| 항목 | 우선순위 |
|---|---|
| 재고 관리 고도화 | 중 — `recon_history` 누적 저장 + 불일치 추이 차트 |
| EMP 로케이션 생성기 KT향 | 낮음 — KT향 품명 예시 필요 |
| 리드타임 집계 분석 | 낮음 — 발주 이력 축적 후 |
| 3단계 수익성 분석 | 낮음 — 원가 파일 파서 개발 필요 |

---

## 문서 유지보수 가이드

| 작업 | 업데이트 대상 |
|------|-------------|
| 새 탭 추가 | CLAUDE.md 탭 구조 + 디렉토리 구조 + 구현 상태 |
| 새 컴포넌트/파일 추가 | CLAUDE.md 디렉토리 구조 |
| 타입 변경 | CLAUDE.md 핵심 타입 섹션 (반드시 파일 직접 읽어서 복사) |
| 공식·계산 로직 변경 | CLAUDE.md 해당 섹션 |
| 단계 완료 | CLAUDE.md 구현 상태 + memory/project.md |

### Claude에게 요청하는 방법

- **"문서 업데이트해줘"** — CLAUDE.md + memory/project.md 갱신
- **"메모리 최신화해줘"** — memory/project.md 구현 현황·작업 이력 갱신
- **"CLAUDE.md 검토해줘"** — 실제 코드와 대조해서 오래된 내용 찾기
