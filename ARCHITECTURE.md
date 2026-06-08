# AJW SCM 어시스턴트 — ARCHITECTURE.md

React 19 / TypeScript / Vite / Tailwind CSS 4 / ExcelJS / SheetJS / Supabase

---

## 탭 구조

```
App.tsx
├── home        → Dashboard.tsx          (KPI·D-day·재고대사 현황·CAGR)
├── steps
│   ├── step1   → Step1.tsx             (ERP 파일 가공·Metadata 생성)
│   ├── step2   → Step2.tsx             (판매 분석·SalesAgg 생성)
│   └── step3   → Step3.tsx             (발주계획.xlsx 생성)
├── sales       → SalesAnalysisTab.tsx  (판매현황 분석·연도 필터·다운로드)
├── recon       → InventoryReconciliationTab.tsx  (EMP↔이카운트 재고대사)
├── import      → ImportTab.tsx
│   ├── plan          (발주계획)
│   ├── intermittent  (간헐적 수요)
│   ├── orders  → ImportOrdersView.tsx  (발주현황 CRUD·D-day)
│   └── profitability → ProfitabilityAnalysis.tsx  (수익성 분석)
├── partnum     → PartNumberGenerator.tsx
│   ├── OJC 품번 자동/수동 생성
│   └── EmpCodeGenerator.tsx            (EMP 로트코드 생성)
└── admin       (VITE_ADMIN_PASS 게이트)
    ├── admin_material → MaterialManager.tsx
    ├── admin_ojcrules → OjcRulesEditor.tsx
    └── admin_settings → Settings.tsx
```

---

## 데이터 흐름

### 생산자재 발주계획 (3-Step)

```
[STEP 1]
ERP 파일(.xlsx) ──→ parseERP() ──→ ErpRow[]
                                    ↓
                              buildPivot() ──→ 가공파일.xlsx (다운로드)
                                    ↓
                              runStep1() ──→ 신규 품번 타입 감지
                                    ↓
                         setMetadata() + saveMetadata()
                         └─ localStorage: ajw_metadata
                         └─ Supabase: app_data[id='metadata']

[STEP 2]
판매분석.xlsx ──→ buildSalesAgg() ──→ SalesAggResult
                                       ↓
                              setSalesAgg() + saveSalesAgg()
                              └─ localStorage: ajw_sales_agg

[STEP 3]
가공파일 + SalesAgg + Metadata + Inventory
    ↓
buildStep3Plan() ──→ Step3Row[]
    안전재고 = monthlyAvg × (LT/30) × safetyK
    발주필요량 = max(0, forecastAnnual - 현재고 - 기발주)
    ↓
발주계획.xlsx (다운로드)
```

### 완제품 수입 발주계획

```
재고파일 ──→ parseStockFile()  ──→ Map<code, qty>   (EMP/이카운트 자동감지)
판매파일 ──→ parseDetailedSalesFile() ──→ DetailedSalesRow[]
                                          └─ IndexedDB: ajw_detailed_sales
    ↓
classifyOjc(name) 필터링
    ↓
PlanRow[]
    커버리지 = 현재고 ÷ 월간최고
    발주필요량 = (월간최고 × targetMonths) − 현재고
    상태: urgent / warning / ok / intermittent / no_data
    ↓
완제품수입발주계획.xlsx (다운로드)
```

### 수익성 분석

```
발주계획파일 + 맥산원가파일 + FLC원가파일
    ↓
parseCostFile() ──→ Map<code, CostEntry>
    ↓
DetailedSalesRow[] (IndexedDB) ──→ 판매가 (계약단가 우선 → 판매이력 평균)
    ↓
마진율 = (판매가 − 원가) / 판매가 × 100
    원가 적용: KT향 → 표준원가 / LG향·기타 → 생산원가
    importThreshold = thresholdEnabled ? 105 : 100  (맥산 생산 가중치 5%, 기본 OFF)
    ↓
수익성분석.xlsx (다운로드)
```

---

## 저장소 분리 기준

| 데이터 | 저장소 | 키 |
|---|---|---|
| AppSettings | localStorage + Supabase | `ajw_settings` |
| Metadata (품번) | localStorage + Supabase | `ajw_metadata` |
| Inventory (재고) | localStorage + Supabase | `ajw_inventory` |
| SalesAnalysis | localStorage + Supabase | `ajw_sales` |
| SalesAggResult | localStorage + Supabase | `ajw_sales_agg` |
| OjcRules | localStorage + Supabase | `ajw_ojc_rules` |
| OjcProduct[] | localStorage + Supabase | `ajw_ojc_products` |
| ItemCode[] | localStorage + Supabase | `ajw_item_codes` |
| ImportOrder[] | localStorage 전용 | `ajw_import_orders` |
| ReconHistory | localStorage 전용 | `ajw_recon_history` |
| DetailedSalesRow[] | **IndexedDB** 전용 | `ajw_detailed_sales` |
| vendors | Supabase 전용 | `vendors` 테이블 |

**로드 패턴**: localStorage 우선 반환 → Supabase 연동 ON이면 원격 갱신 후 캐시 업데이트

---

## 파싱 모듈 (`lib/parse/`)

| 파일 | 입력 | 출력 | 비고 |
|---|---|---|---|
| `parseERP.ts` | ArrayBuffer | `ErpRow[]` | year=4자리, `detectYears()`로 연도 추출 |
| `parseStockFile.ts` | File | `Map<code, qty>` | EMP/이카운트 자동감지 |
| `parseDetailedSales.ts` | ArrayBuffer | `DetailedSalesRow[]` | year=2자리 YY |
| `parseCostFile.ts` | ArrayBuffer | `Map<code, CostEntry>` | 맥산/FLC 형식 분기 |
| `parseItemCodes.ts` | ArrayBuffer | `ItemCode[]` | 이카운트/EMP 품목 리스트 |

---

## 핵심 타입 위치

| 타입 | 위치 |
|---|---|
| `Metadata`, `Inventory`, `AppSettings` | `lib/types.ts` |
| `Step3Row`, `CodeCableEntry` | `lib/step3Core.ts` |
| `DetailedSalesRow` | `lib/parse/parseDetailedSales.ts` |
| `ImportOrder`, `ItemCode`, `OjcProduct` | `lib/supabase.ts` |
| `ReconHistorySummary` | `components/InventoryReconciliationTab.tsx` |

---

## OJC 분류 (`lib/ojcFilter.ts`)

```
classifyOjc(name)        → 일반 분류 (KT OJC / LG OJC / DROP / 피그테일 / ...)
classifyDetailed(name)   → 세분화 분류 (KT OJC-SP / KT OJC-DP / ...)
looksLikeOjc(name)       → OJC 의심 품목 감지 (미분류 경고용)
normalizeName(name)      → 대소문자·공백·언더스코어 정규화 후 prefix 매칭
```

---

## 상태 관리

- **전역 상태**: App.tsx에서 `useState`로 관리, props로 하위 컴포넌트에 전달
- **저장 패턴**: `setState(x)` 먼저 → `sbSave(x)` 비동기 후행
- **초기 로드**: App.tsx `useEffect`에서 `load*()` 함수 순차 호출
- **Supabase 연동 토글**: `getSyncEnabled()` / `setSyncEnabled()` — localStorage `ajw_supabase_sync`
