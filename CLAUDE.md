# AJW SCM 어시스턴트 — CLAUDE.md

(주)에이제이월드 SCM팀 운영 도우미 웹 — React 19 / TypeScript / Vite / Tailwind CSS 4 / ExcelJS / SheetJS / Supabase

```bash
# react-app/
npm run dev      # localhost:5173
npm run build
npx tsc --noEmit
```

---

## 탭 구조 (App.tsx)

**NAV_TABS**: home / steps / sales / recon / import / partnum / admin  
**STEP_TABS** (steps 하위): step1 / step2 / step3  
**ADMIN_TABS** (admin 하위): admin_material / admin_ojcrules / admin_settings

```typescript
type StepId     = 'step1' | 'step2' | 'step3'
type AdminTabId = 'admin_material' | 'admin_ojcrules' | 'admin_settings'
type TabId      = Exclude<typeof NAV_TABS[number]['id'], 'steps' | 'admin'> | StepId | AdminTabId
function isStepId(id: string): id is StepId
function isAdminTabId(id: string): id is AdminTabId
```

admin 탭: `VITE_ADMIN_PASS` 비밀번호 게이트 (session-level, 페이지 리로드 시 초기화)

---

## 핵심 파일

| 파일 | 역할 |
|------|------|
| `components/Dashboard.tsx` | 홈 — KPI 6개 + 수입 D-day + 재고대사 + OJC CAGR |
| `components/ImportTab.tsx` | 수입 관리 — 서브탭: 발주계획 / 간헐적수요 / 발주현황 / 수익성 분석 |
| `components/ImportOrdersView.tsx` | 발주 현황 (업체+차수 단위 등록·D-day·입고처리) |
| `components/ProfitabilityAnalysis.tsx` | 수익성 분석 — 원가 파일 업로드 → 마진율 + 권고 테이블 + Excel |
| `lib/parse/parseCostFile.ts` | 맥산 생산원가 / FLC 수입원가 / 계약 단가 파싱 → `Map<code, CostEntry\|ContractItem>` |
| `components/InventoryReconciliationTab.tsx` | 재고 대사 (EMP ↔ 이카운트) + 이력 누적 차트 |
| `components/PartNumberGenerator.tsx` | OJC 품번(자동/수동) + EMP 코드 생성기 서브탭 |
| `components/OjcRulesEditor.tsx` | OJC 코드표 편집 (관리자 탭 하위, 독립 컴포넌트) |
| `components/MaterialManager.tsx` | 자재 관리 (케이블/하우징/페룰 메타+재고) |
| `lib/supabase.ts` | Supabase 클라이언트 + localStorage CRUD 전체 |
| `lib/ojcFilter.ts` | `classifyOjc` / `looksLikeOjc` — normalizeName 후 prefix 매칭 |
| `lib/parse/parseStockFile.ts` | 이카운트 수불부 / EMP 재고현황 자동감지 → `Map<code, qty>` |
| `lib/parse/parseItemCodes.ts` | 이카운트/EMP 품목 리스트 파싱 → `ItemCode[]` |

---

## 3-STEP 데이터 흐름 (생산자재 발주계획)

```
STEP 1: ERP 파일 → parseERP → classify → 가공파일.xlsx + metadata 저장
STEP 2: 판매파일 → parseSales → CAGR → 판매분석.xlsx + SalesAggResult 저장
STEP 3: 가공파일 + 판매분석 + metadata + inventory
        → 안전재고 = 3개년 피크평균 × (LT/30) × safetyK
        → 발주필요량 = max(0, 예측연간 + 안전재고 − 현재고 − 기발주)
        → 발주계획.xlsx
```

---

## 수입 관리 탭 (ImportTab)

**발주계획**: 재고파일(parseStockFile 자동감지) + 판매현황 → OJC 필터 → 월간최고 기반 커버리지  
- 커버리지 = 현재고 ÷ 월간최고 / 발주필요량 = (월간최고 × 목표개월) − 현재고  
- 간헐적 수요 필터: 판매월수 < 기준(기본 3개월) → 발주 제외, 별도 탭으로 분리

**발주 현황** (`ImportOrdersView`): 업체+연도+차수 단위 (예: `FLC 26-05차`)  
- 차수 자동 제안, D-day 카운트, 입고처리 → 리드타임 자동 계산  
- `localStorage: ajw_import_orders`

**수익성 분석** (`ProfitabilityAnalysis`): 맥산(국내) + FLC(수입) 원가 + 계약 단가 파일 업로드 → 마진율·권고  
- 원가 적용: KT향 → 표준원가, LG향·기타 → 생산원가
- 판매가: 계약 단가 우선 → 없으면 판매 이력 평균 (DetailedSalesRow)
- `마진율 = (판매가 − 원가) / 판매가 × 100`
- `생산유지기준 = max(20, minMarginPct × 3)` (검토필요 구간 자동 확장)
- `importThreshold = thresholdEnabled ? 100 + bufferPct : 100` (맥산 생산 가중치 ON/OFF)

---

## Supabase / localStorage

`CAN_WRITE = true` — 관리자 접근은 비밀번호(`VITE_ADMIN_PASS`)로 제어  
연동 토글: `getSyncEnabled()` / `setSyncEnabled()`  
`.env`: `VITE_SUPABASE_URL`, `VITE_SUPABASE_KEY`, `VITE_ADMIN_PASS`

**app_data 테이블 키**: settings / metadata / inventory / sales / sales_agg / ojc_products / item_codes  
**vendors 테이블**: `{ code, name }` — 수입 업체코드  
**localStorage 전용**: `ajw_detailed_sales` (DetailedSalesRow[]) / `ajw_import_orders` (ImportOrder[]) / `ajw_recon_history` (ReconHistorySummary[]) / `ajw_recon_rows_${date}` (대사 상세)

---

## 핵심 타입

`Metadata` / `Inventory` → `lib/types.ts`  
`Step3Row` → `lib/step3Core.ts`  
`ItemCode { code, name, spec }` → `lib/supabase.ts` + `lib/parse/parseItemCodes.ts`  
`ImportOrder { id, vendorCode, vendorName, year, seq, orderDate, expectedDate, actualDate? }` → `lib/supabase.ts`  
`DetailedSalesRow { customer, code, name, year, month, qty, price, isForeign }` → `lib/parse/parseDetailedSales.ts`

---

## OJC / EMP 도메인

**OJC 품번**: `{타입}{코어타입}{코어수}{재질}{커넥터}{마킹}-{길이코드}` 예: `C-ALNN-3`  
**EMP 로트코드**: `{기본코드} ({업체코드}-{YYMM}-{발주차수}_{입고회차})` 예: `14-K-107 (FLC-2601-01_02)`  
**classifyOjc(name)**: `lib/ojcFilter.ts` — `normalizeName`(대소문자·공백·언더스코어 정규화) 후 prefix 매칭  
**looksLikeOjc(name)**: 분류 실패했지만 OJC 의심 품목 감지 → ImportTab에서 경고 표시  
**EMP 로케이션**: `lib/assignLocation.ts` — KT향 재구성 예정, 탭 숨김 상태

---

## 재고 대사 공식

`EMP조정 = EMP재고 − 출고합산` / `EC조정 = EC재고 + 미출하` / `차이 = EMP조정 − EC조정`

입력 4종: EMP 재고현황(.xls) / 이카운트 수불부(.xlsx 복수) / 미출하현황(.xlsx) / 일일입출고확인서(.xlsx)

---

## 코딩 컨벤션

- 한국어 변수명: `품번`, `품명`, `현재고`, `기발주`, `안전재고`, `리드타임`
- Excel 파일명: `기능명_${today()}.xlsx` — `today()` = `YYYYMMDD_HHmm`
- Supabase 저장: `setState(x)` 먼저, `sbSave(x)` 비동기 후행

---

## 구현 상태 (2026-05-28)

| 탭 | 상태 |
|---|---|
| 홈 대시보드 | ✅ 완료 (KPI 6개 / 수입 D-day / 재고대사 현황 / OJC CAGR) |
| STEP 1/2/3 | 🔄 피드백 수정 중 |
| 판매 현황 분석 | 🔄 피드백 수정 중 |
| 재고 대사 | 🔄 피드백 수정 중 (이력 누적 + 추이 차트 완료) |
| 수입 관리 | ✅ 완료 (발주계획/간헐적수요/발주현황/수익성 분석) |
| 품번 생성기 | 🔄 피드백 수정 중 (EMP 로케이션 탭 숨김 — KT향 재구성 예정) |
| 관리자 | ✅ 완료 (자재관리 / OJC 코드표 / 설정) |

**미구현**: EMP 로케이션 KT향 / 리드타임 집계 분석

---

## 문서 업데이트 방법

작업 후 **"문서 업데이트해줘"** 또는 **"메모리 최신화해줘"** 라고 말하면 됩니다.  
타입 변경 시 반드시 소스 파일을 직접 읽어서 복사 — 추론 금지.
