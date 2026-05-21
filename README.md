# 📦 AJW 운영 도우미 웹

**(주)에이제이월드 SCM팀 — 운영 도우미 웹 개발**

| 단계 | 내용 | 상태 |
|------|------|------|
| 1단계 | 생산 자재 발주 계획 자동 생성 | ✅ 완료 |
| 2단계 | 수입 관리 + 재고 관리 고도화 | 🔲 진행 예정 |
| 3단계 | 완제품 수입 어시스턴트 | 🔲 진행 예정 |

---

## 빠른 시작

```bash
# 1. 의존성 설치
cd react-app
npm install

# 2. 환경변수 설정
cp .env.example .env.local
# .env.local 파일을 열어 실제 값 입력

# 3. Supabase DB 초기화
# Supabase 콘솔 > SQL Editor에서 supabase/schema.sql 전체 실행

# 4. 개발 서버 실행
npm run dev
# → http://localhost:5173
```

---

## 환경변수 (react-app/.env.local)

```env
VITE_SUPABASE_URL=https://your-project-id.supabase.co
VITE_SUPABASE_KEY=your-anon-key
VITE_ADMIN_PASS=자재관리탭_비밀번호
```

> `.env.example` 파일을 복사해서 사용. 실제 키는 Supabase 콘솔 **Project Settings > API** 에서 확인.

---

## 기술 스택

| 분류 | 기술 |
|------|------|
| 프론트엔드 | React 19, TypeScript 6, Vite 8 |
| 스타일 | Tailwind CSS 4 |
| Excel 생성 | ExcelJS 4 |
| Excel 파싱 | SheetJS (xlsx 0.18) |
| 백엔드 / DB | Supabase (PostgreSQL) + localStorage fallback |

---

## 탭 구성

| 탭 | 기능 |
|----|------|
| 🏠 홈 | 대시보드 — 등록 자재 현황·재고 요약 |
| 📤 STEP 1 | ERP 파일 가공 → 가공파일.xlsx 생성 |
| 📈 STEP 2 | 판매·생산 파일 분석 → 판매분석.xlsx 생성 |
| 📊 STEP 3 | 발주계획 생성 → 발주계획.xlsx 생성 |
| 🔍 판매 현황 분석 | 품목별·채널별 판매 추이 열람 |
| 🏷 품번 생성기 | OJC 광점퍼코드 품번 자동 생성 |
| 📦 자재 관리 | 품번·품명·구매처·리드타임·현재고·기발주 등록 (관리자 전용) |
| 🗂 재고 관리 | EMP ↔ 이카운트 재고 대사 (4파일 업로드) |
| 🚢 수입 관리 | EMP 코드 생성 + 발주·입고 관리 (2단계 구현 예정) |
| ⚙️ 파라미터 설정 | 안전재고 계수(k)·리드타임 기본값·색상 테마 |

---

## 3-STEP 워크플로우

```
ERP 구매현황.xlsx
       │
       ▼
┌──────────────────┐
│  STEP 1          │  ERP 파싱 → 케이블 종류·파이·코어수·길이 분류
│  ERP 파일 가공   │  → 연도별 월별 사용량 집계
│                  │  출력: 가공파일.xlsx (집계 + 코드매핑 + 요약 시트)
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  STEP 2          │  판매·생산 파일 업로드
│  판매 분석       │  → OJC 제품 필터링 → 타입별 CAGR 계산
│                  │  출력: 판매분석.xlsx
└────────┬─────────┘
         │ 가공파일.xlsx + 판매분석.xlsx (선택)
         ▼
┌──────────────────┐
│  STEP 3          │  안전재고 = 3개년 피크평균 × (LT/30) × k
│  발주계획 생성   │  발주필요량 = max(0, 예측연간 + 안전재고 - 현재고 - 기발주)
│                  │  출력: 발주계획.xlsx (4개 시트)
└──────────────────┘
```

---

## 안전재고 계산식

```
안전재고 = 3개년 피크평균 × (리드타임 / 30) × k

  3개년 피크평균 = 각 연도 월별 최대값의 평균
  k            = 안전재고 계수 (기본값 1.5, 설정 탭에서 변경)
  리드타임     = 자재별 등록값 (미등록 시 기본값 사용, 단위: 일)
```

---

## 발주계획.xlsx 시트 구성

| 시트 | 내용 |
|------|------|
| 케이블 사용내역 | 케이블 타입별 연도별 연간·피크 사용량, 안전재고, 발주 필요량 |
| 하우징 사용내역 | 하우징 타입별 동일 구성 |
| 📦 품번별 발주 집계 | 품번별 합산 (케이블·하우징·페롤 섹션) + k 입력 셀 |
| 2026 월별 발주계획 | 과거 계절 패턴 기반 월별 자동 분배 |

---

## Supabase 스키마

`supabase/schema.sql` 참고. 주요 테이블:

| 테이블 | 용도 | 단계 |
|--------|------|------|
| `app_data` | 설정·메타·재고·판매 캐시 (키-값) | 1단계 ✅ |
| `recon_history` | 재고 대사 이력 | 2단계 |
| `vendors` | 수입 업체 코드 | 2단계 |
| `import_orders` | 수입 발주 현황 | 2단계 |
| `import_receipts` | 분할입고 기록 | 2단계 |

> `app_data` 테이블의 `id` 값: `settings` / `metadata` / `inventory` / `sales` / `sales_agg` / `ojc_products`

**쓰기 권한**: `CAN_WRITE = import.meta.env.DEV` — 개발 모드에서만 Supabase 쓰기 가능, 프로덕션은 읽기 전용.

---

## 파일 구조

```
react-app/src/
├── App.tsx
├── components/
│   ├── Step1.tsx / Step2.tsx / Step3.tsx
│   ├── Dashboard.tsx
│   ├── SalesAnalysisTab.tsx
│   ├── PartNumberGenerator.tsx
│   ├── MaterialManager.tsx
│   ├── InventoryReconciliationTab.tsx  # 재고 대사
│   ├── ImportTab.tsx                   # 수입 관리 (2단계)
│   ├── Settings.tsx
│   └── FileUploader.tsx
└── lib/
    ├── types.ts / supabase.ts
    ├── step1Core.ts / step2Core.ts / step3Core.ts
    ├── ojcFilter.ts / ojcAutoDetect.ts / download.ts
    ├── aggregate/
    │   ├── salesAgg.ts   # CAGR 계산
    │   └── pivot.ts
    ├── parse/
    │   ├── parseERP.ts / parseSales.ts / parseDetailedSales.ts
    │   └── classify.ts   # deriveKind / derivePai / deriveCore / deriveLength
    └── output/
        ├── writeGaong.ts       # 가공파일.xlsx
        ├── writeSalesAgg.ts    # 판매분석.xlsx
        └── writeStep3Excel.ts  # 발주계획.xlsx

supabase/
└── schema.sql   # DB 초기화 SQL (새 환경 세팅 시 실행)

CLAUDE.md        # Claude Code 세션용 프로젝트 맥락 문서
```

---

## 개발 명령어

```bash
cd react-app
npm run dev      # 개발 서버
npm run build    # 프로덕션 빌드
npm run lint     # ESLint
npx tsc --noEmit # 타입 체크
```

---

## 담당자

박정원 (SCM팀) — jw.park@ajw.co.kr
