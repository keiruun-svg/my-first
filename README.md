# 📦 AJW 생산자재 발주계획 시스템

**(주)에이제이월드 SCM팀 — 생산자재 발주계획 자동화 웹앱**

ERP 구매현황 파일을 업로드하면 자재별 사용량 집계 → 판매 분석 → 연간 발주계획 Excel을 자동 생성합니다.  
Streamlit / Python 버전에서 **React + TypeScript + Vite** 로 전면 재구축되었습니다.

---

## 기술 스택

| 분류 | 기술 |
|------|------|
| 프론트엔드 | React 19, TypeScript 6, Vite 8 |
| 스타일 | Tailwind CSS 4 |
| Excel 생성 | ExcelJS 4 (서식·수식 포함 xlsx 생성) |
| Excel 파싱 | SheetJS (xlsx 0.18) |
| 백엔드 / DB | Supabase (PostgreSQL — 설정·메타데이터·재고 영속화) |

---

## 3-STEP 워크플로우

```
ERP 구매현황.xlsx
       │
       ▼
┌──────────────────┐
│  STEP 1          │  ERP 파싱 → 케이블 종류·파이·코어수·길이 분류
│  ERP 파일 가공   │  → 연도별 월별 사용량 집계
│                  │  출력: 가공파일.xlsx (집계 시트 + 원본 시트)
└────────┬─────────┘
         │ 가공파일.xlsx
         ▼
┌──────────────────┐
│  STEP 2          │  판매·생산 파일 업로드
│  판매 분석       │  → OJC 제품 필터링 → 타입별 CAGR 계산
│                  │  → 생산비중·수입의존도 분석
│                  │  출력: 판매분석.xlsx (7개 시트)
└────────┬─────────┘
         │ 가공파일.xlsx  +  판매분석.xlsx (선택)
         ▼
┌──────────────────┐
│  STEP 3          │  안전재고 계산 (월평균 × LT/30 × k)
│  발주계획 생성   │  CAGR 반영 예측량 → 필요발주량 산정
│                  │  STEP 2 연동 시 수요기반 분석 5컬럼 추가
│                  │  출력: 발주계획.xlsx (5개 시트)
└──────────────────┘
```

---

## 탭 구성

| 탭 | 기능 |
|----|------|
| 🏠 홈 | 대시보드 — 등록 자재 현황·재고 요약 |
| 📤 STEP 1 | ERP 파일 가공 → 가공파일.xlsx 생성 |
| 📈 STEP 2 | 판매·생산 파일 분석 → 판매분석.xlsx 생성 |
| 📊 STEP 3 | 가공파일 업로드 → 발주계획.xlsx 생성 |
| 🔍 판매 현황 분석 | 품목별·채널별 판매 추이 열람 |
| 🏷 품번 생성기 | OJC 품번 자동 생성 규칙 적용 |
| 📦 자재 관리 | 품번·품명·구매처·리드타임·현재고·기발주 등록 (관리자) |
| ⚙️ 파라미터 설정 | 안전재고 계수(k)·리드타임 기본값·색상 테마 |

---

## 파일 구조

```
react-app/src/
├── App.tsx                     # 탭 라우팅, Supabase 초기 로드
├── components/
│   ├── Step1.tsx               # STEP 1 UI
│   ├── Step2.tsx               # STEP 2 UI
│   ├── Step3.tsx               # STEP 3 UI (+ salesAgg 수요 분석 연동)
│   ├── Dashboard.tsx           # 홈 대시보드
│   ├── MaterialManager.tsx     # 자재 관리 (메타 + 재고)
│   ├── SalesAnalysisTab.tsx    # 판매 현황 분석
│   ├── Settings.tsx            # 파라미터 설정
│   ├── PartNumberGenerator.tsx # 품번 생성기
│   └── FileUploader.tsx        # 공용 파일 업로드 컴포넌트
│
└── lib/
    ├── types.ts                # 공용 타입 (Metadata, Inventory, AppSettings)
    ├── supabase.ts             # Supabase 클라이언트 + CRUD 헬퍼
    ├── step1Core.ts            # STEP 1 — ERP 파싱 및 집계 로직
    ├── step2Core.ts            # STEP 2 — 판매 분석 오케스트레이션
    ├── step3Core.ts            # STEP 3 — 발주계획 계산 (CodeCableEntry 포함)
    ├── ojcFilter.ts            # OJC 완제품 판별 (품명 패턴 매칭)
    ├── ojcAutoDetect.ts        # OJC 품번 자동 감지
    ├── download.ts             # xlsx 파일 다운로드 유틸
    ├── aggregate/
    │   ├── salesAgg.ts         # 타입별 판매·생산 집계, CAGR 계산
    │   └── pivot.ts            # ERP 데이터 연도별 피벗
    ├── parse/
    │   ├── parseERP.ts         # ERP 구매현황/구매조회 파싱
    │   ├── parseSales.ts       # 판매 파일 파싱
    │   ├── parseDetailedSales.ts
    │   └── classify.ts         # 케이블 종류·코어수 분류 (deriveKind/deriveCore)
    └── output/
        ├── writeGaong.ts       # 가공파일.xlsx 생성 (STEP 1 출력)
        ├── writeSalesAgg.ts    # 판매분석.xlsx 생성 (STEP 2 출력)
        └── writeStep3Excel.ts  # 발주계획.xlsx 생성 (STEP 3 출력)
```

---

## STEP 3 — 발주계획.xlsx 시트 구성

| 시트 | 내용 |
|------|------|
| 케이블 사용내역 | 케이블 타입별 연도별 연간·피크 사용량, 안전재고, 발주 필요량 |
| 하우징 사용내역 | 하우징 타입별 동일 구성 |
| 📦 품번별 발주 집계 | 하우징 품번별 합산 + 페롤(커넥터 타입별) 섹션 |
| 2026 월별 발주계획 | 과거 계절 패턴 기반 월별 자동 분배 (연간목표 입력 시 자동 계산) |
| ⚠ 이상항목 검토 | 사용량 급감·품번 미등록 등 주의 항목 |

**STEP 2 결과(판매분석.xlsx)가 로드된 경우** 케이블 사용내역 시트에 수요기반 분석 5컬럼이 추가됩니다:

| 컬럼 | 산식 |
|------|------|
| 수요기반 제안량 | Σ (판매실적 × (1+트렌드) × 생산비중 × 자재단가) |
| vs 3개년평균 | 제안량 / 3개년평균 − 1 |
| 생산비중 | 맥산 생산량 / 전체 판매량 (가중평균) |
| 판매트렌드 | 2개년 CAGR (없으면 1개년 증감률) |
| 수입의존위험도 | 🔴 고위험(<30%) / 🟠 주의(<70%) / 🟢 안전 |

---

## 안전재고 계산식

```
안전재고 = 월평균사용량 × (리드타임 / 30) × k

  월평균사용량 = 최근년 연간사용량 / 12
  k            = 안전재고 계수 (기본값 1.5, 설정 탭에서 변경)
  리드타임     = 자재별 등록값 (미등록 시 기본값 사용)
```

STEP 2 판매CAGR이 연동된 경우 케이블의 월평균사용량은 예측량(`latestAnnual × (1 + CAGR)`) 기준으로 계산됩니다.

---

## 개발 환경 실행

```bash
cd react-app
npm install
npm run dev
```

접속: http://localhost:5173

### 환경 변수

`react-app/.env.local` 파일 생성:

```env
VITE_SUPABASE_URL=https://your-project.supabase.co
VITE_SUPABASE_ANON_KEY=your-anon-key
VITE_ADMIN_PASS=자재관리탭_비밀번호
```

### 빌드

```bash
npm run build   # dist/ 폴더 생성
npm run preview # 빌드 결과 로컬 미리보기
```

---

## Supabase 데이터 구조

| 테이블 | 내용 |
|--------|------|
| `settings` | 안전재고 계수·리드타임 기본값·색상 테마 |
| `metadata` | 케이블·하우징·페롤별 품번·품명·구매처·리드타임 |
| `inventory` | 케이블·하우징·페롤별 현재고·기발주 |
| `sales_analysis` | STEP 2 판매 분석 결과 캐시 |
| `sales_agg` | 타입별 판매·생산 CAGR 집계 캐시 |
| `ojc_products` | OJC 품번 생성기용 제품 목록 |

---

## 주요 분류 로직

### 케이블 종류 (`classify.ts`)
- `deriveKind(name, spec, core)` — 품명/규격에서 a1 / b3 / om1 / om3 / drop / pigtail / a2 분류
- `deriveCore(name)` — 품명에서 코어수(1/2/4/8…) 추출
- A1 케이블은 색상(청/녹/적/자)별로 세분류

### OJC 판별 (`ojcFilter.ts`)
- 품명 패턴 매칭으로 광점퍼코드 완제품 구분
- 맥산 생산분 / 수입 완제품 구분 → 생산비중 계산

### 피그테일 집계 (`writeGaong.ts`)
- 0.9mm 피그테일은 코어수만큼 색상별 분리 집계
- 색상 매핑: 연청 → 청록, 연등 → 분홍

---

## 담당자

박정원 (SCM팀) — jw.park@ajw.co.kr
