-- ============================================================
-- AJW 운영 도우미 웹 — Supabase 스키마
-- ============================================================
-- 새 Supabase 프로젝트 세팅 시 이 파일 전체를 SQL Editor에서 실행
-- ============================================================


-- ──────────────────────────────────────────────────────────
-- 1. app_data  (설정·메타데이터·재고·판매 캐시 — 키-값 스토어)
-- ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS app_data (
  id         text        PRIMARY KEY,
  data       jsonb       NOT NULL DEFAULT '{}',
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- 초기 레코드 (앱 최초 실행 전 미리 삽입)
INSERT INTO app_data (id, data) VALUES
  ('settings',     '{}'),
  ('metadata',     '{"cable":{},"housing":{},"ferrule":{}}'),
  ('inventory',    '{"cable":{},"housing":{},"ferrule":{}}'),
  ('sales',        '{}'),
  ('sales_agg',    'null'),
  ('ojc_products', '[]')
ON CONFLICT (id) DO NOTHING;

-- RLS: anon 읽기 허용 / authenticated 쓰기 허용
ALTER TABLE app_data ENABLE ROW LEVEL SECURITY;

CREATE POLICY "anon_read"  ON app_data FOR SELECT USING (true);
CREATE POLICY "auth_write" ON app_data FOR ALL    USING (auth.role() = 'authenticated');


-- ──────────────────────────────────────────────────────────
-- 2. recon_history  (재고 대사 이력 — 재고 관리 탭)
-- ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS recon_history (
  id         uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  date       date        UNIQUE NOT NULL,
  summary    jsonb       NOT NULL,  -- { total, match, diff, emp_only, ecount_only }
  rows       jsonb       NOT NULL,  -- ReconRow[] 배열
  created_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE recon_history ENABLE ROW LEVEL SECURITY;

CREATE POLICY "anon_read"  ON recon_history FOR SELECT USING (true);
CREATE POLICY "auth_write" ON recon_history FOR ALL    USING (auth.role() = 'authenticated');


-- ──────────────────────────────────────────────────────────
-- 3. vendors  (수입 업체 코드 — 수입 관리 탭, 2단계)
-- ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS vendors (
  code       text PRIMARY KEY,  -- ex: 'FLC'
  name       text NOT NULL      -- ex: 'FIBERCAN'
);

ALTER TABLE vendors ENABLE ROW LEVEL SECURITY;

CREATE POLICY "anon_read"  ON vendors FOR SELECT USING (true);
CREATE POLICY "auth_write" ON vendors FOR ALL    USING (auth.role() = 'authenticated');


-- ──────────────────────────────────────────────────────────
-- 4. import_orders  (수입 발주 현황 — 수입 관리 탭, 2단계)
-- ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS import_orders (
  id            uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  lot_code      text        NOT NULL,  -- ex: '14-K-107 (FLC-2601-01_01)'
  vendor_code   text        NOT NULL REFERENCES vendors(code),
  item_code     text        NOT NULL,
  order_date    date        NOT NULL,
  qty           integer     NOT NULL,
  expected_date date,
  status        text        NOT NULL DEFAULT 'pending',
  -- pending | partial | completed
  created_at    timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE import_orders ENABLE ROW LEVEL SECURITY;

CREATE POLICY "anon_read"  ON import_orders FOR SELECT USING (true);
CREATE POLICY "auth_write" ON import_orders FOR ALL    USING (auth.role() = 'authenticated');


-- ──────────────────────────────────────────────────────────
-- 5. import_receipts  (분할입고 기록 — 수입 관리 탭, 2단계)
-- ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS import_receipts (
  id           uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  order_id     uuid        NOT NULL REFERENCES import_orders(id) ON DELETE CASCADE,
  receipt_date date        NOT NULL,
  qty          integer     NOT NULL,
  created_at   timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE import_receipts ENABLE ROW LEVEL SECURITY;

CREATE POLICY "anon_read"  ON import_receipts FOR SELECT USING (true);
CREATE POLICY "auth_write" ON import_receipts FOR ALL    USING (auth.role() = 'authenticated');
