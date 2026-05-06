#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJW 생산자재 발주계획 웹앱
실행: streamlit run web_app.py --server.address=0.0.0.0 --server.port=8501
"""
import streamlit as st
import json, tempfile, os, pandas as pd
from pathlib import Path
from datetime import datetime

SETTINGS_FILE = Path(__file__).parent / "settings.json"
METADATA_FILE = Path(__file__).parent / "metadata.json"

DEFAULT_SETTINGS = {
    "lead_time_default": 60,
    "colors": {
        "main_header": "1F3864",
        "year_23": "2F5597",
        "year_24": "2E75B6",
        "year_25": "155480",
    }
}

def load_settings() -> dict:
    if SETTINGS_FILE.exists():
        try:
            return json.loads(SETTINGS_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return DEFAULT_SETTINGS.copy()

def save_settings(s: dict):
    SETTINGS_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding='utf-8')

def load_metadata() -> dict:
    if METADATA_FILE.exists():
        try:
            return json.loads(METADATA_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return {"cable": {}, "housing": {}}

def save_metadata(m: dict):
    METADATA_FILE.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding='utf-8')

def meta_to_step1(metadata: dict):
    """metadata.json → step1_core용 cable_meta / housing_meta_in 딕셔너리"""
    cable_meta = {}
    for k, v in metadata.get('cable', {}).items():
        parts = k.split('|', 1)
        if len(parts) == 2:
            cable_meta[(parts[0], parts[1])] = v
    housing_meta_in = {}
    for k, v in metadata.get('housing', {}).items():
        parts = k.split('|', 1)
        if len(parts) == 2:
            housing_meta_in[(parts[0], parts[1])] = v if isinstance(v, list) else [v]
    return cable_meta, housing_meta_in

def save_upload(uploaded_file, tmp_dir: str, filename: str) -> str:
    path = os.path.join(tmp_dir, filename)
    with open(path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    return path

def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

# ── 페이지 설정 ──────────────────────────────────────────────
st.set_page_config(
    page_title="AJW 발주계획 시스템",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .main-title { font-size: 1.6rem; font-weight: 700; color: #1F3864; margin-bottom: 0.2rem; }
    .sub-title  { font-size: 0.95rem; color: #666; margin-bottom: 1.5rem; }
    .step-box   { background: #f0f4fa; border-left: 4px solid #2E75B6;
                  padding: 1rem 1.2rem; border-radius: 6px; margin-bottom: 1rem; }
    .log-box    { background: #1e1e1e; color: #d4d4d4; font-family: monospace;
                  font-size: 0.82rem; padding: 0.8rem; border-radius: 6px;
                  white-space: pre-wrap; max-height: 200px; overflow-y: auto; }
    .warn       { color: #e67e22; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📦 AJW 생산자재 발주계획 시스템</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">(주)에이제이월드 SCM팀 — 로우데이터 업로드 후 버튼 클릭으로 Excel 자동 생성</div>', unsafe_allow_html=True)

settings = load_settings()

tab1, tab2, tab3, tab4 = st.tabs([
    "📤 STEP 1 — ERP 파일 가공",
    "📊 STEP 2 — 발주계획 생성",
    "📋 품번 관리",
    "⚙️ 파라미터 & 양식 설정",
])

# ═══════════════════════════════════════════════════════════
# STEP 1
# ═══════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="step-box"><b>STEP 1 — ERP 파일 가공</b>: 맥산 ERP에서 추출한 <b>구매조회</b> 또는 <b>구매현황</b> 파일을 업로드하면 생산자재_사용내역.xlsx를 자동 생성합니다.<br>두 형식 모두 지원하며 자동으로 감지합니다. 품번·품명·구매처·리드타임은 <b>📋 품번 관리</b> 탭 정보로 자동 채워집니다.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown("**필수 파일**")
        row_file = st.file_uploader(
            "구매조회 / 구매현황 파일 (ERP 원본)",
            type=['xlsx'], key='s1_row',
            help="맥산 ERP → 구매조회 또는 구매현황 시트가 포함된 파일 업로드 (자동 감지)"
        )
    with col2:
        st.markdown("**선택 파일 (A1-4C / OM4-DP 보완용)**")
        ojc_file = st.file_uploader(
            "비밀_OJC 참고파일.xlsx",
            type=['xlsx'], key='s1_ojc',
            help="비밀_OJC 3개년 생산자제 사용 내역_*.xlsx — A1-4C, OM4-DP 누락 항목 보완"
        )
        meta_info = load_metadata()
        n_cable_meta = len(meta_info.get('cable', {}))
        n_housing_meta = len(meta_info.get('housing', {}))
        st.info(
            f"현재 저장된 품번 — 케이블 **{n_cable_meta}** 타입 / 하우징 **{n_housing_meta}** 타입\n\n"
            f"리드타임 기본값: **{settings.get('lead_time_default', 60)}일** (⚙️ 탭에서 변경)"
        )

    st.divider()

    if st.button("▶ STEP 1 실행 — ERP 파일 가공 & 사용내역 생성", type="primary", disabled=(row_file is None)):
        if row_file is None:
            st.error("구매조회 / 구매현황 파일을 업로드해주세요.")
        else:
            prog_bar  = st.progress(0)
            prog_text = st.empty()

            def _on_progress(pct: int, msg: str):
                prog_bar.progress(min(pct, 100))
                prog_text.markdown(
                    f'<div style="font-size:0.85rem;color:#555;margin-top:4px">⏳ {msg}</div>',
                    unsafe_allow_html=True,
                )

            try:
                import importlib, step1_core
                importlib.reload(step1_core)
                metadata = load_metadata()
                cable_meta, housing_meta_in = meta_to_step1(metadata)
                with tempfile.TemporaryDirectory() as tmp:
                    row_path = save_upload(row_file, tmp, "가공파일_통합_v양식.xlsx")
                    ojc_path = save_upload(ojc_file, tmp, "ojc_ref.xlsx") if ojc_file else None
                    result_bytes, logs, cable_keys, housing_keys, years = step1_core.run(
                        row_path, cable_meta, housing_meta_in, ojc_path, settings,
                        progress_cb=_on_progress)

                # 신규 타입 metadata 자동 추가
                new_types = 0
                for (pai, ct) in cable_keys:
                    k = f"{pai}|{ct}"
                    if k not in metadata.get('cable', {}):
                        metadata.setdefault('cable', {})[k] = {
                            '품번': '', '품명': '', '구매처': '', '리드타임': None}
                        new_types += 1
                for (pai, htype) in housing_keys:
                    k = f"{pai}|{htype}"
                    if k not in metadata.get('housing', {}):
                        metadata.setdefault('housing', {})[k] = [
                            {'품번': '', '품명': '', '구매처': '', '리드타임': None}]
                        new_types += 1
                if new_types:
                    save_metadata(metadata)
                    logs.append(f"ℹ️ 📋 품번 관리 탭에 신규 타입 {new_types}개 추가 — 품번을 입력해주세요.")

                # 완료 표기
                prog_bar.progress(100)
                prog_text.markdown(
                    '<div style="font-size:0.9rem;font-weight:600;color:#1a7a3c;'
                    'background:#e8f5e9;padding:6px 12px;border-radius:6px;'
                    'border-left:4px solid #1a7a3c;margin-top:6px">'
                    '✅ 처리 완료!</div>',
                    unsafe_allow_html=True,
                )

                n_yr = len(years)
                fname = f"{n_yr}개년_생산자재_사용내역_{ts()}.xlsx"
                st.download_button(
                    label="⬇ 사용내역.xlsx 다운로드",
                    data=result_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                log_text = "\n".join(logs)
                st.markdown(f'<div class="log-box">{log_text}</div>', unsafe_allow_html=True)

                warns = [l for l in logs if '⚠' in l]
                if warns:
                    st.markdown(f'<span class="warn">⚠ 주의: {warns[-1]}</span>', unsafe_allow_html=True)
                    st.markdown("품번이 없는 항목은 📋 **품번 관리** 탭에서 입력 후 다시 실행하세요.")
            except Exception as e:
                prog_bar.progress(0)
                prog_text.empty()
                st.error(f"오류 발생: {e}")

# ═══════════════════════════════════════════════════════════
# STEP 2
# ═══════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="step-box"><b>STEP 2 — 발주계획 생성</b>: STEP 1 결과(현재고·기발주 입력 완료)와 ERP 원본 파일을 업로드하면 2026_연간발주계획.xlsx를 생성합니다.<br><b>구매조회</b> 또는 <b>구매현황</b> 파일을 그대로 업로드하면 자동으로 변환됩니다.</div>', unsafe_allow_html=True)

    col3, col4 = st.columns([1, 1])
    with col3:
        st.markdown("**필수 파일**")
        row_file2 = st.file_uploader(
            "구매조회 / 구매현황 파일 (ERP 원본)",
            type=['xlsx'], key='s2_row',
            help="맥산 ERP → 구매조회 또는 구매현황 시트가 포함된 파일 (자동 감지 및 변환)"
        )
        usage_file = st.file_uploader(
            "3개년_사용내역.xlsx (현재고·기발주 입력 완료)",
            type=['xlsx'], key='s2_usage',
            help="STEP 1 결과 파일에 현재고·기발주를 직접 입력한 버전"
        )
    with col4:
        st.markdown("**선택 파일**")
        ojc_file2 = st.file_uploader(
            "비밀_OJC 참고파일.xlsx",
            type=['xlsx'], key='s2_ojc',
            help="0.9mm 피그테일 케이블 데이터 포함"
        )
        cmp_file = st.file_uploader(
            "생산_판매_비교.xlsx",
            type=['xlsx'], key='s2_cmp',
            help="있으면 수요 기반 분석(제안량·트렌드·위험도) 시트가 자동 추가됩니다."
        )

    st.divider()

    ready2 = (row_file2 is not None) and (usage_file is not None)
    if st.button("▶ STEP 2 실행 — 발주계획 생성", type="primary", disabled=(not ready2)):
        if not ready2:
            st.error("가공파일과 사용내역 파일을 모두 업로드해주세요.")
        else:
            with st.spinner("발주계획 생성 중..."):
                try:
                    import step2_core
                    with tempfile.TemporaryDirectory() as tmp:
                        row_path2   = save_upload(row_file2,  tmp, "가공파일_통합_v양식.xlsx")
                        usage_path  = save_upload(usage_file, tmp, "사용내역.xlsx")
                        ojc_path2   = save_upload(ojc_file2,  tmp, "ojc_ref.xlsx")  if ojc_file2 else None
                        cmp_path    = save_upload(cmp_file,   tmp, "생산_판매_비교.xlsx") if cmp_file  else None
                        result_bytes2, logs2 = step2_core.run(row_path2, usage_path, ojc_path2, cmp_path, settings)

                    st.success("✅ 발주계획 생성 완료!")
                    fname2 = f"2026_연간발주계획_{ts()}.xlsx"
                    st.download_button(
                        label="⬇ 2026_연간발주계획.xlsx 다운로드",
                        data=result_bytes2,
                        file_name=fname2,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    log_text2 = "\n".join(logs2)
                    st.markdown(f'<div class="log-box">{log_text2}</div>', unsafe_allow_html=True)
                    st.markdown("💡 노란색 셀(2026 목표 발주량)에 목표량을 입력하면 필요 발주량이 자동 계산됩니다.")
                except Exception as e:
                    st.error(f"오류 발생: {e}")

# ═══════════════════════════════════════════════════════════
# 품번 관리
# ═══════════════════════════════════════════════════════════
with tab3:
    st.markdown("### 📋 품번 관리")
    st.info(
        "STEP 1 실행 후 발견된 신규 타입이 자동으로 추가됩니다. "
        "품번·품명·구매처·리드타임을 직접 입력하고 **저장** 버튼을 누르세요. "
        "다음 STEP 1 실행 시 자동으로 적용됩니다."
    )

    metadata = load_metadata()

    # ── 케이블 ──────────────────────────────────────────────
    st.markdown("#### 케이블")
    with st.expander("입력 예시 보기"):
        st.dataframe(pd.DataFrame([{
            '파이': '2.0mm', '케이블종류': 'B3-SP',
            '품번': 'N90-29-1100', '품명': 'OPTICAL FIBER CABLE, B3, SP, 2.0MM',
            '구매처': 'CCTC', '리드타임': 40,
        }]), hide_index=True, use_container_width=True)
    cable_rows = []
    for k, v in sorted(metadata.get('cable', {}).items()):
        pai, ct = k.split('|', 1) if '|' in k else (k, '')
        cable_rows.append({
            '파이': pai,
            '케이블종류': ct,
            '품번': v.get('품번') or '',
            '품명': v.get('품명') or '',
            '구매처': v.get('구매처') or '',
            '리드타임': int(v['리드타임']) if v.get('리드타임') else None,
            '상태': '✓' if v.get('품번') else '⚠ 미입력',
        })

    if not cable_rows:
        df_cable = pd.DataFrame(columns=['파이','케이블종류','품번','품명','구매처','리드타임'])
    else:
        df_cable = pd.DataFrame([{k: v for k, v in r.items() if k != '상태'} for r in cable_rows])

    edited_cable = st.data_editor(
        df_cable,
        key='cable_editor',
        column_config={
            '파이':      st.column_config.SelectboxColumn('파이', options=['2.0mm','3.0mm','0.9mm'], width='small', required=True),
            '케이블종류': st.column_config.TextColumn('케이블종류', width='medium'),
            '품번':      st.column_config.TextColumn('품번', width='medium'),
            '품명':      st.column_config.TextColumn('품명', width='large'),
            '구매처':    st.column_config.TextColumn('구매처', width='medium'),
            '리드타임':  st.column_config.NumberColumn('리드타임', min_value=1, max_value=365, step=1, format='%d일'),
        },
        num_rows='dynamic',
        hide_index=True,
        use_container_width=True,
    )

    # ── 하우징 ──────────────────────────────────────────────
    st.markdown("#### 하우징")
    st.caption("하나의 하우징 타입에 여러 부품(하우징 본체·스프링·더스트캡 등)이 있으면 행을 추가하세요.")
    with st.expander("입력 예시 보기"):
        st.dataframe(pd.DataFrame([
            {'파이': '2.0mm', '하우징타입': 'LC/PC 청색', '품번': 'N90-29-1213',  '품명': 'FERRULE (W/ FLANGE,LC/PC TYPE)',          '구매처': 'CCTC',     '리드타임': 40},
            {'파이': '2.0mm', '하우징타입': 'LC/PC 청색', '품번': 'P14-LS-4228',  '품명': 'OJC HOUSING KIT (OJC LC/PC 2.0 BLUE)',    '구매처': 'FIBERCAN', '리드타임': 60},
            {'파이': '2.0mm', '하우징타입': 'LC/PC 청색', '품번': 'N94-1-04075',  '품명': 'SPRING',                                   '구매처': '상영스프링','리드타임': 30},
            {'파이': '2.0mm', '하우징타입': 'SC/PC 청색', '품번': 'N90-29-1315',  '품명': 'FERRULE (W/ FLANGE,SC/PC TYPE)',           '구매처': 'CCTC',     '리드타임': 40},
            {'파이': '2.0mm', '하우징타입': 'SC/APC 녹색','품번': 'N90-29-1323',  '품명': 'FERRULE (W/ FLANGE,SC/APC TYPE)',          '구매처': 'CCTC',     '리드타임': 40},
            {'파이': '2.0mm', '하우징타입': 'LC/PC 청색', '품번': 'P15-RM-4021',  '품명': 'DUST CAP(PC)',                             '구매처': 'FIBERCAN', '리드타임': 60},
        ]), hide_index=True, use_container_width=True)
        st.caption("같은 하우징 타입에 부품이 여러 개면 파이·하우징타입을 동일하게 입력하고 행을 추가합니다.")
    housing_rows = []
    for k, v in sorted(metadata.get('housing', {}).items()):
        pai, htype = k.split('|', 1) if '|' in k else (k, '')
        items = v if isinstance(v, list) else [v]
        for item in items:
            housing_rows.append({
                '파이': pai,
                '하우징타입': htype,
                '품번': item.get('품번') or '',
                '품명': item.get('품명') or '',
                '구매처': item.get('구매처') or '',
                '리드타임': int(item['리드타임']) if item.get('리드타임') else None,
            })

    if housing_rows:
        df_housing = pd.DataFrame(housing_rows)
    else:
        df_housing = pd.DataFrame(columns=['파이','하우징타입','품번','품명','구매처','리드타임'])

    edited_housing = st.data_editor(
        df_housing,
        key='housing_editor',
        column_config={
            '파이':      st.column_config.SelectboxColumn('파이', options=['2.0mm','3.0mm','0.9mm'], width='small', required=True),
            '하우징타입': st.column_config.TextColumn('하우징타입', width='medium'),
            '품번':      st.column_config.TextColumn('품번', width='medium'),
            '품명':      st.column_config.TextColumn('품명', width='large'),
            '구매처':    st.column_config.TextColumn('구매처', width='medium'),
            '리드타임':  st.column_config.NumberColumn('리드타임', min_value=1, max_value=365, step=1, format='%d일'),
        },
        num_rows='dynamic',
        hide_index=True,
        use_container_width=True,
    )

    st.divider()
    if st.button("💾 품번 저장", type="primary"):
        new_meta = load_metadata()

        if edited_cable is not None and not edited_cable.empty:
            cable_dict = {}
            for _, row in edited_cable.iterrows():
                if not row.get('파이') or not row.get('케이블종류'): continue
                k = f"{row['파이']}|{row['케이블종류']}"
                cable_dict[k] = {
                    '품번':   str(row['품번']).strip()  if pd.notna(row['품번'])   else '',
                    '품명':   str(row['품명']).strip()  if pd.notna(row['품명'])   else '',
                    '구매처': str(row['구매처']).strip() if pd.notna(row['구매처']) else '',
                    '리드타임': int(row['리드타임']) if pd.notna(row['리드타임']) else None,
                }
            new_meta['cable'] = cable_dict

        if edited_housing is not None and not edited_housing.empty:
            housing_dict = {}
            for _, row in edited_housing.iterrows():
                if not row.get('파이') or not row.get('하우징타입'): continue
                k = f"{row['파이']}|{row['하우징타입']}"
                housing_dict.setdefault(k, []).append({
                    '품번':   str(row['품번']).strip()  if pd.notna(row['품번'])   else '',
                    '품명':   str(row['품명']).strip()  if pd.notna(row['품명'])   else '',
                    '구매처': str(row['구매처']).strip() if pd.notna(row['구매처']) else '',
                    '리드타임': int(row['리드타임']) if pd.notna(row['리드타임']) else None,
                })
            new_meta['housing'] = housing_dict

        save_metadata(new_meta)
        st.success("✅ 품번 정보가 저장됐습니다. 다음 STEP 1 실행 시 자동 적용됩니다.")
        st.rerun()

# ═══════════════════════════════════════════════════════════
# 설정
# ═══════════════════════════════════════════════════════════
with tab4:
    st.markdown("### ⚙️ 파라미터 설정")
    st.markdown("변경 후 **저장** 버튼을 누르면 다음 실행부터 적용됩니다.")
    st.divider()

    col5, col6 = st.columns([1, 1])
    with col5:
        st.markdown("**계산 파라미터**")
        lt_val = st.number_input(
            "리드타임 기본값 (일) — 리드타임 미입력 항목에 적용",
            min_value=1, max_value=365,
            value=int(settings.get('lead_time_default', 60)),
            step=1,
        )
        st.markdown("**안전재고 계산 공식**")
        st.code("안전재고 = 3개년 피크 사용량 평균 × (리드타임 ÷ 30)", language=None)

    with col6:
        st.markdown("**Excel 헤더 색상** (hex 6자리, # 제외)")
        colors = settings.get('colors', DEFAULT_SETTINGS['colors'])
        c_main = st.text_input("메인 헤더 색상", value=colors.get('main_header', '1F3864'), max_chars=6)
        c_23   = st.text_input("2023년 컬럼 색상", value=colors.get('year_23', '2F5597'), max_chars=6)
        c_24   = st.text_input("2024년 컬럼 색상", value=colors.get('year_24', '2E75B6'), max_chars=6)
        c_25   = st.text_input("2025년 컬럼 색상", value=colors.get('year_25', '155480'), max_chars=6)

        st.markdown("미리보기")
        cols_prev = st.columns(4)
        for col_p, label, hex_val in zip(
            cols_prev,
            ['메인', '23년', '24년', '25년'],
            [c_main, c_23, c_24, c_25]
        ):
            col_p.markdown(
                f'<div style="background:#{hex_val};color:#fff;text-align:center;'
                f'padding:6px 2px;border-radius:4px;font-size:0.8rem">{label}</div>',
                unsafe_allow_html=True,
            )

    st.divider()
    if st.button("💾 설정 저장", type="primary"):
        new_settings = {
            "lead_time_default": lt_val,
            "colors": {
                "main_header": c_main,
                "year_23": c_23,
                "year_24": c_24,
                "year_25": c_25,
            }
        }
        save_settings(new_settings)
        settings.update(new_settings)
        st.success("✅ 설정이 저장됐습니다. 다음 STEP 실행부터 적용됩니다.")

    st.divider()
    st.markdown("**현재 저장된 설정**")
    st.json(load_settings())

    st.divider()
    st.markdown("**워크플로우 안내**")
    st.markdown("""
    ```
    ① 맥산 ERP에서 구매조회 또는 구매현황 파일 추출
              ↓
    ② STEP 1 실행 (ERP 원본 파일 업로드)
       → 자동 변환 후 사용내역.xlsx 다운로드
       → 신규 타입은 📋 품번 관리 탭에 자동 추가
              ↓
    ③ 📋 품번 관리 탭에서 품번·품명·구매처·리드타임 입력 후 저장
       (이후 실행부터 자동 적용)
              ↓
    ④ 사용내역.xlsx에 현재고·기발주 직접 입력
              ↓
    ⑤ STEP 2 실행 (동일 ERP 원본 파일 + 수정된 사용내역.xlsx 업로드)
       → 2026_연간발주계획.xlsx 다운로드
              ↓
    ⑥ 노란색 셀에 2026 목표 발주량 입력
    ```
    """)
