# AJW 발주계획 자동화 시스템

(주)에이제이월드 SCM팀 — 생산자재 발주계획 자동화 웹앱

## 실행 방법

```bash
"AI/연간 발주 계획 새우기/.venv/Scripts/streamlit.exe" run "AI/연간 발주 계획 새우기/web_app.py" --server.address=0.0.0.0 --server.port=8501
```

접속: http://localhost:8501

## 3-STEP 워크플로우

| STEP | 입력 | 출력 |
|------|------|------|
| STEP 1 — ERP 파일 가공 | 맥산 구매현황 Excel | 생산자재 사용내역 + 가공파일 |
| STEP 2 — 판매 분석 | 전체 판매량 Excel | OJC 판매량 정리 + 생산비중 분석 |
| STEP 3 — 발주계획 생성 | STEP 1 가공파일 | 연간발주계획 Excel |

## 주요 파일

```
AI/연간 발주 계획 새우기/
├── web_app.py        # Streamlit 웹앱 진입점
├── step1_core.py     # STEP 1 — ERP 파싱 및 가공파일 생성
├── step2_core.py     # STEP 3 — 발주계획 Excel 생성
├── convert_core.py   # ERP 원본 파싱 (구매조회/구매현황)
├── ojc_filter.py     # OJC 완제품 판매량 필터링
├── metadata.json     # 품번·품명·구매처·리드타임
├── inventory.json    # 현재고·기발주
└── sales_analysis.json  # STEP 2 판매 분석 결과
```

## 담당자

박정원 (SCM팀) — jw.park@ajw.co.kr
