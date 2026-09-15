# SHIPPING PLAN 자동화 툴

본사→해외법인·해외법인→고객(BMW) 주간 스케줄 자동화. 매주 받는 **TRACING REPORT 2개**와 **SHIPPING PLAN**을 인보이스 번호로 매칭해, 도착 일정(ETA) 기준으로 재배선된 주차별 물량·인보이스가 반영된 완성형 쉬핑플랜을 만듭니다.

## 구성

| 파일 | 용도 |
|---|---|
| `index.html` | **브라우저용 웹 툴** — 파일 3개를 드래그드롭하면 자동으로 반영·다운로드 (GitHub Pages에서 바로 동작) |
| `apply_to_plan.py` | 로컬 자동화 — 파일 3개 경로를 인자로 주면 반영된 `SHIPPING_PLAN_REF.xlsx` 생성 |
| `auto_schedule.py` | 로컬 산출 — 8형번 스케줄 + 매칭요약 + 변경감지(`SHIPPING_SCHEDULE_AUTO.xlsx`) 생성 |
| `README.md` | 이 문서 |

> ⚠️ **업로드 금지**: `shipping_plan.xlsx`, `tracing_*.xlsx/xlsb`(실 인보이스·물량 데이터)는 이 저장소에 절대 커밋하지 마세요. 가짜 샘플 데이터로 대체하거나 private repo를 사용하세요.

## 사용법 (github Page 사용)

1. **배포**: GitHub에 이 폴더를 푸시 → Settings → Pages → **Deploy from a branch** → `main` / `(root)` → Save.
2. 링크: `https://<아이디>.github.io/<repo명>/` 에서 접속.
3. 화면에 **SHIPPING PLAN(.xlsx)**, **TRACING ①(.xlsx)**, **TRACING ②(.xlsb)** 3개를 드래그드롭 → 실행 → 반영된 엑셀을 다운로드.
   - 인터넷 필요(SheetJS 라이브러리를 CDN에서 로드).
   - 모든 처리는 브라우저 안(로컬)에서만 이루어집니다.

## 사용법 (로컬 python)

```bash
python3 -m pip install openpyxl
python3 apply_to_plan.py \
  --plan input/shipping_plan.xlsx \
  --tracing1 input/tracing_1.xlsx \
  --tracing2 input/tracing_2.xlsb \
  --out SHIPPING_PLAN_REF.xlsx
```

- 날짜 컬럼은 수식(`=AN260+7` 등)으로 저장되므로 **Excel에서 열면 자동 재계산**됩니다.
- `.xlsb`(TRACING ②)는 로컬 python이 가장 확실하게 처리합니다(브라우저 SheetJS 로딩에 문제가 있을 경우 대체 경로).

## 반영 규칙 (확정 로직)

- **수량** = SHIPPING PLAN 인보이스 텍스트 수량
- **주차 배선** = TRACING 도착 일정 기준 (ETD/계획 픽업 아님)
  - GEIS(문 도착): 도착일이 속한 주
  - POD(로테르담): 로테르담 도착 + 14일 이후 첫 월요일
- ETA 표기: DOOR 날짜가 있으면 `ETA GEIS`, 없으면 `ETA POD`(같은 인보이스 다건은 가장 늦은 일자)
- 도착일이 바뀌면 해당 주차의 물량·인보이스가 함께 이동(옛 주차 잔재 자동 제거)
