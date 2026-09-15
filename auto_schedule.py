#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SHIPPING PLAN 자동화 툴
- 본사-해외법인(CIF, Korea→Germany) TRACING REPORT 2개 + SHIPPING PLAN 엑셀 매칭
- 인보이스별 도착예정일(ETA GEIS / ETA POD) 및 주차별 물량을 주차별 그리드로 추출

[요구사항]
1) TRACING REPORT 2개 파일을 SHIPPING PLAN 인보이스 넘버와 매칭하여
   도착예정일·도착예정 물량을 복붙 가능한 형태로 추출
2) ETA ROTTERDAM까지만 표기 시 "ETA POD", DOOR까지 업데이트 시 "ETA GEIS"
   같은 인보이스의 여러 건 일자가 상이하면 가장 늦은 일자 기준
3) ETA POD 기준 건은 SHIPPING PLAN상 2주 후 도착으로 반영(계획 시트 배치 유지)
4) 물량 없는 주차도 빈칸으로 생성, 8개 형번별 탭 분리

사용법:
  python3 auto_schedule.py --plan <shipping_plan.xlsx> \
      --tracing1 <tracing_report1.xlsx> --tracing2 <tracing_report2.xlsb> \
      [-o <출력.xlsx>]
  (파일명 자동감지: 커맨드라인 미지정 시 ./input/ 기본 경로 사용)

출력:
  <출력.xlsx>  - 8개 형번 탭 + "매칭요약" 탭
"""
import argparse, re, sys, os
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------------- config
# 형번 탭명 -> SHIPPING PLAN 시트명 후보
FORM_SHEETS = [
    ("FAAR WE MID ND", ["FAAR WE ND "]),
    ("FAAR WE MID D",  ["FAAR WE D "]),   # 두 시트(증량관련/154K) 병합, 직납증량 제외됨
    ("35UP HD",        ["High D "]),
    ("G8X D",          ["G8X D "]),
    ("G8X ND",         ["G8X ND "]),
    ("XHIGH TK120",    ["XHigh TK120 "]),
    ("NCAR ND",        ["NCAR ND "]),
    ("NCAR D",         ["NCAR D "]),
]
INV_RE = re.compile(r'(IJG[-_ ]?\d{5,6}[A-Z0-9_\-()]*)', re.I)
QTY_RE = re.compile(r'([\d.,]+)\s*ETA\b')
TAG_RE = re.compile(r'\s*ET[AD]\s+(POD|GEIS)\s+(\d{1,2})\.(\d{1,2})', re.I)

# ---------------------------------------------------------------- helpers
def to_date(v):
    """엑셀 셀 값 -> date. (숫자 serial / datetime / str 모두 처리)"""
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        if 20000 < v < 60000:          # 엑셀 날짜 serial
            try:
                return (datetime(1899, 12, 30) + timedelta(days=v)).date()
            except OverflowError:
                return None
        return None
    s = str(v).strip()
    # 날짜 + 시간 문자열 ("2026-09-14 00:00:00") → 날짜부만
    m = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    for f in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None

def norm_inv(txt):
    m = INV_RE.search(str(txt))
    return m.group(1).upper().replace(" ", "") if m else None

def collapse_key(inv):
    """연속 중복 영문자를 압축 (오타 내성: 'FAAAR'→'FAAR').
    예) TRACING에 'IJG-260729FAAAR(ND)'로 기재된 경우 계획 'IJG-260729FAAR(ND)'와 매칭"""
    return re.sub(r'(?i)([A-Z])\1+', r'\1', inv) if inv else inv

def inv_keys(txt):
    """인보이스 텍스트에서 등록용 키 목록(정확·압축) 반환"""
    inv = norm_inv(txt)
    if not inv:
        return []
    keys = [inv]
    c = collapse_key(inv)
    if c and c != inv:
        keys.append(c)
    return keys

def month_day(d):
    """날짜를 'M.D' 형식으로 (8.13 / 9.4 / 10.10)"""
    return f"{d.month}.{d.day}"

def parse_qty(line):
    """문장에서 '3600 ETA' / '7,200 ETA' / '10.800 ETA' 앞의 수량 추출"""
    m = QTY_RE.search(line)
    if not m:
        return None
    try:
        return int(m.group(1).replace(",", "").replace(".", ""))
    except ValueError:
        return None

def tag_date(line, week):
    """계획 셀의 'ETA GEIS M.D' / 'ETA POD M.D' 태그 → (label, date). 태그 없으면 None.
    태그는 년도가 없으므로 주차(week)를 기준으로 가장 가까운 연도를 추정한다."""
    m = TAG_RE.search(line)
    if not m:
        return None
    mm, dd = int(m.group(2)), int(m.group(3))
    best = None
    for y in (week.year - 1, week.year, week.year + 1):
        try:
            d = datetime(y, mm, dd).date()
        except ValueError:
            continue
        if best is None or abs((d - week).days) < abs((best - week).days):
            best = d
    return (m.group(1).upper(), best)

# ---------------------------------------------------------------- tracing read
def read_tracing1(path):
    """(VOB) ILJIN GLOBAL TRACING REPORT (xlsx)"""
    out = {}   # inv -> list of (pod_date, door_date)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=12, max_col=26, values_only=True):
        inv_txt = row[4] if len(row) > 4 else None
        keys = inv_keys(inv_txt)
        if not keys:
            continue
        pod = to_date(row[14] if len(row) > 14 else None)
        door = to_date(row[19] if len(row) > 19 else None)
        # DOOR 도착완료일시(실도착)가 있을 경우 그게 최신의 확정 일자
        arrived = row[22] if len(row) > 22 else None
        if door is None:
            door = to_date(arrived)
        for k in keys:
            out.setdefault(k, []).append((pod, door))
    wb.close()
    return out

def read_tracing2(path):
    """.xlsb (일진글로벌 Shipping Status Report) - 모든 시트 처리"""
    out = {}
    try:
        from pyxlsb import open_workbook as xlsb_open
    except ImportError:
        print("[경고] pyxlsb 미설치. tracing2(.xlsb)는 건너뜁니다.")
        return out
    with xlsb_open(path) as wb:
        for shname in wb.sheets:
            with wb.get_sheet(shname) as sh:
                for row in sh.rows():
                    if len(row) < 10:
                        continue
                    keys = inv_keys(row[4].v)
                    if not keys:
                        continue
                    pod = to_date(row[8].v)
                    door = to_date(row[9].v)
                    for k in keys:
                        out.setdefault(k, []).append((pod, door))
    return out

def merge_tracing(t1, t2):
    merged = {}
    for src in (t1, t2):
        for inv, pairs in src.items():
            merged.setdefault(inv, []).extend(pairs)
    return merged

def resolve_eta(inv, tracing):
    """
    인보이스별 최종 도착 기준 반환: (label, date)
      - DOOR 존재 시 "GEIS" + 최신 DOOR
      - DOOR 없고 ETA ROTTERDAM만 존재 시 "POD" + 최신 POD
      - 둘 다 없으면 (None, None)
    정확 키 우선, 없으면 연속중복 압축 키로 재시도(오타 내성)
    """
    pairs = []
    keys = [inv] if inv else []
    ck = collapse_key(inv)
    if ck and ck != inv and ck not in keys:
        keys.append(ck)
    for k in keys:
        pairs.extend(tracing.get(k, []))
    doors = [d for _, d in pairs if d]
    pods = [p for p, _ in pairs if p]
    if doors:
        return "GEIS", max(doors)
    if pods:
        return "POD", max(pods)
    return None, None

# ---------------------------------------------------------------- plan read
def find_header_cols(ws):
    """시트 헤더(상위 5행)에서 Invoice / Pick up / Geis Inbound / ETA GEIS 컬럼 탐색
    (FAAR D 계열은 'Pick up'이 없고 'Date' 헤더를 사용하므로 Date도 주간 그리드로 감지)"""
    inv_col = pick_col = qty_col = eta_col = None
    for r in range(1, 6):
        for row in ws.iter_rows(min_row=r, max_row=r):
            for cell in row:
                if cell.value is None:
                    continue
                s = str(cell.value).lower()
                if inv_col is None and "invoice" in s:
                    inv_col = cell.column
                if pick_col is None and "pick up" in s:
                    pick_col = cell.column
                if qty_col is None and "geis" in s and "inbound" in s:
                    qty_col = cell.column
                if eta_col is None and s.strip().lower() == "eta geis":
                    eta_col = cell.column
    # 주간 그리드 컬럼: 'Pick up'이 없으면 'Date' 정확일치 헤더 사용 (FAAR D 계열)
    if pick_col is None:
        for r in range(1, 6):
            for row in ws.iter_rows(min_row=r, max_row=r):
                for cell in row:
                    if cell.value is not None and str(cell.value).strip().lower() == "date":
                        pick_col = cell.column
                        break
                if pick_col:
                    break
            if pick_col:
                break
    if qty_col is None and inv_col and inv_col > 1:
        qty_col = inv_col - 1
    return inv_col, pick_col, qty_col, eta_col

def extract_plan_sheet(ws, horizon_start=None, horizon_end=None):
    """
    시트에서 주차별 그리드 + 인보이스 행 추출.
    Returns: list of dict {week(date Monday), lines:[(inv, line_text)]}
    """
    inv_col, pick_col, qty_col, eta_col = find_header_cols(ws)
    grid_col = pick_col or eta_col
    if not (inv_col and grid_col):
        return [], (inv_col, pick_col, qty_col, eta_col)

    rows = []
    for r, row in enumerate(ws.iter_rows(min_row=6, max_row=ws.max_row,
                                         min_col=1, max_col=max(inv_col, grid_col or 0) + 3,
                                         values_only=True), start=6):
        week = None
        if grid_col <= len(row):
            week = to_date(row[grid_col - 1])
        inv_text = None
        if inv_col <= len(row):
            inv_text = row[inv_col - 1]
        if week is None and not (inv_text and str(inv_text).strip()):
            continue
        rows.append({"row": r, "week": week, "inv_text": inv_text})

    # 주차별 집계 (인보이스 라인 분해)
    grid = {}
    for item in rows:
        week = item["week"]
        txt = item["inv_text"]
        if week is None or txt is None:
            continue
        if horizon_start and week < horizon_start:
            continue
        if horizon_end and week > horizon_end:
            continue
        # 인보이스 라인 분리 (여러 줄 셀)
        for line in str(txt).split("\n"):
            line = line.strip()
            if not line:
                continue
            inv = norm_inv(line)
            if not inv:
                continue
            grid.setdefault(week, []).append((inv, line))
    return grid, (inv_col, pick_col, qty_col, eta_col)

def placement_week(label, eta_date, fallback_week):
    """tracing 도착 기준 배선(픽업) 주차 계산
      - GEIS(문 도착) : 도착일이 속한 주(월요일)
      - POD(로테르담) : 로테르담 도착 + 2주 버퍼. 주차 라벨(M)을 해당 주 구간 [M-6, M]으
        로 보므로(예: 09-07 = 09-01~09-07), 배선 주차는 (POD+14일) 이상인 첫 월요일 →
        버퍼가 항상 14~20일 (2주 이상) 보장
      - 미매칭       : 계획서 픽업 주차 유지(fallback_week)
    """
    if label is None or eta_date is None:
        return fallback_week
    if label == "GEIS":
        base = eta_date
    else:
        base = eta_date + timedelta(days=14)
    # 해당 날짜(구간 [base-6, base] 중 최소 라벨 = base 이상인 첫 월요일)
    monday = base - timedelta(days=base.weekday())
    if monday < base:
        monday += timedelta(days=7)
    return monday

# ---------------------------------------------------------------- output build
def build_form_tab(form_name, week_lines, tracing, sheet_meta, win_start, win_end):
    """week_lines: {주차(Monday) -> [(inv, line_text)]}, tracing: 인보이스별 (pod,door)
    - 주차 배정은 tracing 도착 일정 기준(placement_week)으로 재계산 → 물량·내용 함께 이동
    - win_start~win_end 구간 전체 주차 그리드 생성(빈 주차 포함), 미매칭은 계획 주차 유지
    Returns: excel_rows[(week,total,text)], unmatched, inv_det
    """
    buckets = {}
    excel_rows = []
    unmatched = []
    inv_det = []                  # 변경감지용: 인보이스별 기준(계획) / 신규(TRACING) ETA + 주차
    for week, lines in week_lines.items():
        for inv, line in lines:
            label, eta_date = resolve_eta(inv, tracing)
            plan_tag = tag_date(line, week)      # 계획서에 수기 기록된 ETA 태그(기준)
            kept = False
            if label is None and plan_tag is not None:
                # TRACING 미기재지만 계획서에 ETA 기록이 있으면 → 계획 ETA·계획 주차 유지해 표시
                label, eta_date = plan_tag[0], plan_tag[1]
                pw = week
                kept = True
            else:
                pw = placement_week(label, eta_date, week)
            inv_det.append({"form": form_name, "inv": inv, "week": week, "line": line,
                            "plan_tag": plan_tag, "new_tag": (label, eta_date),
                            "new_week": pw, "kept": kept})
            if label is None:
                unmatched.append((inv, line, week))
                continue          # 일자 정보가 전혀 없음 → 출력 제외(주차는 빈칸 유지)
            base = TAG_RE.sub(" ", line).strip()              # 기존 ETA/ETD 태그 제거
            base = re.sub(r"\s{2,}", " ", base).strip()
            new_line = f"{base} ETA {label} {month_day(eta_date)}"
            buckets.setdefault(pw, []).append((inv, new_line))

    # 그리드: win_start~win_end (tracing 기준 배선이 창 밖이면 확장)
    g_end = win_end
    if buckets:
        g_end = max(win_end, max(buckets))
    d = win_start
    while d <= g_end:
        lines = buckets.get(d, [])
        keep = []
        total = 0
        for inv, new_line in lines:
            keep.append((inv, new_line))
            q = parse_qty(new_line)
            if q is not None:
                total += q
        if keep:
            text = "\n".join(l for _, l in keep)
        else:
            text = None
        excel_rows.append((d, total, text))
        d += timedelta(days=7)
    return excel_rows, unmatched, inv_det

# ---------------------------------------------------------------- excel writer
HEADER_FILL = PatternFill("solid", fgColor="5B5B6D")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BORDER = Border(*[Side(style="thin", color="BFBFBF")] * 4)
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

def write_workbook(forms_data, out_path, today_txt, change_rows=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    meta_header = ["Pick up from ILJIN Germany W/H for Tier1", "Geis Inbound Qty", "Invoice"]
    widths = [26, 16, 95]
    all_unmatched = {}

    for form_name, rows, unmatched in forms_data:
        ws = wb.create_sheet(title=form_name[:31])
        ws.append(meta_header)
        for c in range(1, 4):
            cell = ws.cell(row=1, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN
        ws.column_dimensions["A"].width = widths[0]
        ws.column_dimensions["B"].width = widths[1]
        ws.column_dimensions["C"].width = widths[2]
        r = 2
        for week, qty_num, text in rows:
            ws.cell(row=r, column=1, value=week.strftime("%Y-%m-%d"))
            # 수량은 '숫자'로 저장 → 엑셀 상태표시줄 합계 / SUM 자동 계산 됨
            qcell = ws.cell(row=r, column=2)
            if qty_num:
                qcell.value = qty_num
                qcell.number_format = "#,##0"
                qcell.font = Font(bold=True)
            else:
                qcell.value = 0
                qcell.number_format = "#,##0"
                qcell.font = Font(color="9E9E9E")
            ws.cell(row=r, column=3, value=text)
            for c in range(1, 4):
                cell = ws.cell(row=r, column=c)
                cell.border = THIN
                cell.alignment = Alignment(vertical="top", wrap_text=(c == 3))
                if c == 1:
                    cell.alignment = Alignment(vertical="top")
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="right", vertical="top")
            r += 1
        ws.freeze_panes = "A2"

        # 주차별 수량 합계행 (수기 파일 물량과 대조용)
        sr = ws.max_row + 2
        ws.cell(row=sr, column=1, value="합계 (Geis Inbound Qty)")
        ws.cell(row=sr, column=1).font = Font(bold=True)
        sc = ws.cell(row=sr, column=2, value=f"=SUM(B2:B{sr-2})")
        sc.number_format = "#,##0"
        sc.font = Font(bold=True)
        sc.fill = PatternFill("solid", fgColor="DDEBF7")
        for c in (1, 2):
            ws.cell(row=sr, column=c).border = THIN
        if unmatched:
            all_unmatched[form_name] = unmatched

    # 매칭요약 탭
    ws = wb.create_sheet(title="매칭요약")
    ws.append(["형번", "추출 인보이스 수", "TRACING 매칭 성공", "미매칭(미확정) 건수"])
    for c in range(1, 5):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.border = THIN
    r = 2
    for form_name, rows, unmatched in forms_data:
        matched = sum(1 for _, q, t in rows if t)
        ws.cell(row=r, column=1, value=form_name)
        ws.cell(row=r, column=2, value=len(rows))
        ws.cell(row=r, column=3, value=matched)
        ws.cell(row=r, column=4, value=len(unmatched))
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = THIN
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="미매칭 인보이스 상세(같은 번호가 TRACING에 없어 일자 미확정)")
    ws.cell(row=r, column=1).font = Font(bold=True)
    r += 1
    ws.append(["형번", "인보이스", "계획상 주차", "계획 셀 원문(전체)"])
    for c in range(1, 5):
        ws.cell(row=r, column=c).font = Font(bold=True)
        ws.cell(row=r, column=c).border = THIN
    r += 1
    for form_name, um in all_unmatched.items():
        for inv, line, week in um:
            ws.cell(row=r, column=1, value=form_name)
            ws.cell(row=r, column=2, value=inv)
            ws.cell(row=r, column=3, value=week.strftime("%Y-%m-%d") if week else "")
            ws.cell(row=r, column=4, value=line)
            for c in range(1, 5):
                ws.cell(row=r, column=c).border = THIN
            r += 1
    for col, w in zip("ABCD", [20, 26, 14, 120]):
        ws.column_dimensions[col].width = w

    # 변경감지 탭: 기준(계획서) ETA vs 신규(TRACING) ETA - 변경/지연 인보이스만
    ws = wb.create_sheet(title="변경감지")
    DELAY_FILL = PatternFill("solid", fgColor="FFC7CE")   # 지연(빨강)
    EARLY_FILL = PatternFill("solid", fgColor="C6EFCE")   # 앞당김(초록)
    MISS_FILL = PatternFill("solid", fgColor="FFEB9C")    # 미기재(노랑)
    headers = ["형번", "인보이스", "배선 계획 주차", "기준(계획서) ETA",
               "신규(TRACING) ETA", "상태", "변화(일)", "비고",
               "배선 주차(계획→TRACING)", "계획 셀 원문"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.border = THIN
    r = 2
    for row in (change_rows or []):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = THIN
            cell.alignment = Alignment(vertical="top", wrap_text=(c == len(headers)))
        status = str(row[5] if row[5] else "")
        fill = None
        if status.startswith("지연"): fill = DELAY_FILL
        elif status.startswith("앞당김"): fill = EARLY_FILL
        elif "미기재" in status: fill = MISS_FILL
        if fill:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = fill
        r += 1
    for col, w in zip("ABCDEFGHI", [14, 24, 14, 16, 16, 16, 9, 30, 22, 120]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    wb.save(out_path)
    return out_path

# ---------------------------------------------------------------- main
def default_input_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "input")

def main():
    ap = argparse.ArgumentParser(description="SHIPPING PLAN 자동화 툴")
    ap.add_argument("--plan", help="SHIPPING PLAN xlsx 경로")
    ap.add_argument("--tracing1", help="TRACING REPORT 1 (xlsx, VOB)")
    ap.add_argument("--tracing2", help="TRACING REPORT 2 (xlsb)")
    ap.add_argument("-o", "--out", default="SHIPPING_SCHEDULE_AUTO.xlsx", help="출력 xlsx")
    ap.add_argument("--start", help="그리드 시작일(YYYY-MM-DD), 기본=계획 최소주차")
    ap.add_argument("--end", help="그리드 종료일(YYYY-MM-DD), 기본=계획 최대주차")
    args = ap.parse_args()

    indir = default_input_dir()
    plan = args.plan or os.path.join(indir, "shipping_plan.xlsx")
    tr1 = args.tracing1 or os.path.join(indir, "tracing_1.xlsx")
    tr2 = args.tracing2 or os.path.join(indir, "tracing_2.xlsb")

    for p in (plan, tr1, tr2):
        if not os.path.exists(p):
            sys.exit(f"[오류] 파일 없음: {p}")

    print("== TRACING 읽는 중 ==")
    t1 = read_tracing1(tr1)
    t2 = read_tracing2(tr2)
    tracing = merge_tracing(t1, t2)
    print(f"  TRACING1 인보이스 {len(t1)}건 / TRACING2 {len(t2)}건 / 병합 {len(tracing)}건")

    print("== SHIPPING PLAN 읽는 중 ==")
    wb = openpyxl.load_workbook(plan, read_only=True, data_only=True)
    available = set(wb.sheetnames)
    horizon_start = to_date(args.start)
    horizon_end = to_date(args.end)

    today_d = date.today()
    # 기본 그리드 창: (오늘 - 7일) 이후의 픽업부터 ~ 올해 마지막 월요일
    win_start = (today_d - timedelta(days=7)) - timedelta(days=(today_d - timedelta(days=7)).weekday())
    win_end = date(today_d.year, 12, 31)
    win_end = win_end - timedelta(days=(win_end.weekday() - 0) % 7)
    if horizon_start is None:
        horizon_start = win_start
    if horizon_end is None:
        horizon_end = win_end
    print(f"  그리드 창: {horizon_start} ~ {horizon_end}")

    forms_data = []
    all_inv_det = []               # 변경감지용 전체 수집
    for form_name, candidates in FORM_SHEETS:
        combined = {}
        used = []
        for cand in candidates:
            for sn in available:
                if not sn.startswith(cand):
                    continue
                ws = wb[sn]
                if ws.max_column < 30:      # 'FAAR WE D 직납증량' 같은 요약시트 제외
                    continue
                grid, meta = extract_plan_sheet(ws, horizon_start, horizon_end)
                for week, lines in grid.items():
                    # 동일 인보이스 중복 제거(시트 병합 시)
                    seen = {l for _, l in combined.get(week, [])}
                    for line in lines:
                        if line not in seen:
                            combined.setdefault(week, []).append(line)
                            seen.add(line)
                used.append(sn)
        if not used:
            print(f"  [경고] {form_name}: 시트 없음")
            forms_data.append((form_name, [], []))
            continue
        rows, unmatched, inv_det = build_form_tab(form_name, combined, tracing, None,
                                                  horizon_start, horizon_end)
        forms_data.append((form_name, rows, unmatched))
        all_inv_det.extend(inv_det)
        matched = sum(1 for _, q, t in rows if t)
        print(f"  {form_name:12}: 시트 {used} / 주차 {len(rows)} / 채워진 주차 {matched} / 미매칭 {len(unmatched)}")

    wb.close()

    # 변경감지: 기준(계획서 기록 ETA)이 있고 신규 TRACING ETA와 달라진 인보이스만
    change_rows = []
    for det in all_inv_det:
        plan = det["plan_tag"]
        nl, nd = det["new_tag"]
        if plan is None:
            continue                       # 기준 태그가 없는 건은 변경감지에서 제외
        if nd is None:
            status = "트레이싱 미기재 (기준만)"
            note = "TRACING에 해당 인보이스가 없음(출항 전 or 누락)"
            delta = ""
        else:
            delta = (nd - plan[1]).days
            if delta == 0:
                continue                   # 동일 → 제외
            if nd > plan[1]:
                status = f"지연 {delta}일"
            else:
                status = f"앞당김 {-delta}일"
            note = ""
            if nl != plan[0]:
                note = f"기준 {plan[0]} → 신규 {nl}"
        change_rows.append([det["form"], det["inv"],
                            det["week"].strftime("%Y-%m-%d") if det["week"] else "",
                            f"{plan[0]} {month_day(plan[1])}",
                            f"{nl} {month_day(nd)}" if nd else "-",
                            status, delta, note,
                            det["new_week"].strftime("%Y-%m-%d") if det.get("new_week") else "",
                            det["line"]])
    change_rows.sort(key=lambda x: (x[4] or ""))     # 신규 ETA 순 정렬
    print(f"변경감지: 총 {len(change_rows)}건 (계획서 태그 대비 변경/지연/미기재)")

    write_workbook(forms_data, args.out, datetime.today().strftime("%Y-%m-%d"), change_rows)
    print(f"\n완료 → {args.out}")

if __name__ == "__main__":
    main()
