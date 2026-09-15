#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
인플레이스 반영 스크립트 - 파일 3개만 넣으면 완성형 쉬핑플랜을 바로 추출
- 원본 쉬핑플랜 복사본의 각 형번 시트 주차 행에, TRACING 도착 기준으로 재배선된
  스케줄(ETA GEIS/POD)·물량·인보이스를 직접 기록 (복붙 불필요)
- 읽기는 data_only=True(캐시값), 쓰기는 data_only=False(수식/서식 보존) 워크북 병행
"""
import argparse, re, sys, os
from datetime import date, datetime, timedelta
import openpyxl
from auto_schedule import (to_date, norm_inv, tag_date, read_tracing1, read_tracing2,
                           merge_tracing, resolve_eta, placement_week, find_header_cols, FORM_SHEETS)

FORMS = FORM_SHEETS

def parse_qty(line):
    m = re.search(r'([\d.,]+)\s*ET[AD]', line)
    return int(m.group(1).replace(",", "").replace(".", "")) if m else None

def sheet_buckets(read_ws, write_ws, win_start, win_end, buffer, tracing):
    """시트별: 주차(월요일 라벨) → [line] (TRACING 기준 재배선). 그리드 행 인덱스도 반환"""
    ic, pc, qc, ec = find_header_cols(write_ws)
    if not (isinstance(ic, int) and ic > 0 and isinstance(pc, int) and pc > 0):
        return None
    # 1) 그리드 행(주차 라벨 → 행번호) 수집 (캐시값 기준)
    grid_rows = []
    for r, row in enumerate(read_ws.iter_rows(min_row=6, max_row=read_ws.max_row,
                                              max_col=max(ic, pc) + 2, values_only=True), start=6):
        wk = to_date(row[pc - 1]) if pc <= len(row) else None
        if wk is None:
            continue
        if win_start and wk < win_start:
            continue
        if win_end and wk > win_end:
            continue
        grid_rows.append((r, wk))
    # 2) 인보이스 라인 재배선
    inv_buckets = {}   # inv -> (원주차, line)
    for r, wk in grid_rows:
        txt = None
        # 캐시값 워크북에서 인보이스 텍스트 읽기
        val = read_ws.cell(row=r, column=ic).value
        if isinstance(val, str):
            txt = val
        if not txt:
            continue
        for line in txt.split("\n"):
            line = line.strip()
            if not line:
                continue
            inv = norm_inv(line)
            if not inv:
                continue
            label, eta = resolve_eta(inv, tracing)
            plan_tag = tag_date(line, wk)
            if label is None and plan_tag:
                label, eta = plan_tag[0], plan_tag[1]
            pw = placement_week(label, eta, wk) if label else wk
            inv_buckets.setdefault(inv, []).append((wk, pw, line))
    # 3) 주차별로 ETA 재표기
    buckets = {}
    for inv, items in inv_buckets.items():
        for wk, pw, line in items:
            label, eta = resolve_eta(inv, tracing)
            plan_tag = tag_date(line, wk)
            if label is None and plan_tag:
                label, eta = plan_tag[0], plan_tag[1]
            if label and eta:
                base = re.sub(r'\s*ET[AD]\s+(POD|GEIS)\s+\d{1,2}\.\d{1,2}', '', line, flags=re.I).strip()
                line = f"{base} ETA {label} {eta.month}.{eta.day}"
            buckets.setdefault(pw, []).append(line)
    return {"buckets": buckets, "grid": grid_rows, "cols": (ic, pc, qc, ec)}

def apply_to_plan(write_ws, sheet_data):
    """주차 행의 invoice/qty 셀에 재배선 결과 기록"""
    ic, pc, qc, ec = sheet_data["cols"]
    week_row = {wk: r for r, wk in sheet_data["grid"]}
    if not week_row:
        return 0
    updated = 0
    # 1) 창 내 모든 그리드 행의 옛 인보이스/수량을 먼저 비움 (재배선으로 이동한 옛 주차의 잔재 제거)
    for wk, r in week_row.items():
        write_ws.cell(row=r, column=ic, value=None)
        write_ws.cell(row=r, column=qc, value=None)
    # 2) 재배선 결과를 각 주차 행에 기록
    for d, lines in sheet_data["buckets"].items():
        r = week_row.get(d)
        if r is None:
            continue
        text = "\n".join(lines) or None
        total = sum(parse_qty(l) or 0 for l in lines)
        write_ws.cell(row=r, column=ic, value=text)
        write_ws.cell(row=r, column=qc, value=total if total else None)
        updated += 1
    return updated

def monday(dd):
    x = date(dd.year, dd.month, dd.day)
    return x - timedelta(days=x.weekday())

def add_days(dd, n):
    x = date(dd.year, dd.month, dd.day)
    return x + timedelta(days=n)

def main():
    ap = argparse.ArgumentParser(description="파일 3개 → 반영된 쉬핑플랜 추출")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--tracing1", required=True)
    ap.add_argument("--tracing2", required=True)
    ap.add_argument("--out", default="SHIPPING_PLAN_REF.xlsx")
    ap.add_argument("--buffer", type=int, default=14)
    ap.add_argument("--start-off", type=int, default=7)
    ap.add_argument("--end")
    args = ap.parse_args()

    tracing = merge_tracing(read_tracing1(args.tracing1), read_tracing2(args.tracing2))
    print(f"TRACING 병합: {len(tracing)} 인보이스")

    today = date.today()
    win_start = monday(add_days(today, -args.start_off))
    if args.end:
        win_end = to_date(args.end)
    else:
        e = date(today.year, 12, 31)
        win_end = monday(e)
        if win_end < e:
            win_end = add_days(win_end, 7)
    print(f"반영 창: {win_start} ~ {win_end}")

    wb_w = openpyxl.load_workbook(args.plan, data_only=False)  # 쓰기용(수식 보존)
    wb_r = openpyxl.load_workbook(args.plan, data_only=True, read_only=True)  # 읽기용(캐시값)
    avail = set(wb_w.sheetnames)
    total = 0
    for form_name, prefixes in FORMS:
        for pre in prefixes:
            for sn in avail:
                if not sn.startswith(pre):
                    continue
                ws_w, ws_r = wb_w[sn], wb_r[sn]
                if ws_w.max_column < 30:
                    continue
                sd = sheet_buckets(ws_r, ws_w, win_start, win_end, args.buffer, tracing)
                if not sd:
                    continue
                n = apply_to_plan(ws_w, sd)
                if n:
                    total += n
                    print(f"  [{form_name}] '{sn}' 행 {n}개 반영")
    wb_r.close()
    wb_w.save(args.out)
    print(f"\n완료 → {args.out} (주차 {total}행 반영)")

if __name__ == "__main__":
    main()
