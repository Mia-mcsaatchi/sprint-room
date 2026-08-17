#!/usr/bin/env python3
"""
link_embeddings.py — make the new slide's charts "Edit Data in Excel"-linked with the CURRENT
(new-month) values.

build_skeleton_raw gives every new chart its own embedding, but that copy still holds last
month's numbers. This writes each chart's own values into its own backing workbook by reading the
chart's <c:f> cell references + the paired num/strCache — so the workbook always matches the chart
exactly, whatever the per-carline cell layout is. The volume 'Average' row is written as a live
=AVERAGE() formula so it recomputes if an analyst edits a competitor value.

Usage:
    python3 link_embeddings.py unpacked/ --slide slide6.xml
"""
import argparse, os, re, sys
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

def parse_range(ref):
    """'Sheet1'!$B$11:$E$11 -> (sheet, [(row,col),...]) in idx order (row- or column-vector)."""
    m = re.match(r"(?:'([^']+)'|([^!]+))!(.+)", ref)
    if not m: return None, []
    sheet = m.group(1) or m.group(2)
    rng = m.group(3).replace('$', '')
    if ':' not in rng:
        c = re.match(r'([A-Z]+)(\d+)', rng)
        return sheet, [(int(c.group(2)), column_index_from_string(c.group(1)))]
    a, b = rng.split(':')
    ca = re.match(r'([A-Z]+)(\d+)', a); cb = re.match(r'([A-Z]+)(\d+)', b)
    c1, r1 = column_index_from_string(ca.group(1)), int(ca.group(2))
    c2, r2 = column_index_from_string(cb.group(1)), int(cb.group(2))
    cells = []
    if r1 == r2:      # row vector
        for c in range(c1, c2 + 1): cells.append((r1, c))
    else:             # column vector
        for r in range(r1, r2 + 1): cells.append((r, c1))
    return sheet, cells

def num(v):
    try:
        f = float(v); return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return v

def link_chart(chart_xml_path, emb_path):
    xml = open(chart_xml_path, encoding='utf-8').read()
    pairs = re.findall(r'<c:f>([^<]+)</c:f>\s*<c:(numCache|strCache)>(.*?)</c:\2>', xml, re.S)
    wb = openpyxl.load_workbook(emb_path)
    avg_rows = {}  # sheet -> (row, [model_rows]) for the volume Average
    label_rows = {}  # sheet -> {row: labelcol A text}
    # first, write every referenced cell from its cache
    for ref, _kind, cache in pairs:
        sheet, cells = parse_range(ref)
        if not sheet: continue
        ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
        pts = dict(re.findall(r'<c:pt idx="(\d+)"[^>]*>\s*<c:v>([^<]*)</c:v>', cache))
        for i, (r, c) in enumerate(cells):
            if str(i) in pts:
                ws.cell(row=r, column=c, value=num(pts[str(i)]))
    # detect the volume Average row (a column-A label == 'Average') and make it a live formula
    for ref, kind, cache in pairs:
        sheet, cells = parse_range(ref)
        if not sheet or not cells: continue
        if kind == 'strCache' and len(cells) == 1 and cells[0][1] == 1:  # single A-col label
            pts = re.findall(r'<c:v>([^<]*)</c:v>', cache)
            label_rows.setdefault(sheet, {})[cells[0][0]] = pts[0] if pts else ''
    for sheet, rows in label_rows.items():
        ws = wb[sheet]
        avg_row = next((r for r, lab in rows.items() if lab.strip().lower() == 'average'), None)
        if not avg_row: continue
        model_rows = sorted(r for r in rows if r < avg_row)
        if not model_rows: continue
        r0, r1 = model_rows[0], model_rows[-1]
        for col in range(2, 6):  # B..E
            L = get_column_letter(col)
            ws.cell(row=avg_row, column=col, value=f"=AVERAGE({L}{r0}:{L}{r1})")
    wb.save(emb_path)
    return len(pairs)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('unpacked')
    ap.add_argument('--slide', required=True, help='the new-month slide, e.g. slide6.xml')
    a = ap.parse_args()
    up = a.unpacked
    rels = open(f"{up}/ppt/slides/_rels/{a.slide}.rels").read()
    charts = re.findall(r'charts/(chart\d+\.xml)', rels)
    if not charts:
        sys.exit(f"No charts on {a.slide}")
    for c in charts:
        crels = open(f"{up}/ppt/charts/_rels/{c}.rels").read()
        em = re.search(r'embeddings/([^"]+\.xlsx)', crels)
        if not em:
            print(f"  {c}: no embedding rel — skipped"); continue
        emb = f"{up}/ppt/embeddings/{em.group(1)}"
        n = link_chart(f"{up}/ppt/charts/{c}", emb)
        print(f"  {c} → {em.group(1)}: wrote {n} ranges")
    print("Embeddings now carry the new-month values (charts are Edit-Data linked).")

if __name__ == '__main__':
    main()
