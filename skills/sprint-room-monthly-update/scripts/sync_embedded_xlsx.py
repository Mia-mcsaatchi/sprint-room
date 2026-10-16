#!/usr/bin/env python3
"""
Bring each new chart's embedded workbook ("Edit Data in Excel") in line with the
chart's updated values, for decks whose charts already carry their own embedded
xlsx (build_skeleton.py gives charts 9-12 private copies of last month's).

Values are placed by label, not by the chart's cell references (those are not
always reliable in hand-edited decks):
  - grid layout (row 1 = category labels, col A = series names): volume and
    sentiment charts. Row 1 is rewritten with the chart's categories, then each
    series row gets its values.
  - list layout (col A = category label, col B = value): single-series pies.
Formula cells (e.g. the volume Average row) are left alone so Excel recomputes them.

Usage:
    python3 sync_embedded_xlsx.py path/to/unpacked/ [--charts 9 10 11 12]
"""
import argparse, os, re, sys
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from chart_cache import cache_values, norm


def chart_series(xml):
    out = []
    for block in re.findall(r"<c:ser>.*?</c:ser>", xml, re.DOTALL):
        nm = re.search(r"<c:tx>.*?<c:v>([^<]+)</c:v>", block, re.DOTALL)
        cat = re.search(r"<c:cat>(.*?)</c:cat>", block, re.DOTALL)
        cats = re.findall(r"<c:v>([^<]+)</c:v>", cat.group(1)) if cat else []
        vals = cache_values(block, "c:val") or cache_values(block, "c:xVal")
        out.append((nm.group(1).strip() if nm else None, cats, [float(v) if v else 0 for v in vals]))
    return out


def is_formula(cell):
    return isinstance(cell.value, str) and cell.value.startswith("=")


def sync(wb_path, xml):
    wb = openpyxl.load_workbook(wb_path)
    ws = wb.worksheets[0]
    series = chart_series(xml)
    col_a = {norm(str(ws.cell(r, 1).value)): r for r in range(1, ws.max_row + 1)
             if ws.cell(r, 1).value is not None}
    cats = next((c for _, c, _ in series if c), [])
    written = 0
    grid = any(name and norm(name) in col_a for name, _, _ in series) and len(series) > 1
    if grid:
        for i, c in enumerate(cats):
            ws.cell(1, 2 + i).value = c
        for name, _, vals in series:
            r = col_a.get(norm(name or ""))
            if r is None:
                sys.exit(f"{wb_path}: no row for series {name!r}")
            for i, v in enumerate(vals):
                cell = ws.cell(r, 2 + i)
                if not is_formula(cell):
                    cell.value = v
                    written += 1
    else:
        name, cats, vals = series[0]
        for c, v in zip(cats, vals):
            r = col_a.get(norm(c))
            if r is None:
                sys.exit(f"{wb_path}: no row labelled {c!r}")
            ws.cell(r, 2).value = v
            written += 1
    wb.save(wb_path)
    return "grid" if grid else "list", written


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("unpacked")
    ap.add_argument("--charts", nargs="+", type=int, default=[9, 10, 11, 12])
    a = ap.parse_args()
    for n in a.charts:
        rels = open(os.path.join(a.unpacked, f"ppt/charts/_rels/chart{n}.xml.rels")).read()
        m = re.search(r'Type="[^"]*/package" Target="\.\./(embeddings/[^"]+)"', rels)
        if not m:
            sys.exit(f"chart{n} has no embedded workbook - use embed_backing_xlsx.py instead")
        wb = os.path.join(a.unpacked, "ppt", m.group(1))
        xml = open(os.path.join(a.unpacked, f"ppt/charts/chart{n}.xml")).read()
        layout, k = sync(wb, xml)
        print(f"chart{n} -> {m.group(1)}: {layout} layout, {k} cells written")


if __name__ == "__main__":
    main()
