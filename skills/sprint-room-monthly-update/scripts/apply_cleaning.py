#!/usr/bin/env python3
"""
Apply analyst cleaning decisions (Phase 1 exclusions + Phase 3 sentiment retags)
to a Brandwatch raw export, producing a cleaned export in the same layout.

Decisions CSV columns:
    row      0-based data-row index in the raw export (row order below the header)
    url      the row's Url - checked against the export so a stale decision can't
             hit the wrong row
    action   exclude | retag
    sentiment  new sentiment for retag (positive / neutral / negative)
    reason   free text, carried into the log

Outputs the cleaned xlsx (metadata block + header + kept rows, retags applied)
and a log CSV of every change. Downstream scripts read the cleaned file exactly
like the original.

Usage:
    python3 apply_cleaning.py --raw raw.xlsx --decisions cleaning.csv \\
        --out raw_clean.xlsx --log cleaning_log.csv
"""
import argparse, csv, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from raw_loader import _iter_rows
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", required=True)
    ap.add_argument("--decisions", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--log", required=True)
    a = ap.parse_args()

    with open(a.decisions, newline="") as f:
        decisions = {int(d["row"]): d for d in csv.DictReader(f)}

    meta, hdr, data = [], None, []
    for row in _iter_rows(a.raw):
        if hdr is None:
            if (row.get("B") or "").strip() == "Query Name":
                hdr = row
            else:
                meta.append(row)
            continue
        data.append(row)
    if hdr is None:
        sys.exit(f"{a.raw}: no header row with column B == 'Query Name'")
    col = {v: k for k, v in hdr.items()}
    url_c, sent_c, q_c = col["Url"], col["Sentiment"], col["Query Name"]

    errors = [f"row {i}: out of range" for i in decisions if i >= len(data)]
    errors += [f"row {i}: url mismatch ({d['url']!r} vs {data[i].get(url_c)!r})"
               for i, d in decisions.items() if i < len(data) and d["url"] != data[i].get(url_c)]
    errors += [f"row {i}: bad action/sentiment {d['action']!r}/{d.get('sentiment')!r}"
               for i, d in decisions.items()
               if d["action"] not in ("exclude", "retag")
               or (d["action"] == "retag" and d.get("sentiment") not in ("positive", "neutral", "negative"))]
    if errors:
        sys.exit("Decisions don't match this export:\n  " + "\n  ".join(errors))

    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    width = max(column_index_from_string(c) for r in [hdr] + data for c in r)

    def emit(row):
        out = [None] * width
        for c, v in row.items():
            out[column_index_from_string(c) - 1] = v
        ws.append(out)

    log, kept = [], 0
    for r in meta:
        emit(r)
    emit(hdr)
    for i, row in enumerate(data):
        d = decisions.get(i)
        if d and d["action"] == "exclude":
            log.append([i, row.get(q_c), "exclude", row.get(sent_c), "", row.get(url_c), d["reason"]])
            continue
        if d and d["action"] == "retag" and d["sentiment"] != row.get(sent_c):
            log.append([i, row.get(q_c), "retag", row.get(sent_c), d["sentiment"], row.get(url_c), d["reason"]])
            row = {**row, sent_c: d["sentiment"]}
        emit(row)
        kept += 1
    wb.save(a.out)

    with open(a.log, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["row", "query", "action", "old_sentiment", "new_sentiment", "url", "reason"])
        w.writerows(log)
    n_ex = sum(1 for l in log if l[2] == "exclude")
    print(f"{len(data)} rows in, {kept} kept, {n_ex} excluded, {len(log) - n_ex} retagged -> {a.out}")


if __name__ == "__main__":
    main()
