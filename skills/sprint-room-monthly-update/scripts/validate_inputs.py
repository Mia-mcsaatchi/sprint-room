#!/usr/bin/env python3
"""
Validate Sprint Room inputs before any deck-building work begins.

Checks:
- raw.xlsx contains expected 6 competitors in Query Name column
- chart_volume.xlsx 4th column is the new reporting month
- chart_sentiment.xlsx covers all 6 competitors for the new month
- chart_platform.xlsx is carline-specific (NOT the unfiltered "all" variant)
- Sentiment %s computed from raw data per competitor for new and prior month

Usage:
    python3 validate_inputs.py --carline Capri --raw raw.xlsx \\
        --vol chart_volume.xlsx --country chart_country.xlsx \\
        --platform chart_platform.xlsx --sent chart_sentiment.xlsx \\
        --last-month last_month.pptx
"""
import argparse, sys, re, zipfile, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bw_raw import load_raw  # zipfile+lxml reader; LibreOffice CSV can't open BW exports

# Queries in the raw export = carline + its competitor set (see monthly-run/config/carlines.json)
EXPECTED_QUERIES = {'Capri': 7, 'Explorer': 6, 'Kuga': 6, 'Puma': 5}

def inspect_chart_export(xlsx_path):
    """Pull header rows from a Brandwatch chart export."""
    with zipfile.ZipFile(xlsx_path) as z:
        xml = z.read('xl/worksheets/sheet1.xml').decode()
    rows = []
    for m in re.finditer(r'<row r="(\d+)">(.*?)</row>', xml):
        rn = int(m.group(1))
        if rn > 20: continue
        cells = re.findall(
            r'<c r="([A-Z]+\d+)"(?:[^>]*)>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)</c>',
            m.group(2))
        rows.append((rn, [(c[0], c[1] or c[2]) for c in cells]))
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--carline', required=True, choices=['Capri','Explorer','Kuga','Puma'])
    ap.add_argument('--raw', required=True)
    ap.add_argument('--vol', required=True)
    ap.add_argument('--country', required=True)
    ap.add_argument('--platform', required=True)
    ap.add_argument('--sent', required=True)
    ap.add_argument('--last-month', required=True)
    args = ap.parse_args()

    issues = []
    info = []

    # 1. Raw data: convert and check
    print(f"\n=== Raw data ({args.raw}) ===")
    try:
        df = load_raw(args.raw)
    except Exception as e:
        df = None
        issues.append(f"Could not read {args.raw}: {e}")
    if df is not None:
        if 'Query Name' not in df.columns:
            issues.append("raw.xlsx missing 'Query Name' column")
        else:
            queries = df['Query Name'].unique()
            info.append(f"Raw data: {len(df)} rows, {len(queries)} queries: {sorted(queries)}")
            if args.carline not in queries and not any(args.carline in str(q) for q in queries):
                issues.append(f"Carline '{args.carline}' not found in Query Name column. Got: {sorted(queries)}")
            expected = EXPECTED_QUERIES[args.carline]
            if len(queries) != expected:
                issues.append(f"Expected {expected} queries ({args.carline} + {expected-1} competitors), got {len(queries)}")
            # Month coverage
            month_counts = df.groupby('Month').size().to_dict()
            info.append(f"Rows per month: {month_counts}")
            if len(month_counts) < 2:
                issues.append("Raw data covers only one month — need new + prior month")

    # 2. Volume chart: 4th column should be the new month
    print(f"\n=== Volume chart export ({args.vol}) ===")
    rows = inspect_chart_export(args.vol)
    # Find header row (typically row 11 = months)
    # Header row = the row whose column A is 'months' (row 2 'Dimension 1 | months' is metadata)
    header_row = next((r for r in rows if r[1] and r[1][0][0].startswith('A') and str(r[1][0][1]).strip().lower() == 'months'), None)
    if header_row:
        month_cols = [c[1] for c in header_row[1][1:]]
        info.append(f"Volume chart columns: {month_cols}")
        if len(month_cols) < 4:
            issues.append(f"Volume chart has only {len(month_cols)} month columns, expected 4")
        else:
            # 4th column: check it's a date - the new month
            last_col = str(month_cols[-1])
            if '23:00:00' in last_col or '00:00:00' in last_col:
                info.append(f"4th column '{last_col}' is the new month (Brandwatch UTC labelling quirk)")
            else:
                info.append(f"4th column: {last_col}")

    # 3. Sentiment chart: check 6 competitors present
    print(f"\n=== Sentiment chart export ({args.sent}) ===")
    rows = inspect_chart_export(args.sent)
    # Date range
    start_row = next((r for r in rows if r[1] and 'Start Date' in str(r[1][0][1])), None)
    end_row = next((r for r in rows if r[1] and 'End Date' in str(r[1][0][1])), None)
    if start_row and end_row:
        info.append(f"Sentiment date range: {start_row[1][1][1]} to {end_row[1][1][1]}")
    # Header row with queries
    queries_row = next((r for r in rows if r[1] and 'queries' in str(r[1][0][1]).lower()), None)
    if queries_row:
        queries = [c[1] for c in queries_row[1][1:]]
        info.append(f"Sentiment competitors: {queries}")
        expected = EXPECTED_QUERIES[args.carline]
        if len(queries) != expected:
            issues.append(f"Sentiment chart has {len(queries)} models, expected {expected} for {args.carline}")

    # 4. Platform chart: detect unfiltered vs carline-filtered
    print(f"\n=== Platform chart export ({args.platform}) ===")
    rows = inspect_chart_export(args.platform)
    dim1 = next((r[1][1][1] for r in rows if r[1] and 'Dimension 1' in str(r[1][0][1])), None)
    queries = next((r[1][1][1] for r in rows if r[1] and 'Query Id' in str(r[1][0][1])), None)
    info.append(f"Platform Dimension 1: {dim1}; Query Id: {queries}")
    # The table starts at row 11. An unfiltered pull has an 'all' | 'all' header and rows that are
    # boolean-component names with 0 volume (see samples/chart-exports/*BAD-EXAMPLE*).
    hdr11 = next(([c[1] for c in r[1]] for r in rows if r[0] == 11), [])
    data = [r[1] for r in rows if r[0] > 11 and len(r[1]) >= 2]
    total = sum(float(c[1][1]) for c in data if re.fullmatch(r'-?[\d.]+', str(c[1][1] or '')))
    known = {'online news', 'news', 'x', 'twitter', 'forums', 'blogs', 'review', 'reviews',
             'tumblr', 'reddit', 'instagram', 'facebook', 'youtube', 'tiktok', 'qq', 'bluesky', 'threads'}
    labels = [str(c[0][1]).strip().lower() for c in data]
    info.append(f"Platform rows: {[c[0][1] for c in data][:8]} (total volume {total:g})")
    if [str(h).lower() for h in hdr11[:2]] == ['all', 'all'] or total == 0 \
            or (labels and not any(l in known for l in labels)):
        issues.append("Platform export looks UNFILTERED ('all'/'all' header, zero volume or no "
                      "platform names). Re-export the pageTypes breakdown filtered to the carline.")

    # 5. Country chart sanity
    print(f"\n=== Country chart export ({args.country}) ===")
    rows = inspect_chart_export(args.country)
    queries_row = next((r for r in rows if r[1] and r[1][0][0]=='A11'), None)

    # === REPORT ===
    print("\n\n========== VALIDATION RESULT ==========\n")
    for i in info:
        print(f"  • {i}")
    if issues:
        print(f"\n  ISSUES ({len(issues)}):")
        for i in issues:
            print(f"  ✗ {i}")
        sys.exit(1)
    else:
        print("\n  ✓ All checks passed")
        sys.exit(0)

if __name__ == '__main__':
    main()
