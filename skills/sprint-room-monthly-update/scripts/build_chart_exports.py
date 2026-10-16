#!/usr/bin/env python3
"""
Build the four chart-data exports from a (cleaned) raw export, in Brandwatch's
"Chart Data Export" layout, for when the Brandwatch dashboard doesn't reflect the
analyst's cleaning (exclusions / retags applied with apply_cleaning.py).

Writes into --out-dir:
    chart_country.xlsx    carline x country, reporting month
    chart_platform.xlsx   carline x platform, reporting month
    chart_sentiment.xlsx  all queries x sentiment, reporting month
    chart_volume.xlsx     4 rolling months x queries - only with --last-month:
                          the 3 earlier months are carried forward from the
                          published volume chart in last month's deck (never
                          re-derived from raw), the reporting month comes from raw.

Files use inline strings and plain <row r="N"> tags, which is what the
update_*_chart.py readers parse.

Usage:
    python3 build_chart_exports.py --carline Capri --raw raw_clean.xlsx \\
        --month 2026-09 --out-dir inputs/Capri/2026-09 [--last-month last_month.pptx]
"""
import argparse, os, re, sys, zipfile
from xml.sax.saxutils import escape
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from raw_loader import load_raw

COUNTRIES = ["Germany", "United Kingdom", "Spain", "Italy", "Netherlands", "France", "Norway"]
PLATFORMS = {"news": "Online News", "twitter": "X", "forum": "Forums", "blog": "Blogs",
             "review": "Review", "tumblr": "Tumblr"}
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


def _col(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def write_export(path, meta, table):
    """meta: [(label, value)] for rows 1-9; table: list of rows starting at row 11."""
    rows = [["Chart Data Export"]] + [list(m) for m in meta] + [[]] + table
    xml = []
    for rn, row in enumerate(rows, 1):
        cells = []
        for ci, v in enumerate(row):
            ref = f"{_col(ci)}{rn}"
            if isinstance(v, (int, float)):
                cells.append(f'<c r="{ref}"><v>{v}</v></c>')
            elif v is not None:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(v))}</t></is></c>')
        xml.append(f'<row r="{rn}">{"".join(cells)}</row>')
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<sheetData>{"".join(xml)}</sheetData></worksheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>')
        z.writestr("_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
        z.writestr("xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '</Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def prior_volume(pptx, queries, last_label):
    """{query: [v1, v2, v3, v4]} and month names from last month's published volume
    chart - the 4-month chart whose final month is `last_label` (the deck also
    carries the month before's slide, with its own volume chart)."""
    with zipfile.ZipFile(pptx) as z:
        charts = [n for n in z.namelist() if re.fullmatch(r"ppt/charts/chart\d+\.xml", n)]
        for name in charts:
            xml = z.read(name).decode("utf-8")
            cat = re.search(r"<c:cat>(.*?)</c:cat>", xml, re.S)
            if not cat:
                continue
            labels = re.findall(r"<c:v>([^<]+)</c:v>", cat.group(1))
            if len(labels) != 4 or labels[-1] != last_label or not all(l in MONTHS for l in labels):
                continue
            out = {}
            for ser in re.findall(r"<c:ser>.*?</c:ser>", xml, re.S):
                nm = re.search(r"<c:tx>.*?<c:v>([^<]+)</c:v>", ser, re.S)
                val = re.search(r"<c:val>(.*?)</c:val>", ser, re.S)
                if not (nm and val):
                    continue
                vals = [float(v) for v in re.findall(r"<c:v>([^<]+)</c:v>", val.group(1))]
                key = next((q for q in queries if _norm(nm.group(1)) in _norm(q)), None)
                if key and len(vals) == 4:
                    out[key] = vals
            return name, labels, out
    sys.exit(f"No 4-month volume chart ending in {last_label} found in {pptx}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carline", required=True)
    ap.add_argument("--raw", required=True)
    ap.add_argument("--month", required=True, help="reporting month YYYY-MM")
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--last-month", help="last month's deck, for the volume chart history")
    a = ap.parse_args()

    yr, mm = map(int, a.month.split("-"))
    df = load_raw(a.raw)
    cur = df[(df["Date"].dt.year == yr) & (df["Month"] == mm)]
    queries = list(dict.fromkeys(df["Query Name"]))
    carline = next(q for q in queries if a.carline.lower() in q.lower())
    car = cur[cur["Query Name"] == carline]
    meta = lambda d1, d2: [("Dimension 1", d1), ("Dimension 2", d2), ("Aggregate", "volume"),
                           ("Start Date", f"{a.month}-01"), ("End Date", f"{a.month} (end)"),
                           ("Project Id", ""), ("Location", "[FRA, GBR, DEU, ESP, NLD, NOR, ITA]"),
                           ("Query Id", f"built from {os.path.basename(a.raw)}")]
    out = lambda f: os.path.join(a.out_dir, f)

    counts = car["Country"].value_counts()
    write_export(out("chart_country.xlsx"), meta("queries", "countries"),
                 [["queries", carline]] + [[c, int(counts.get(c, 0))] for c in COUNTRIES])

    counts = car["Page Type"].map(PLATFORMS).value_counts()
    write_export(out("chart_platform.xlsx"), meta("queries", "pageTypes"),
                 [["queries", carline]] + [[p, int(counts.get(p, 0))] for p in PLATFORMS.values()])

    sent = cur.groupby("Query Name")["Sentiment"].value_counts().unstack(fill_value=0)
    write_export(out("chart_sentiment.xlsx"), meta("queries", "sentiment"),
                 [["queries"] + queries] +
                 [[s] + [int(sent.loc[q].get(s, 0)) if q in sent.index else 0 for q in queries]
                  for s in ("positive", "neutral", "negative")])
    print(f"Wrote chart_country / chart_platform / chart_sentiment for {carline}, {a.month}")

    if not a.last_month:
        print("chart_volume.xlsx NOT written: pass --last-month to carry forward the earlier months.")
        return
    src, labels, hist = prior_volume(a.last_month, queries, MONTHS[(mm - 2) % 12])
    missing = [q for q in queries if q not in hist]
    if missing:
        sys.exit(f"{src}: no series matched {missing} (series names vs query names)")
    vol = cur["Query Name"].value_counts()
    month_cols = []
    for back in (3, 2, 1, 0):
        m, y = mm - back, yr
        if m <= 0:
            m, y = m + 12, y - 1
        month_cols.append(f"{y}-{m:02d}-01 00:00:00.0")
    table = [["months"] + month_cols]
    for q in queries:
        table.append([q] + [int(v) for v in hist[q][1:]] + [int(vol.get(q, 0))])
    write_export(out("chart_volume.xlsx"), meta("months", "queries"), table)
    print(f"Wrote chart_volume: {labels[1:]} carried forward from {src}, {MONTHS[mm - 1]} from raw")


if __name__ == "__main__":
    main()
