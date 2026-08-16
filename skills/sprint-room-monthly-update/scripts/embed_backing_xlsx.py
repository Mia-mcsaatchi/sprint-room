#!/usr/bin/env python3
"""
Embed editable backing xlsx files so PowerPoint's "Edit Data in Excel" works.

CRITICAL: each chart gets its OWN xlsx file. Do not share one file across charts —
PowerPoint rejects that as malformed. (See references/embed_xlsx_pitfalls.md.)

The xlsx files mirror what the chart's <c:f> formulas reference:
- chart9 (volume) → 'Explorer July' sheet, A11:E18
- chart10 (country) → 'Explorer July' sheet, A22:B28
- chart11 (sentiment) → 'Explorer July' sheet, I11:O17
- chart12 (platform) → 'Explorer' sheet, A12:B17

Sheet names are inherited from the original template — they aren't carline-specific.

Usage:
    python3 embed_backing_xlsx.py path/to/unpacked/
"""
import argparse, os, re, sys
import openpyxl
from openpyxl.styles import Font

def extract_chart_data(chart_xml_path):
    """Read current values from a chart so we can mirror them in xlsx."""
    with open(chart_xml_path) as f:
        xml = f.read()
    return xml

def get_volume_data(xml):
    """From volume chart9.xml, get month labels + per-series values."""
    cat_m = re.search(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    months = re.findall(r'<c:v>([^<]+)</c:v>', cat_m.group(1)) if cat_m else []
    series = []
    for ser in re.finditer(r'<c:ser>(.*?)</c:ser>', xml, re.DOTALL):
        b = ser.group(1)
        name = re.search(r'<c:tx>.*?<c:v>([^<]+)</c:v>', b, re.DOTALL)
        val_m = re.search(r'<c:val>(.*?)</c:val>', b, re.DOTALL)
        if name and val_m:
            vals = re.findall(r'<c:v>([^<]+)</c:v>', val_m.group(1))
            series.append((name.group(1).strip(), [float(v) for v in vals]))
    return months, series

def get_country_data(xml):
    cat_m = re.search(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    cats = re.findall(r'<c:v>([^<]+)</c:v>', cat_m.group(1)) if cat_m else []
    val_m = re.search(r'<c:val>(.*?)</c:val>', xml, re.DOTALL)
    vals = re.findall(r'<c:v>([^<]+)</c:v>', val_m.group(1)) if val_m else []
    return list(zip(cats, [int(float(v)) for v in vals]))

def get_platform_data(xml):
    cat_m = re.search(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    cats = re.findall(r'<c:v>([^<]+)</c:v>', cat_m.group(1)) if cat_m else []
    # numCache only has pts for non-deleted categories
    pts = re.findall(r'<c:pt idx="(\d+)">\s*<c:v>([^<]+)</c:v>\s*</c:pt>', xml)
    # Map idx to category
    result = []
    for cat_idx, cat_label in enumerate(cats):
        v = 0
        for pt_idx, pt_v in pts:
            if int(pt_idx) == cat_idx:
                try: v = int(float(pt_v))
                except: v = 0
                break
        result.append((cat_label, v))
    return result

def get_sentiment_data(xml):
    """Returns dict: {pos:[...], neu:[...], neg:[...], pos_pre:[...], neg_pre:[...]}"""
    out = {}
    series_map = {'positive':'pos','neutral':'neu','negative':'neg',
                  'positive pre':'pos_pre','Negative Pre':'neg_pre'}
    cat_m = re.search(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    out['categories'] = re.findall(r'<c:v>([^<]+)</c:v>', cat_m.group(1)) if cat_m else []
    for ser in re.finditer(r'<c:ser>(.*?)</c:ser>', xml, re.DOTALL):
        b = ser.group(1)
        name_m = re.search(r'<c:tx>.*?<c:v>([^<]+)</c:v>', b, re.DOTALL)
        if not name_m: continue
        name = name_m.group(1).strip()
        if name not in series_map: continue
        key = series_map[name]
        # Y vals or X vals?
        is_scatter = name in ('positive pre','Negative Pre')
        tag = 'c:xVal' if is_scatter else 'c:val'
        val_m = re.search(rf'<{tag}>(.*?)</{tag}>', b, re.DOTALL)
        if val_m:
            out[key] = [float(v) for v in re.findall(r'<c:v>([^<]+)</c:v>', val_m.group(1))]
    return out

FONT = Font(name='Arial', size=10)
def stylize(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.font = FONT

def build_volume_xlsx(out_path, months, series):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Explorer July'
    ws['A11'] = 'months'
    for i, mo in enumerate(months):
        ws.cell(row=11, column=2+i).value = mo
    # series in chart order
    for i, (model, vals) in enumerate(series):
        r = 12 + i
        ws.cell(row=r, column=1).value = model
        for j, v in enumerate(vals):
            ws.cell(row=r, column=2+j).value = round(v) if model != 'Average' else v
    # Set Average row to live formulas if it exists
    # Find Average row
    for r in range(12, 19):
        if ws.cell(row=r, column=1).value == 'Average':
            last_data_r = r - 1
            for j in range(len(months)):
                col_letter = chr(ord('B')+j)
                ws.cell(row=r, column=2+j).value = f'=AVERAGE({col_letter}12:{col_letter}{last_data_r})'
            break
    stylize(ws)
    wb.save(out_path)

def build_country_xlsx(out_path, country_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Explorer July'
    for i, (k,v) in enumerate(country_data):
        ws.cell(row=22+i, column=1).value = k
        ws.cell(row=22+i, column=2).value = v
    stylize(ws)
    wb.save(out_path)

def build_sentiment_xlsx(out_path, sent_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Explorer July'
    # I12-I17 = sentiment category labels
    labels = ['positive','neutral','negative','positive pre','Negative Pre','Y']
    for i, lab in enumerate(labels):
        ws.cell(row=12+i, column=9).value = lab
    # J11:O11 = competitor names
    cats = sent_data.get('categories', [])
    for i, c in enumerate(cats):
        ws.cell(row=11, column=10+i).value = c
    # Data rows
    for key, row in [('pos',12),('neu',13),('neg',14),('pos_pre',15),('neg_pre',16)]:
        vals = sent_data.get(key, [])
        for i, v in enumerate(vals):
            ws.cell(row=row, column=10+i).value = v
            if row < 17:
                ws.cell(row=row, column=10+i).number_format = '0.0%'
    # Y values
    y_vals = [3, 8.75, 14.6, 20.5, 26.3, 32.1]
    for i, v in enumerate(y_vals):
        ws.cell(row=17, column=10+i).value = v
    stylize(ws)
    wb.save(out_path)

def build_platform_xlsx(out_path, platform_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Explorer'
    for i, (k,v) in enumerate(platform_data):
        ws.cell(row=12+i, column=1).value = k
        ws.cell(row=12+i, column=2).value = v
    stylize(ws)
    wb.save(out_path)

def repoint_rels(unpacked, chart_n, xlsx_name):
    rels_path = os.path.join(unpacked, 'ppt/charts/_rels', f'chart{chart_n}.xml.rels')
    with open(rels_path) as f:
        c = f.read()
    new_rel = f'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/{xlsx_name}"/>'
    c = re.sub(r'<Relationship Id="rId4"[^>]*/>', new_rel, c)
    with open(rels_path,'w') as f:
        f.write(c)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('unpacked')
    args = ap.parse_args()

    embeddings_dir = os.path.join(args.unpacked, 'ppt/embeddings')
    os.makedirs(embeddings_dir, exist_ok=True)
    # Clear any prior xlsx
    for f in os.listdir(embeddings_dir):
        if f.endswith('.xlsx'):
            os.remove(os.path.join(embeddings_dir, f))

    # Build each xlsx from the current chart data
    charts_dir = os.path.join(args.unpacked, 'ppt/charts')

    # Volume
    xml = extract_chart_data(os.path.join(charts_dir, 'chart9.xml'))
    months, series = get_volume_data(xml)
    build_volume_xlsx(os.path.join(embeddings_dir, 'Microsoft_Excel_Worksheet.xlsx'), months, series)
    print(f"Built volume backing xlsx: {len(months)} months, {len(series)} series")

    # Country
    xml = extract_chart_data(os.path.join(charts_dir, 'chart10.xml'))
    country = get_country_data(xml)
    build_country_xlsx(os.path.join(embeddings_dir, 'Microsoft_Excel_Worksheet1.xlsx'), country)
    print(f"Built country backing xlsx: {len(country)} entries")

    # Sentiment
    xml = extract_chart_data(os.path.join(charts_dir, 'chart11.xml'))
    sent = get_sentiment_data(xml)
    build_sentiment_xlsx(os.path.join(embeddings_dir, 'Microsoft_Excel_Worksheet2.xlsx'), sent)
    print(f"Built sentiment backing xlsx: {len(sent.get('categories',[]))} models")

    # Platform
    xml = extract_chart_data(os.path.join(charts_dir, 'chart12.xml'))
    platform = get_platform_data(xml)
    build_platform_xlsx(os.path.join(embeddings_dir, 'Microsoft_Excel_Worksheet3.xlsx'), platform)
    print(f"Built platform backing xlsx: {len(platform)} entries")

    # Repoint chart rels
    repoint_rels(args.unpacked, 9, 'Microsoft_Excel_Worksheet.xlsx')
    repoint_rels(args.unpacked, 10, 'Microsoft_Excel_Worksheet1.xlsx')
    repoint_rels(args.unpacked, 11, 'Microsoft_Excel_Worksheet2.xlsx')
    repoint_rels(args.unpacked, 12, 'Microsoft_Excel_Worksheet3.xlsx')
    print("Repointed chart9-12 rId4 → internal package")

    # Register in Content_Types (cleanup-and-add to avoid duplicates)
    ct_path = os.path.join(args.unpacked, '[Content_Types].xml')
    with open(ct_path) as f:
        ct = f.read()
    # Remove ANY prior xlsx Override (handles whitespace variants)
    ct = re.sub(r'\s*<Override PartName="/ppt/embeddings/Microsoft_Excel_Worksheet[0-9]*\.xlsx"[^>]*/>', '', ct)
    # Add fresh
    new_entries = ''
    for n in ['', '1', '2', '3']:
        new_entries += f'  <Override PartName="/ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>\n'
    ct = ct.replace('</Types>', new_entries + '</Types>')
    with open(ct_path,'w') as f:
        f.write(ct)

    # Verify no duplicate Override entries
    overrides = re.findall(r'<Override PartName="/ppt/embeddings/[^"]+"', ct)
    if len(overrides) != len(set(overrides)):
        print(f"ERROR: duplicate Override entries: {overrides}")
        sys.exit(1)
    print(f"Registered {len(overrides)} xlsx Override entries (unique)")

if __name__ == '__main__':
    main()
