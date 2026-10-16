#!/usr/bin/env python3
"""
Update the volume line chart with new month's values.

Reads the chart_volume.xlsx export (4 months × 6 competitors) and:
- Replaces the 4 month labels with calendar names (Jan/Feb/Mar/Apr etc.)
- Replaces each series's 4 numeric values
- Replaces the Average row's 4 values (computed from competitors)

Series-name matching is done via the <c:v> series identifier inside <c:ser>.

Usage:
    python3 update_volume_chart.py path/to/chart9.xml --vol chart_volume.xlsx
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from chart_cache import set_cache
import argparse, re, zipfile

MONTH_NAMES = ['January','February','March','April','May','June','July','August',
               'September','October','November','December']

def read_volume_export(path):
    """Return (month_dates, competitor_data) where month_dates are date strings
    and competitor_data is {model_name: [v1,v2,v3,v4]}"""
    with zipfile.ZipFile(path) as z:
        xml = z.read('xl/worksheets/sheet1.xml').decode()
    rows = {}
    for m in re.finditer(r'<row r="(\d+)">(.*?)</row>', xml):
        rn = int(m.group(1))
        cells = re.findall(
            r'<c r="([A-Z]+\d+)"(?:[^>]*)>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)</c>',
            m.group(2))
        rows[rn] = [(c[0], c[1] or c[2]) for c in cells]
    # Row 11 = headers, 12+ = data
    headers = rows.get(11, [])
    month_labels = [c[1] for c in headers[1:5]]
    data = {}
    for rn in sorted(rows.keys()):
        if rn < 12: continue
        cells = rows[rn]
        if len(cells) < 5: continue
        model = cells[0][1]
        vals = [int(float(c[1])) for c in cells[1:5]]
        data[model] = vals
    return month_labels, data

def dates_to_month_names(date_strs):
    """'2026-04-30 23:00:00.0' → 'April'. The 4th column may have a different date than calendar."""
    names = []
    for s in date_strs:
        m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
        if m:
            mo = int(m.group(2))
            day = int(m.group(3))
            # If day is 31 (UTC quirk), bump to next month
            if day >= 30 and mo < 12:
                # likely captures next month's data
                mo = mo + 1 if mo < 12 else mo
            elif day >= 30 and mo == 12:
                mo = 1
            names.append(MONTH_NAMES[mo-1])
        else:
            names.append(s)
    return names

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('chart_xml')
    ap.add_argument('--vol', required=True)
    args = ap.parse_args()

    month_dates, comp_data = read_volume_export(args.vol)
    month_names = dates_to_month_names(month_dates)
    print(f"Month names: {month_names}")
    print(f"Competitors found: {list(comp_data.keys())}")

    with open(args.chart_xml) as f:
        xml = f.read()

    # 1. Replace month labels — use placeholder swap to avoid double-application
    # Read existing month labels from the chart
    cur_months = re.findall(r'<c:v>(January|February|March|April|May|June|July|August|September|October|November|December)</c:v>', xml)
    # Each unique month appears N times (one per series). Map old position to new.
    if cur_months:
        unique_cur = []
        for m in cur_months:
            if m not in unique_cur:
                unique_cur.append(m)
        # Replace via placeholder
        for i, old in enumerate(unique_cur):
            xml = xml.replace(f'<c:v>{old}</c:v>', f'<c:v>__M{i}__</c:v>')
        for i, new in enumerate(month_names):
            xml = xml.replace(f'<c:v>__M{i}__</c:v>', f'<c:v>{new}</c:v>')

    # 2. Replace each series's values — match by series name + old vals sequence
    # Find each <c:ser> block, identify its model name, swap in new values
    def replace_series_vals(block, new_vals):
        return set_cache(block, 'c:val', new_vals)

    # Map chart series name (e.g. "Enyaq", "ID.4") to export name (e.g. "Skoda Enyaq", "VW id4")
    name_map = {
        'Enyaq': ['Skoda Enyaq','Enyaq'],
        'XC40 Recharge/ EX40': ['Volvo XC40 Recharge/ EX40','XC40 Recharge/ EX40'],
        'E-3008': ['Peugeot E3008','Peugeot E-3008','E-3008'],
        'Megane E-Tech': ['Megane E-Tech'],
        'ID.4': ['VW id4','ID.4','ID4'],
    }
    # And the carline — search for the carline-like name in chart
    sers = re.findall(r'<c:ser>.*?</c:ser>', xml, re.DOTALL)
    new_sers = []
    for block in sers:
        name_m = re.search(r'<c:tx>.*?<c:v>([^<]+)</c:v>', block, re.DOTALL)
        if not name_m:
            new_sers.append(block); continue
        chart_name = name_m.group(1).strip()
        # Find matching export key
        export_name = None
        for k, aliases in name_map.items():
            if chart_name == k or chart_name in aliases:
                # Find export key matching
                for exp_k in comp_data:
                    if exp_k in aliases or k in exp_k:
                        export_name = exp_k; break
                break
        # If not in name_map, try direct match
        if not export_name:
            for exp_k in comp_data:
                if exp_k == chart_name or chart_name in exp_k or exp_k in chart_name:
                    export_name = exp_k; break
        if not export_name:
            # Probably the carline or Average
            if chart_name == 'Average':
                # compute from competitors
                avgs = [sum(v[i] for v in comp_data.values())/len(comp_data) for i in range(4)]
                new_block = replace_series_vals(block, [f'{a:.10g}' for a in avgs])
            else:
                # carline — match by name fragment
                for exp_k in comp_data:
                    if exp_k.lower() in chart_name.lower() or chart_name.lower() in exp_k.lower():
                        new_block = replace_series_vals(block, [str(v) for v in comp_data[exp_k]])
                        break
                else:
                    print(f"  warning: no export match for chart series '{chart_name}'")
                    new_sers.append(block); continue
        else:
            new_block = replace_series_vals(block, [str(v) for v in comp_data[export_name]])
            print(f"  {chart_name} ← {export_name}: {comp_data[export_name]}")
        new_sers.append(new_block)

    for old, new in zip(sers, new_sers):
        if old != new:
            xml = xml.replace(old, new, 1)

    with open(args.chart_xml, 'w') as f:
        f.write(xml)
    print(f"Updated {args.chart_xml}")

if __name__ == '__main__':
    main()
