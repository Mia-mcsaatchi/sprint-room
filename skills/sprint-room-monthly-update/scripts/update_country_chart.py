#!/usr/bin/env python3
"""
Update the country pie chart with new month's values + strip manualLayout from labels.

Usage:
    python3 update_country_chart.py path/to/chart10.xml --country chart_country.xlsx
"""
import argparse, re, zipfile

def read_country_export(path):
    """Return {country_code_or_name: count}."""
    with zipfile.ZipFile(path) as z:
        xml = z.read('xl/worksheets/sheet1.xml').decode()
    data = {}
    for m in re.finditer(r'<row r="(\d+)">(.*?)</row>', xml):
        rn = int(m.group(1))
        if rn < 12: continue
        cells = re.findall(
            r'<c r="([A-Z]+\d+)"(?:[^>]*)>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)</c>',
            m.group(2))
        cells = [(c[0], c[1] or c[2]) for c in cells]
        if len(cells) >= 2 and cells[0][1]:
            try:
                data[cells[0][1]] = int(float(cells[1][1]))
            except ValueError:
                pass
    return data

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('chart_xml')
    ap.add_argument('--country', required=True)
    args = ap.parse_args()

    new_data = read_country_export(args.country)
    print(f"Country data: {new_data}")

    with open(args.chart_xml) as f:
        xml = f.read()

    # Find current category order
    cats = re.findall(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    if cats:
        cat_labels = re.findall(r'<c:v>([^<]+)</c:v>', cats[0])
        print(f"Chart categories (current order): {cat_labels}")

        # Find current val numCache block
        val_match = re.search(r'<c:val>(.*?)</c:val>', xml, re.DOTALL)
        if val_match:
            val_block = val_match.group(0)
            old_nums = re.findall(r'<c:v>([^<]+)</c:v>', val_block)
            # Build new values in same category order
            country_aliases = {
                'UK': ['United Kingdom','UK','GBR'],
                'ES': ['Spain','ES','ESP'],
                'DE': ['Germany','DE','DEU'],
                'FR': ['France','FR','FRA'],
                'IT': ['Italy','IT','ITA'],
                'NL': ['Netherlands','NL','NLD'],
                'NO': ['Norway','NO','NOR'],
            }
            new_vals = []
            for label in cat_labels:
                v = None
                for alias_list in [country_aliases.get(label, [label])]:
                    for k in alias_list:
                        if k in new_data:
                            v = new_data[k]; break
                    if v is not None: break
                if v is None:
                    # Try direct
                    v = new_data.get(label, 0)
                new_vals.append(v)
            print(f"New values in category order: {list(zip(cat_labels, new_vals))}")

            new_val_block = val_block
            for o, n in zip(old_nums, new_vals):
                new_val_block = new_val_block.replace(f'<c:v>{o}</c:v>', f'<c:v>{n}</c:v>', 1)
            xml = xml.replace(val_block, new_val_block, 1)

    # Strip manualLayout from data labels (let labels auto-position around new slices)
    before = xml.count('<c:manualLayout>')
    xml = re.sub(r'\s*<c:layout>\s*<c:manualLayout>.*?</c:manualLayout>\s*</c:layout>', '',
                 xml, flags=re.DOTALL)
    after = xml.count('<c:manualLayout>')
    print(f"Stripped {before-after} manualLayout blocks (kept {after} non-label ones)")

    with open(args.chart_xml,'w') as f:
        f.write(xml)
    print(f"Updated {args.chart_xml}")

if __name__ == '__main__':
    main()
