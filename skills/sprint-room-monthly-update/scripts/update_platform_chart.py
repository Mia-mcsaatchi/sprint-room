#!/usr/bin/env python3
"""
Update the platform pie chart with new month's values + strip manualLayout from labels.

The platform chart's numCache has only 4 visible values typically (Online News, X,
Forums, Blogs) — Review and Tumblr are hidden via <c:delete val="1"/>. Keep the
delete flags; only update the visible values.

Usage:
    python3 update_platform_chart.py path/to/chart12.xml --platform chart_platform.xlsx
"""
import argparse, re, zipfile

def read_platform_export(path):
    """Return {platform_label: count}. Filter out zero rows."""
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
                v = int(float(cells[1][1]))
                data[cells[0][1]] = v
            except ValueError:
                pass
    return data

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('chart_xml')
    ap.add_argument('--platform', required=True)
    args = ap.parse_args()

    new_data = read_platform_export(args.platform)
    print(f"Platform data: {new_data}")

    with open(args.chart_xml) as f:
        xml = f.read()

    # Category labels in chart (typically Online News, X, Forums, Blogs, Review, Tumblr)
    cats = re.findall(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    if cats:
        cat_labels = re.findall(r'<c:v>([^<]+)</c:v>', cats[0])
        print(f"Chart categories: {cat_labels}")

    # Update numCache — but the cache may only have N entries even with M categories.
    # Map by index in cat_labels, but skip pt indices that don't exist in the original cache.
    cache_match = re.search(r'<c:numCache>(.*?)</c:numCache>', xml, re.DOTALL)
    if cache_match:
        cache = cache_match.group(1)
        # Get existing pts with their idx
        pts = re.findall(r'<c:pt idx="(\d+)">\s*<c:v>([^<]+)</c:v>\s*</c:pt>', cache)
        print(f"Existing pts: {pts}")
        # Build new pts: for each existing idx, find matching category label, get new count
        new_pts_xml = []
        platform_aliases = {
            'X': ['X','Twitter','twitter'],
            'Online News': ['Online News','News','news'],
            'Forums': ['Forums','forum','Forum'],
            'Blogs': ['Blogs','blog','Blog'],
        }
        for idx, _old_v in pts:
            i = int(idx)
            if i < len(cat_labels):
                label = cat_labels[i]
                v = None
                for alias_k, aliases in platform_aliases.items():
                    if label in aliases or label == alias_k:
                        for a in aliases:
                            if a in new_data:
                                v = new_data[a]; break
                        break
                if v is None:
                    v = new_data.get(label, 0)
                new_pts_xml.append(f'<c:pt idx="{idx}"><c:v>{v}</c:v></c:pt>')
                print(f"  idx={idx} {label}: {v}")
        # Replace numCache content (keep formatCode, ptCount)
        format_m = re.search(r'<c:formatCode>[^<]+</c:formatCode>', cache)
        ptcount_m = re.search(r'<c:ptCount val="\d+"/>', cache)
        new_cache_inner = ''
        if format_m: new_cache_inner += format_m.group(0) + '\n'
        if ptcount_m: new_cache_inner += ptcount_m.group(0) + '\n'
        new_cache_inner += '\n'.join(new_pts_xml)
        xml = xml.replace(cache, '\n' + new_cache_inner + '\n', 1)

    # Strip manualLayout from non-deleted labels
    # Be careful: deleted labels (<c:delete val="1"/>) should keep their manualLayout
    # since they're suppressed anyway. So we only strip manualLayout from labels without c:delete.
    def strip_layout(match):
        block = match.group(0)
        if '<c:delete val="1"/>' in block:
            return block  # don't touch deleted labels
        # Remove the c:layout > c:manualLayout
        return re.sub(r'\s*<c:layout>\s*<c:manualLayout>.*?</c:manualLayout>\s*</c:layout>', '',
                      block, flags=re.DOTALL)
    xml = re.sub(r'<c:dLbl>.*?</c:dLbl>', strip_layout, xml, flags=re.DOTALL)

    with open(args.chart_xml,'w') as f:
        f.write(xml)
    print(f"Updated {args.chart_xml}")

if __name__ == '__main__':
    main()
