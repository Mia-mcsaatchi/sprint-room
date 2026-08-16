#!/usr/bin/env python3
"""
Update the sentiment combo chart (stacked bars + diamond scatter).

Bars = new month's pos/neu/neg %s.
Diamonds = prior month's pos % (positive pre) and 1 - prior month's neg % (Negative Pre).

The PRIOR month diamond values come from the PRIOR month's published chart (chart3.xml
inside last_month.pptx), NOT from the new raw data. Last month's published values are
what the client saw; using newly-retagged numbers creates a mismatch with the
comparison slide.

Also strips any manual <a:t>X%</a:t> data-label overrides — they go stale.

Usage:
    python3 update_sentiment_chart.py path/to/chart11.xml \\
        --new-month-sent chart_sentiment.xlsx \\
        --prior-from-last-month last_month.pptx
"""
import argparse, re, zipfile

def read_sentiment_export(path):
    """Return {model: {pos:N, neu:N, neg:N}}."""
    with zipfile.ZipFile(path) as z:
        xml = z.read('xl/worksheets/sheet1.xml').decode()
    rows = {}
    for m in re.finditer(r'<row r="(\d+)">(.*?)</row>', xml):
        rn = int(m.group(1))
        cells = re.findall(
            r'<c r="([A-Z]+\d+)"(?:[^>]*)>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)</c>',
            m.group(2))
        rows[rn] = [(c[0], c[1] or c[2]) for c in cells]
    # Row 11 = headers (queries), 12=pos, 13=neu, 14=neg
    header_cells = rows.get(11, [])
    queries = [c[1] for c in header_cells[1:]]
    data = {q: {} for q in queries}
    for rn, sent_key in [(12,'pos'),(13,'neu'),(14,'neg')]:
        if rn not in rows: continue
        cells = rows[rn]
        for i, q in enumerate(queries):
            if i+1 < len(cells):
                try:
                    data[q][sent_key] = int(float(cells[i+1][1]))
                except (ValueError, IndexError):
                    data[q][sent_key] = 0
    return data

def read_prior_chart3(pptx_path):
    """Read the prior month's chart3 (sentiment) data — positive and negative series values."""
    with zipfile.ZipFile(pptx_path) as z:
        xml = z.read('ppt/charts/chart3.xml').decode()
    result = {}
    for ser_match in re.finditer(r'<c:ser>(.*?)</c:ser>', xml, re.DOTALL):
        block = ser_match.group(1)
        name_m = re.search(r'<c:tx>.*?<c:v>([^<]+)</c:v>', block, re.DOTALL)
        if not name_m: continue
        name = name_m.group(1).strip()
        if name not in ('positive','negative'): continue
        val_m = re.search(r'<c:val>(.*?)</c:val>', block, re.DOTALL)
        if val_m:
            vals = re.findall(r'<c:v>([^<]+)</c:v>', val_m.group(1))
            result[name] = vals
    return result

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('chart_xml')
    ap.add_argument('--new-month-sent', required=True)
    ap.add_argument('--prior-from-last-month', required=True)
    args = ap.parse_args()

    new_sent = read_sentiment_export(args.new_month_sent)
    print(f"New sentiment: {new_sent}")
    prior_chart3 = read_prior_chart3(args.prior_from_last_month)
    print(f"Prior chart3 positive: {prior_chart3.get('positive')}")
    print(f"Prior chart3 negative: {prior_chart3.get('negative')}")

    with open(args.chart_xml) as f:
        xml = f.read()

    # Find chart's category order (which models, in what order)
    # We use the existing series val counts to figure out where each model sits
    # in the chart. The chart's series ORDER inside <c:ser> isn't ordered by model;
    # rather each series (positive, neutral, etc) has 6 values, one per model in
    # category order. We need to know that order.
    # The c:cat block inside any series has the model list.
    cat_match = re.search(r'<c:cat>(.*?)</c:cat>', xml, re.DOTALL)
    if not cat_match:
        print("ERROR: no c:cat found in chart"); return
    chart_cats = re.findall(r'<c:v>([^<]+)</c:v>', cat_match.group(1))
    print(f"Chart category order: {chart_cats}")

    # Map chart category names to export query names
    def find_export_match(chart_name, queries):
        for q in queries:
            if q.strip().lower() == chart_name.strip().lower(): return q
        for q in queries:
            cl = chart_name.strip().lower(); ql = q.strip().lower()
            if cl in ql or ql in cl: return q
        return None

    new_vals_pos, new_vals_neu, new_vals_neg = [], [], []
    for cn in chart_cats:
        q = find_export_match(cn, list(new_sent.keys()))
        if not q:
            print(f"  WARN: no export match for chart cat '{cn}'")
            new_vals_pos.append(0); new_vals_neu.append(0); new_vals_neg.append(0)
            continue
        c = new_sent[q]
        tot = c['pos']+c['neu']+c['neg']
        if tot == 0:
            new_vals_pos.append(0); new_vals_neu.append(0); new_vals_neg.append(0)
        else:
            new_vals_pos.append(c['pos']/tot)
            new_vals_neu.append(c['neu']/tot)
            new_vals_neg.append(c['neg']/tot)

    # Diamond values from prior chart3 (in same category order)
    prior_pos = [float(v) for v in prior_chart3.get('positive', [])]
    prior_neg = [float(v) for v in prior_chart3.get('negative', [])]
    new_pos_pre = prior_pos
    new_neg_pre = [1 - n for n in prior_neg]

    print(f"\nNew bar values:")
    for i, cn in enumerate(chart_cats):
        print(f"  {cn}: pos {new_vals_pos[i]*100:.2f}% / neu {new_vals_neu[i]*100:.2f}% / neg {new_vals_neg[i]*100:.2f}%")
    print(f"\nDiamond X (from prior chart3):")
    for i, cn in enumerate(chart_cats):
        print(f"  {cn}: pos {new_pos_pre[i]*100:.2f}% / 1-neg {new_neg_pre[i]*100:.2f}%")

    # Replace each series's values atomically
    def replace_series(content, anchor, new_list, kind='val'):
        # kind: 'val' or 'xVal'
        m = re.search(rf'<c:v>{re.escape(anchor)}</c:v>.*?</c:ser>', content, re.DOTALL)
        if not m: return content, 0
        block = m.group(0)
        target_tag = 'c:xVal' if kind=='xVal' else 'c:val'
        val_m = re.search(rf'<{target_tag}>(.*?)</{target_tag}>', block, re.DOTALL)
        if not val_m: return content, 0
        val_block = val_m.group(0)
        old_nums = re.findall(r'<c:v>([^<]+)</c:v>', val_block)
        new_val_block = val_block
        n = 0
        for o, new_v in zip(old_nums, new_list):
            t = f'<c:v>{o}</c:v>'
            if t in new_val_block:
                new_val_block = new_val_block.replace(t, f'<c:v>{new_v:.17g}</c:v>', 1)
                n += 1
        new_block = block.replace(val_block, new_val_block, 1)
        return content.replace(block, new_block, 1), n

    xml, n1 = replace_series(xml, 'positive', new_vals_pos, 'val')
    xml, n2 = replace_series(xml, 'neutral', new_vals_neu, 'val')
    xml, n3 = replace_series(xml, 'negative', new_vals_neg, 'val')
    xml, n4 = replace_series(xml, 'positive pre', new_pos_pre, 'xVal')
    xml, n5 = replace_series(xml, 'Negative Pre', new_neg_pre, 'xVal')
    print(f"\nReplaced: pos {n1}, neu {n2}, neg {n3}, pos pre {n4}, neg pre {n5}")

    # Strip manual data-label text overrides (they go stale)
    # Find <c:dLbl> containing <a:t>X%</a:t> and remove just the <c:tx>...</c:tx> sub-block
    def strip_manual_text(content):
        new = content
        while True:
            m = re.search(
                r'<c:dLbl>(?:(?!</c:dLbl>).)*?<a:t>\d+%</a:t>(?:(?!</c:dLbl>).)*?</c:dLbl>',
                new, re.DOTALL)
            if not m: break
            block = m.group(0)
            new_block = re.sub(r'<c:tx>.*?</c:tx>\s*', '', block, count=1, flags=re.DOTALL)
            if new_block == block: break
            new = new.replace(block, new_block, 1)
        return new
    before_overrides = len(re.findall(r'<a:t>\d+%</a:t>', xml))
    xml = strip_manual_text(xml)
    after_overrides = len(re.findall(r'<a:t>\d+%</a:t>', xml))
    print(f"Stripped {before_overrides - after_overrides} manual label overrides")

    with open(args.chart_xml, 'w') as f:
        f.write(xml)
    print(f"Updated {args.chart_xml}")

if __name__ == '__main__':
    main()
