#!/usr/bin/env python3
"""
Update the sentiment combo chart (stacked bars + diamond scatter).

Bars = new month's pos/neu/neg %s.
Diamonds = prior month's pos % (positive pre) and 1 - prior month's neg % (Negative Pre).

The PRIOR month diamond values come from the PRIOR month's published chart (the sentiment chart on slide 2,
inside last_month.pptx), NOT from the new raw data. Last month's published values are
what the client saw; using newly-retagged numbers creates a mismatch with the
comparison slide.

Also strips any manual <a:t>X%</a:t> data-label overrides — they go stale.

Usage:
    python3 update_sentiment_chart.py path/to/chart11.xml \\
        --new-month-sent chart_sentiment.xlsx \\
        --prior-from-last-month last_month.pptx
"""
import argparse, os, re, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from chart_cache import set_cache, cache_values, norm

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

def read_prior_sentiment(pptx_path):
    """Last month's published sentiment bars as {category: (pos, neg)}.

    The chart is found from the deck's slide 2 (last month's slide), as the chart
    there with 'positive'/'negative' bar series - its chart number differs by
    carline template (chart3 in Explorer, chart4 in Capri)."""
    with zipfile.ZipFile(pptx_path) as z:
        rels = z.read('ppt/slides/_rels/slide2.xml.rels').decode()
        charts = re.findall(r'Target="\.\./charts/(chart\d+\.xml)"', rels)
        for name in charts:
            xml = z.read(f'ppt/charts/{name}').decode()
            sers = {}
            for block in re.findall(r'<c:ser>.*?</c:ser>', xml, re.DOTALL):
                nm = re.search(r'<c:tx>.*?<c:v>([^<]+)</c:v>', block, re.DOTALL)
                if nm:
                    sers[nm.group(1).strip()] = block
            if 'positive' in sers and 'negative' in sers:
                cats = re.findall(r'<c:v>([^<]+)</c:v>',
                                  re.search(r'<c:cat>(.*?)</c:cat>', sers['positive'], re.DOTALL).group(1))
                pos = [float(v) for v in cache_values(sers['positive'], 'c:val')]
                neg = [float(v) for v in cache_values(sers['negative'], 'c:val')]
                print(f"Prior month's sentiment chart: {name} (slide 2)")
                return {c: (p, n) for c, p, n in zip(cats, pos, neg)}
    sys.exit(f"No sentiment chart (positive/negative series) on slide 2 of {pptx_path}")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('chart_xml')
    ap.add_argument('--new-month-sent', required=True)
    ap.add_argument('--prior-from-last-month', required=True)
    args = ap.parse_args()

    new_sent = read_sentiment_export(args.new_month_sent)
    print(f"New sentiment: {new_sent}")
    prior = {norm(k): v for k, v in read_prior_sentiment(args.prior_from_last_month).items()}

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
            if norm(q) == norm(chart_name): return q
        for q in queries:
            if norm(chart_name) in norm(q) or norm(q) in norm(chart_name): return q
        return None

    new_vals_pos, new_vals_neu, new_vals_neg = [], [], []
    for cn in chart_cats:
        q = find_export_match(cn, list(new_sent.keys()))
        if not q:
            sys.exit(f"No export query matches chart category '{cn}' - exports: {list(new_sent)}")
        c = new_sent[q]
        tot = c['pos']+c['neu']+c['neg']
        if tot == 0:
            new_vals_pos.append(0); new_vals_neu.append(0); new_vals_neg.append(0)
        else:
            new_vals_pos.append(c['pos']/tot)
            new_vals_neu.append(c['neu']/tot)
            new_vals_neg.append(c['neg']/tot)

    # Diamonds = last month's published bars, matched by category name
    missing = [cn for cn in chart_cats if norm(cn) not in prior]
    if missing:
        sys.exit(f"Prior month's chart has no category for {missing}")
    new_pos_pre = [prior[norm(cn)][0] for cn in chart_cats]
    new_neg_pre = [1 - prior[norm(cn)][1] for cn in chart_cats]

    print(f"\nNew bar values:")
    for i, cn in enumerate(chart_cats):
        print(f"  {cn}: pos {new_vals_pos[i]*100:.2f}% / neu {new_vals_neu[i]*100:.2f}% / neg {new_vals_neg[i]*100:.2f}%")
    print(f"\nDiamond X (last month's published bars):")
    for i, cn in enumerate(chart_cats):
        print(f"  {cn}: pos {new_pos_pre[i]*100:.2f}% / 1-neg {new_neg_pre[i]*100:.2f}%")

    # Replace each series's cached values by point index
    def replace_series(content, anchor, new_list, kind='val'):
        tag = 'c:xVal' if kind == 'xVal' else 'c:val'
        for block in re.findall(r'<c:ser>.*?</c:ser>', content, re.DOTALL):
            nm = re.search(r'<c:tx>.*?<c:v>([^<]+)</c:v>', block, re.DOTALL)
            if nm and nm.group(1).strip().lower() == anchor.lower():  # Kuga: 'Positive Pre'
                return content.replace(block, set_cache(block, tag, [f'{v:.17g}' for v in new_list]), 1), len(new_list)
        return content, 0

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
