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

def _find_prior_sentiment_chart(z):
    """Locate the reported-month slide's sentiment chart in a prior deck, BY ROLE.

    The prior deck's chart numbering varies per carline/cycle (Capri=chart4, Puma=chart16,
    Explorer=chart3), so never hardcode a number. Pick the newest month-labelled, chart-bearing
    slide, then the chart whose series are positive/neutral/negative (a bar+scatter combo).
    """
    import re as _re
    months = {m.lower(): i for i, m in enumerate(
        ["","January","February","March","April","May","June","July","August",
         "September","October","November","December"]) if m}
    best = None
    for n in z.namelist():
        m = _re.match(r'ppt/slides/(slide\d+)\.xml$', n)
        if not m:
            continue
        rels_name = f'ppt/slides/_rels/{m.group(1)}.xml.rels'
        if rels_name not in z.namelist():
            continue
        charts = _re.findall(r'Target="\.\./charts/(chart\d+\.xml)"', z.read(rels_name).decode())
        if not charts:
            continue
        txt = z.read(n).decode('utf-8', 'replace')
        mm = _re.search(r'\(EU\d\s+([A-Za-z]+)\s+(20\d\d)\)', txt) or _re.search(r'([A-Za-z]+)\s+(20\d\d)', txt)
        if not mm:
            continue
        mo = months.get(mm.group(1).lower())
        if not mo:
            continue
        key = (int(mm.group(2)), mo)
        if best is None or key > best[0]:
            best = (key, charts)
    if not best:
        return None
    for c in best[1]:
        xml = z.read(f'ppt/charts/{c}').decode()
        sn = [s.lower() for s in re.findall(r'<c:tx>.*?<c:v>([^<]*)</c:v>', xml, re.DOTALL)]
        if 'scatterChart' in xml or any(s in ('positive', 'negative', 'neutral') for s in sn):
            return f'ppt/charts/{c}'
    return None

def read_prior_chart3(pptx_path, prior_chart=None):
    """Read the prior month's sentiment chart — positive and negative series values.

    Chart is located by role (see _find_prior_sentiment_chart); pass prior_chart to override.
    """
    with zipfile.ZipFile(pptx_path) as z:
        target = prior_chart or _find_prior_sentiment_chart(z) or 'ppt/charts/chart3.xml'
        xml = z.read(target).decode()
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
    ap.add_argument('--prior-chart', default=None,
                    help="explicit ppt/charts/chartN.xml inside the prior deck (else auto-detect by role)")
    ap.add_argument('--prior-month-name', default=None,
                    help="prior month name for month-labelled diamond series, e.g. June (else leave name)")
    args = ap.parse_args()

    new_sent = read_sentiment_export(args.new_month_sent)
    print(f"New sentiment: {new_sent}")
    prior_chart3 = read_prior_chart3(args.prior_from_last_month, args.prior_chart)
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

    # Diamonds = the scatter series carrying <c:xVal>. Their names vary per deck:
    # 'positive pre'/'Negative Pre' (Explorer/Capri) or month-labelled 'Positive May' (Puma).
    # Classify by polarity word in the name; update the xVals; and if the name ends in a month,
    # rename it to the new prior month.
    _MN = ['january','february','march','april','may','june','july','august',
           'september','october','november','december']
    def update_diamonds(content, prior_name):
        counts = {'pos': 0, 'neg': 0}
        def repl(mo):
            block = mo.group(0)
            if '<c:xVal>' not in block:
                return block
            nm = re.search(r'<c:tx>.*?<c:v>([^<]*)</c:v>', block, re.DOTALL)
            name = nm.group(1) if nm else ''
            low = name.lower()
            if 'positive' in low:
                new_list, key = new_pos_pre, 'pos'
            elif 'negative' in low:
                new_list, key = new_neg_pre, 'neg'
            else:
                return block
            xv = re.search(r'<c:xVal>(.*?)</c:xVal>', block, re.DOTALL)
            if xv:
                vb = xv.group(0); olds = re.findall(r'<c:v>([^<]+)</c:v>', vb); nvb = vb
                for o, new_v in zip(olds, new_list):
                    nvb = nvb.replace(f'<c:v>{o}</c:v>', f'<c:v>{new_v:.17g}</c:v>', 1)
                block = block.replace(vb, nvb, 1); counts[key] += 1
            if prior_name:
                m2 = re.match(r'(Positive|Negative|positive|negative)\s+(\w+)$', name)
                if m2 and m2.group(2).lower() in _MN:
                    block = block.replace(f'<c:v>{name}</c:v>', f'<c:v>{m2.group(1)} {prior_name}</c:v>')
            return block
        content = re.sub(r'<c:ser>.*?</c:ser>', repl, content, flags=re.DOTALL)
        return content, counts['pos'], counts['neg']
    xml, n4, n5 = update_diamonds(xml, args.prior_month_name)
    print(f"\nReplaced: pos {n1}, neu {n2}, neg {n3}, pos-diamond {n4}, neg-diamond {n5}")

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
