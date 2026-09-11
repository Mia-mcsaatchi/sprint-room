"""Tier-1 content pack for the Capri July + August 2026 slides: every number, cluster and URL the
charts, bullets and notes need, computed from the cleaned dataset. Writes JSON + a readable markdown."""
import sys, re, json, pandas as pd
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))  # tools/bw_read
from openpyxl import load_workbook

CLEAN = sys.argv[1]  # cleaned workbook (Clean sheet)
OUTJ = sys.argv[2]   # content_pack.json
OUTM = sys.argv[3]   # readable markdown
Q = 'Query Name'; C = 'Ford Capri'
# chart category order (sentiment chart4) and line-series names (chart1) from the June deck
SENT_ORDER = ['Audi Q4 Sportback e-tron', 'Kia EV6', 'Peugeot E3008', 'Polestar 2', 'Skoda Enyaq Coupe', 'VW ID.5', 'Ford Capri']
SENT_CATS = ['Audi Q4 Sportback e-tron', 'Kia EV6', 'Peugeot E 3008', 'Polestar 2', 'Skoda Enyaq Coupe', 'VW ID.5', 'Ford Capri']
LINE_SERIES = {'Polestar 2': 'Polestar 2', 'E-3008': 'Peugeot E3008', 'EV6': 'Kia EV6', 'ID.5': 'VW ID.5',
               'Q4 Sportback e-tron': 'Audi Q4 Sportback e-tron', 'Capri': 'Ford Capri', 'Enyaq Coupe': 'Skoda Enyaq Coupe'}
# June deck published values (chart1 Mar..Jun; chart4 pos/neg in SENT_CATS order)
JUNE_VOL = {'Polestar 2': [489, 419, 415, 306], 'E-3008': [326, 354, 280, 243], 'EV6': [203, 213, 213, 155], 'ID.5': [102, 60, 82, 20],
            'Q4 Sportback e-tron': [75, 322, 152, 102], 'Capri': [198, 124, 130, 130], 'Enyaq Coupe': [121, 33, 32, 31]}
JUNE_POS = [0.058823529411764705, 0.16774193548387098, 0.069958847736625515, 0.10784313725490197, 0.032258064516129031, 0.1, 0.1]
JUNE_NEG = [0, 0.10967741935483871, 0.037037037037037035, 0.11437908496732026, 0.12903225806451613, 0, 0.061538461538461542]
JUNE_N = {'Audi Q4 Sportback e-tron': 102, 'Kia EV6': 155, 'Peugeot E 3008': 243, 'Polestar 2': 306, 'Skoda Enyaq Coupe': 31, 'VW ID.5': 20, 'Ford Capri': 130}
COUNTRY_CODES = [('UK', 'United Kingdom'), ('IT', 'Italy'), ('FR', 'France'), ('DE', 'Germany'), ('ES', 'Spain'), ('NL', 'Netherlands'), ('NO', 'Norway')]
PLATFORMS = ['Online News', 'X', 'Forums', 'Blogs', 'Review', 'Tumblr']

wb = load_workbook(CLEAN, read_only=True); ws = wb['Clean']; rows = ws.iter_rows(values_only=True); header = next(rows)
df = pd.DataFrame(list(rows), columns=header)
for c in ['Title', 'Full Text', 'Url', 'Domain', 'Country', 'Page Type', 'Page Type Name', 'Sentiment']:
    df[c] = df[c].fillna('').astype(str)
df['Date'] = pd.to_datetime(df['Date'], errors='coerce'); df['Month'] = (df['Date'] + pd.Timedelta(hours=1)).dt.month
df['txt'] = (df.Title + ' ' + df['Full Text']).str.replace(r'\s+', ' ', regex=True)

# ---- Tier 1 retags (Capri) ----
s = df['Sentiment'].copy(); cap = df[Q] == C; aug, jul = df.Month == 8, df.Month == 7
T = lambda pat: cap & df.txt.str.contains(pat, case=False, regex=True)
s[T(r'electric skip|besmirch|blob|sucks harder|a mustang that isn.t a mustang') & aug & (s == 'positive')] = 'negative'
s[T(r'Audi A2|geen kuren') & aug & (s == 'positive')] = 'neutral'
s[T(r'burlarse de los conductores de un AMI') & aug & (s == 'negative')] = 'neutral'
s[T(r'better than the new|monstrosity|offensive to the heritage|devoid of any style') & jul & (s == 'positive')] = 'negative'
s[T(r'Nani Roma|Marc Roma|Roma y estrena|like an SUV like the Ford Capri') & jul & (s == 'positive')] = 'neutral'
s[T(r'no pasa desapercibido|joyless\? The new Ford Capri suggests we were wrong|dynamische energieprijzen|Explorer/Capri are due to get|Even in my Ford Capri EV|luchtweerstand|asentamiento') & jul & (s == 'negative')] = 'neutral'
df['Sentiment'] = s

def norm(t): return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', str(t).lower())).strip()[:80]
pack = {'months': {}, 'june': {'vol': JUNE_VOL, 'pos': JUNE_POS, 'neg': JUNE_NEG, 'n': JUNE_N}}
md = ['# Capri content pack — Tier 1 (must-fix retags), cleaned data, local-time months\n']
for mon, name, prev in ((7, 'July', None), (8, 'August', 7)):
    m = df[df.Month == mon]; c = m[m[Q] == C]
    P = {'name': name, 'n': int(len(c))}
    P['volume'] = {ser: int(m[m[Q] == q].Url.nunique()) for ser, q in LINE_SERIES.items()}
    P['platform'] = {p: int((c['Page Type Name'] == p).sum()) for p in PLATFORMS}
    P['country'] = {code: int((c.Country == full).sum()) for code, full in COUNTRY_CODES}
    P['country_news'] = {code: int(((c.Country == full) & (c['Page Type'] == 'news')).sum()) for code, full in COUNTRY_CODES}
    P['country_other'] = int(len(c) - sum(P['country'].values()))
    sent = {}
    for cat, q in zip(SENT_CATS, SENT_ORDER):
        x = m[m[Q] == q]; n = len(x); vc = x.Sentiment.value_counts()
        sent[cat] = {'n': int(n), 'pos': int(vc.get('positive', 0)), 'neu': int(vc.get('neutral', 0)), 'neg': int(vc.get('negative', 0)),
                     'pos_pct': vc.get('positive', 0) / n if n else 0, 'neg_pct': vc.get('negative', 0) / n if n else 0}
    P['sentiment'] = sent
    P['capri_sent_by_platform'] = {snt: c[c.Sentiment == snt]['Page Type Name'].value_counts().to_dict() for snt in ('positive', 'negative')}
    P['capri_platform_by_country'] = {p: c[c['Page Type Name'] == p].Country.value_counts().head(5).to_dict() for p in ('Online News', 'X', 'Forums')}
    # news clusters with URL (>=2) + carline-titled singles
    nb = c[c['Page Type'].isin(['news', 'blog'])].assign(k=lambda d: d.Title.map(norm))
    clusters = []
    for k, cnt in nb.k.value_counts().items():
        g = nb[nb.k == k]
        clusters.append({'n': int(cnt), 'title': g.Title.iloc[0][:120], 'domains': g.Domain.value_counts().head(4).to_dict(),
                         'countries': g.Country.value_counts().head(3).to_dict(), 'sentiment': g.Sentiment.value_counts().to_dict(),
                         'urls': g.Url.head(4).tolist(), 'capri_titled': bool(re.search('capri', g.Title.iloc[0], re.I))})
    P['news_clusters'] = clusters[:45]
    # X: backlash + threads
    x = c[c['Page Type'] == 'twitter']
    P['x'] = {'n': int(len(x)), 'uk': int((x.Country == 'United Kingdom').sum()),
              'backlash_terms': int(x.txt.str.contains(r'ugly|monstrosity|shite|horrid|disaster|sacrileg|abomination|insult|besmirch|travesty|blob|should have called|not a capri|proper one|so called|imposter|skip', case=False).sum()),
              'bolderston': int(x.txt.str.contains('bolderston', case=False).sum()),
              'adrianfclarke_insidelucys_ctrlaltrefuel': int(x.txt.str.contains('adrianfclarke|InsideLucysHead|ctrlaltrefuel', case=False).sum()),
              'neg': int((x.Sentiment == 'negative').sum()), 'pos': int((x.Sentiment == 'positive').sum())}
    f = c[c['Page Type'] == 'forum']
    P['forums'] = {'n': int(len(f)), 'domains': f.Domain.value_counts().head(6).to_dict(), 'threads': f.Title.str[:60].value_counts().head(6).to_dict()}
    # competitor movers: biggest |Δ| vs prior month volume
    prev_vol = {ser: JUNE_VOL[ser][3] for ser in LINE_SERIES} if prev is None else pack['months']['7']['volume']
    P['vol_delta'] = {ser: P['volume'][ser] - prev_vol[ser] for ser in LINE_SERIES}
    # competitor news/forum/x split for the biggest mover
    mover = max((k for k in LINE_SERIES if k != 'Capri'), key=lambda k: abs(P['vol_delta'][k]))
    mq = LINE_SERIES[mover]; mm = m[m[Q] == mq]
    mnb = mm[mm['Page Type'].isin(['news', 'blog'])].assign(k=lambda d: d.Title.map(norm))
    P['mover'] = {'series': mover, 'query': mq, 'delta': P['vol_delta'][mover], 'platforms': mm['Page Type Name'].value_counts().to_dict(),
                  'countries': mm.Country.value_counts().head(4).to_dict(),
                  'top_clusters': [{'n': int(cnt), 'title': mnb[mnb.k == k].Title.iloc[0][:110], 'url': mnb[mnb.k == k].Url.iloc[0],
                                    'sent': mnb[mnb.k == k].Sentiment.value_counts().to_dict()} for k, cnt in mnb.k.value_counts().head(5).items()]}
    if prev:
        pm = df[df.Month == prev]; pq = pm[pm[Q] == mq]
        P['mover']['prev_platforms'] = pq['Page Type Name'].value_counts().to_dict()
    pack['months'][str(mon)] = P

    # ---- markdown ----
    md.append(f'## {name}  (Capri n={len(c)})\n')
    md.append('**Volume (unique URLs):** ' + ', '.join(f'{k} {v} ({P["vol_delta"][k]:+d})' for k, v in P['volume'].items()))
    md.append('**Platform:** ' + str(P['platform']) + f"  |  **Country:** {P['country']} other={P['country_other']}  |  news: {P['country_news']}")
    md.append('**Sentiment (Tier 1):** ' + '; '.join(f"{k} n={v['n']} pos {v['pos']} ({v['pos_pct']*100:.1f}%) neg {v['neg']} ({v['neg_pct']*100:.1f}%)" for k, v in sent.items()))
    md.append(f"**Capri sentiment by platform:** {P['capri_sent_by_platform']}")
    md.append(f"**Capri X:** {P['x']}  |  **Forums:** {P['forums']}")
    md.append(f"**Biggest competitor mover:** {P['mover']['series']} {P['mover']['delta']:+d}; platforms {P['mover']['platforms']}; prev {P['mover'].get('prev_platforms')}")
    for cl in P['mover']['top_clusters']: md.append(f"   - [{cl['n']}] {cl['title']} | {cl['sent']} | {cl['url']}")
    md.append('**Capri news clusters:**')
    for cl in clusters[:30]:
        md.append(f"   - [{cl['n']}] {'★' if cl['capri_titled'] else ' '} {cl['title']} | {cl['countries']} | {cl['sentiment']} | {cl['urls'][0]}")
    md.append('')
json.dump(pack, open(OUTJ, 'w'), indent=1, default=str); open(OUTM, 'w').write('\n'.join(md)); print('wrote', OUTJ, OUTM)
