"""Puma August 2026 content pack. Basis = July deck's: non-car noise out, WRC rally coverage IN, passing
news/blog mentions tagged pos/neg -> neutral (Tier 1). Prior month = the PUBLISHED July chart values."""
import re, json, pandas as pd
from openpyxl import load_workbook
import sys
W = sys.argv[1].rstrip('/') + '/'   # work dir with <Carline>_<Mon>_raw_CLEAN.xlsx; PUB below = the last deck's published chart values (read from its chart XML)
Q = 'Query Name'; C = 'Ford Puma'
SER = {'Puma': 'Ford Puma', 'Yaris Cross': 'Toyota Yaris Cross', '2008': 'Peugeot 2008', 'Kona': 'Hyundai Kona', 'T-Cross': 'VW T-Cross'}
CATS = ['Hyundai Kona', 'Peugeot 2008', 'Toyota Yaris Cross', 'VW T-Cross', 'Ford Puma']
PLAT = ['Online News', 'Forums', 'X', 'Blogs', 'Review']; CTRY = [('UK', 'United Kingdom'), ('ES', 'Spain'), ('IT', 'Italy'), ('FR', 'France'), ('DE', 'Germany')]
PUB = {'vol': {'Puma': [1005, 836, 649, 949], 'Yaris Cross': [1008, 517, 429, 718], '2008': [712, 774, 497, 1261], 'Kona': [377, 347, 398, 510], 'T-Cross': [194, 214, 178, 333]},
       'pos': [5.6862745098039215E-2, 6.5820777160983349E-2, 9.1922005571030641E-2, 4.5045045045045043E-2, 9.799789251844046E-2],
       'neg': [5.2941176470588235E-2, 0.17049960348929422, 7.5208913649025072E-2, 3.003003003003003E-2, 9.2729188619599584E-2],
       'puma_pos_n': 93, 'puma_neg_n': 88, 'puma_n': 949}

wb = load_workbook(W + sys.argv[2], read_only=True)   # cleaned workbook filename
def sheet(name):
    ws = wb[name]; rows = ws.iter_rows(values_only=True); hdr = list(next(rows)); return pd.DataFrame(list(rows), columns=hdr)
clean = sheet('Clean'); excl = sheet('Excluded').rename(columns={'Exclusion reason': 'why'})
wrc = excl[excl.why.str.startswith('WRC')].drop(columns=['why'])
df = pd.concat([clean, wrc], ignore_index=True)
for c in ['Title', 'Full Text', 'Url', 'Domain', 'Country', 'Page Type', 'Page Type Name', 'Sentiment']: df[c] = df[c].fillna('').astype(str)
df['Date'] = pd.to_datetime(df['Date'], errors='coerce'); df['Month'] = (df['Date'] + pd.Timedelta(hours=1)).dt.month
df['txt'] = (df.Title + ' ' + df['Full Text']).str.replace(r'\s+', ' ', regex=True)
df['is_wrc'] = df.txt.str.contains(r'rally1|m-sport|\bwrc\b|rallye? (de |du |di )|rally (paraguay|finland|estonia|finlandia)', case=False, regex=True)

# ---- Tier 1 retag on Puma August: passing news/blog mentions tagged pos/neg -> neutral; obvious polarity errors ----
aug = (df.Month == 8) & (df[Q] == C); s = df['Sentiment'].copy()
news = df['Page Type'].isin(['news', 'blog'])
PASSING = (r'Kia EV2|Mazda CX-3|BYD supera|auto elettriche cinesi|Danubio|insurance|fine alert|speed camera|Lost Android|Yaris Cross ofrece|Leapmotor T03|'
           r'Fiat Panda Elettra|Wunderlich|Fiat 600|Tourneo Courier|Northern Ireland|Campagne di veiculi|Used car UK dealers|Jeep Avenger|Ioniq 3|'
           r'Chevrolet|Geely E2|ID\. Cross|Corvette|Mustang|Stonic|Epiq|Peaq|Fathom|Sea-Doo|Pirelli|Best Chinese cars|Toyota Yaris Cross|Mazda nos col|'
           r'Volkswagen Taigo|crossover compatti|Crossover compatti|autonomia|rottamazione|sconti|Ladri|Fuoristrada sulla duna|informazione.news|ID\.Cross|PUMA x Porsche')
mask_pass = aug & news & df.txt.str.contains(PASSING, case=False, regex=True) & ~df.Title.str.contains(r'Puma', case=False) & s.isin(['positive', 'negative'])
n_pass = int(mask_pass.sum()); s[mask_pass] = 'neutral'
# rally results tagged negative/positive are race outcomes, not brand sentiment -> neutral (July precedent: "rally coverage kept", sentiment fragmented)
mask_wrc = aug & df.is_wrc & s.isin(['positive', 'negative']); n_wrc = int(mask_wrc.sum()); s[mask_wrc] = 'neutral'
# polarity fixes
fix_pos = aug & df.txt.str.contains(r'El Ford Puma es muy chulo|super zufrieden|proud owner of a Ford Puma|Ford puma nuevo y vacaciones', case=False) & (s != 'positive'); s[fix_pos] = 'positive'
df['Sentiment'] = s
print(f"Tier 1 retags: passing news pos/neg->neutral {n_pass}, WRC pos/neg->neutral {n_wrc}, ->positive {int(fix_pos.sum())}")

m = df[df.Month == 8]; p = m[m[Q] == C]
P = {'n': int(len(p)), 'volume': {k: int(m[m[Q] == q].Url.nunique()) for k, q in SER.items()},
     'platform': {k: int((p['Page Type Name'] == k).sum()) for k in PLAT},
     'country': {k: int((p.Country == f).sum()) for k, f in CTRY}, 'country_other': int((~p.Country.isin([f for _, f in CTRY])).sum()),
     'country_news': {k: int(((p.Country == f) & (p['Page Type'] == 'news')).sum()) for k, f in CTRY},
     'sent': {}, 'wrc_rows': int(p.is_wrc.sum()), 'gen_e_rows': int(p.txt.str.contains(r'gen-?e\b', case=False).sum())}
for c in CATS:
    x = m[m[Q] == c]; n = len(x); vc = x.Sentiment.value_counts()
    P['sent'][c] = {'n': int(n), 'pos': int(vc.get('positive', 0)), 'neg': int(vc.get('negative', 0)), 'pos_pct': vc.get('positive', 0) / n, 'neg_pct': vc.get('negative', 0) / n}
P['puma_sent_platform'] = {k: p[p.Sentiment == k]['Page Type Name'].value_counts().to_dict() for k in ('positive', 'negative')}
P['vol_delta'] = {k: P['volume'][k] - PUB['vol'][k][3] for k in SER}
def norm(t): return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', str(t).lower())).strip()[:80]
nb = p[p['Page Type'].isin(['news', 'blog'])].assign(k=lambda d: d.Title.map(norm))
P['news_clusters'] = [{'n': int(n), 'title': nb[nb.k == k].Title.iloc[0][:110], 'url': nb[nb.k == k].Url.iloc[0], 'country': nb[nb.k == k].Country.value_counts().head(2).to_dict(),
                       'sent': nb[nb.k == k].Sentiment.value_counts().to_dict()} for k, n in nb.k.value_counts().head(20).items()]
P['puma_negatives'] = [{'d': str(r.Date.date()), 'pt': r['Page Type'], 'dom': r.Domain, 'title': r.Title[:80], 'txt': r['Full Text'][:120]} for _, r in p[p.Sentiment == 'negative'].iterrows()]
P['puma_positives'] = [{'d': str(r.Date.date()), 'pt': r['Page Type'], 'dom': r.Domain, 'title': r.Title[:80], 'txt': r['Full Text'][:120]} for _, r in p[p.Sentiment == 'positive'].iterrows()]
# competitor detail
P['comp'] = {}
for c in CATS[:-1]:
    x = m[m[Q] == c]; xn = x[x['Page Type'].isin(['news', 'blog'])].assign(k=lambda d: d.Title.map(norm))
    P['comp'][c] = {'platforms': x['Page Type Name'].value_counts().to_dict(), 'neg_platforms': x[x.Sentiment == 'negative']['Page Type Name'].value_counts().to_dict(),
                    'top': [{'n': int(n), 'title': xn[xn.k == k].Title.iloc[0][:100], 'url': xn[xn.k == k].Url.iloc[0], 'sent': xn[xn.k == k].Sentiment.value_counts().to_dict()} for k, n in xn.k.value_counts().head(4).items()]}
json.dump({'pub': PUB, 'aug': P}, open(W + 'content_pack.json', 'w'), indent=1, default=str)
print(json.dumps({k: v for k, v in P.items() if k in ('n', 'volume', 'platform', 'country', 'country_other', 'country_news', 'vol_delta', 'wrc_rows', 'gen_e_rows', 'puma_sent_platform')}, indent=0))
for c, x in P['sent'].items(): print(f"  {c:20s} n={x['n']} pos {x['pos']} ({x['pos_pct']*100:.1f}%) neg {x['neg']} ({x['neg_pct']*100:.1f}%)")
print("\nPuma negatives after retag:"); [print(f"   {r['d']} {r['pt']:7s} {r['dom'][:20]:20s} | {r['title'][:60]} || {r['txt'][:90]}") for r in P['puma_negatives']]
print("\nPuma positives after retag:"); [print(f"   {r['d']} {r['pt']:7s} {r['dom'][:20]:20s} | {r['title'][:60]} || {r['txt'][:90]}") for r in P['puma_positives']]
