"""Chart inputs for Capri July + August 2026 slides, from the cleaned dataset (local-time month bucketing,
volume = unique URLs), with sentiment under three retag tiers."""
import os, sys, re, pandas as pd
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))  # tools/bw_read
from bw_read import load_dataframe
from openpyxl import load_workbook

CLEAN = sys.argv[1]  # cleaned workbook (Clean sheet)
OUT = sys.argv[2]    # markdown out
Q = 'Query Name'; C = 'Ford Capri'
ORDER = ['Ford Capri', 'Peugeot E3008', 'VW ID.5', 'Skoda Enyaq Coupe', 'Polestar 2', 'Audi Q4 Sportback e-tron', 'Kia EV6']
LABEL = {'Peugeot E3008': 'Peugeot E-3008', 'Skoda Enyaq Coupe': 'Skoda Enyaq Coupe', 'Audi Q4 Sportback e-tron': 'Audi Q4 Sportback e-tron'}

# read Clean sheet (already noise-free) -> DataFrame
wb = load_workbook(CLEAN, read_only=True); ws = wb['Clean']
rows = ws.iter_rows(values_only=True); header = next(rows)
df = pd.DataFrame(list(rows), columns=header)
for c in ['Title', 'Full Text', 'Url', 'Domain', 'Country', 'Page Type', 'Page Type Name', 'Sentiment']:
    df[c] = df[c].fillna('').astype(str)
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df['Month'] = (df['Date'] + pd.Timedelta(hours=1)).dt.month  # local-time bucketing, as Brandwatch labels months
df['txt'] = (df.Title + ' ' + df['Full Text']).str.replace(r'\s+', ' ', regex=True)

# ---- retag tiers (Capri only; from CAPRI_AUG2026_FLAGS.md) ----
def apply_retags(d, tier):
    d = d.copy(); s = d['Sentiment'].copy()
    cap = d[Q] == C
    aug, jul = d.Month == 8, d.Month == 7
    T = lambda pat: cap & d.txt.str.contains(pat, case=False, regex=True)
    if tier >= 1:  # must-fix
        s[T(r'electric skip|besmirch|blob|sucks harder|a mustang that isn.t a mustang') & aug & (s == 'positive')] = 'negative'
        s[T(r'Audi A2|geen kuren') & aug & (s == 'positive')] = 'neutral'
        s[T(r'burlarse de los conductores de un AMI') & aug & (s == 'negative')] = 'neutral'
        # July equivalents (only matter for the July slide bars)
        s[T(r'better than the new|monstrosity|offensive to the heritage|devoid of any style') & jul & (s == 'positive')] = 'negative'
        s[T(r'Nani Roma|Marc Roma|Roma y estrena|like an SUV like the Ford Capri') & jul & (s == 'positive')] = 'neutral'
        s[T(r'no pasa desapercibido|joyless\? The new Ford Capri suggests we were wrong|dynamische energieprijzen|Explorer/Capri are due to get|Even in my Ford Capri EV|luchtweerstand|asentamiento') & jul & (s == 'negative')] = 'neutral'
    if tier >= 2:  # judgement
        s[T(r'Sacrilege|imposter|Grifters|it is not a Capri. Its an insult|call it something else|Def not the new one|only has the badge in common|really looks like the old version Capri 😂|should have been the new capri from day one|Totally out of place') & aug & (s == 'neutral')] = 'negative'
        s[T(r'Saw the new capri recently 😍|This is the car I.m after|Esperienza di guida molto buona|elettrica confortevole, dinamica ed equilibrata') & aug & (s == 'neutral')] = 'positive'
        s[T(r'verdoppelt|most comfortable ride for our 6ft') & jul & (s == 'neutral')] = 'positive'
    d['Sentiment'] = s
    return d

L = []; w = L.append
w('# Capri — chart inputs for the July and August 2026 slides (from the cleaned dataset)\n')
w('Source: `Capri_Aug2026_raw_CLEAN.xlsx` → Clean sheet (noise excluded). Months bucketed on the local-time boundary, as '
  'Brandwatch labels them (a 23:xx UTC mention on the last day of a month is the next month). Volume = COUNTUNIQUE(URL). '
  'These replace the four Brandwatch chart exports for July and August; May and June come from the June deck (carry-forward).\n')

# 1 volume
w('## 1. Volume (line chart) — July and August columns\n')
w('| Model | July | August | Δ | Δ% |\n|---|---|---|---|---|')
for m in ORDER:
    j = df[(df[Q] == m) & (df.Month == 7)].Url.nunique(); a = df[(df[Q] == m) & (df.Month == 8)].Url.nunique()
    w(f'| {LABEL.get(m, m)} | {j} | {a} | {a-j:+d} | {(a-j)/j*100 if j else 0:+.1f}% |')
w('\nMay and June: copy from the June deck volume chart (`chart1`/`chart9` backing xlsx). The Average row is a formula — do not overwrite.\n')

# 2 country / 3 platform per month
for mon, name in ((7, 'July'), (8, 'August')):
    c = df[(df[Q] == C) & (df.Month == mon)]
    w(f'## 2{"a" if mon == 7 else "b"}. Capri country pie — {name} (n={len(c)})\n')
    w('| Country | Mentions | Share |\n|---|---|---|')
    for k, v in c.Country.replace('', '(unknown)').value_counts().items():
        w(f'| {k} | {v} | {v/len(c)*100:.1f}% |')
    w('')
    w(f'## 3{"a" if mon == 7 else "b"}. Capri platform pie — {name} (n={len(c)})\n')
    w('| Platform | Mentions | Share |\n|---|---|---|')
    for k, v in c['Page Type Name'].value_counts().items():
        w(f'| {k} | {v} | {v/len(c)*100:.1f}% |')
    w('')

# 4 sentiment under tiers
w('## 4. Sentiment — bars (reporting month) and diamonds (prior month)\n')
w('Three tiers. **Tier 0** = as exported. **Tier 1** = must-fix retags (wrong polarity, passing mentions → neutral). '
  '**Tier 2** = Tier 1 + judgement retags (clearly hostile/favourable neutral-tagged items). Competitors are unchanged across tiers '
  '(no retag review done on them). Pick one tier and use it for BOTH months so the diamonds and bars are like-for-like.\n')
for tier in (0, 1, 2):
    d = apply_retags(df, tier)
    w(f'### Tier {tier}\n')
    w('| Model | Month | n | pos | pos % | neu | neu % | neg | neg % | diamond: positive pre = pos % | diamond: Negative Pre = 1 − neg % |')
    w('|---|---|---|---|---|---|---|---|---|---|---|')
    for m in ORDER:
        for mon, name in ((7, 'July'), (8, 'August')):
            s = d[(d[Q] == m) & (d.Month == mon)]; n = len(s); vc = s.Sentiment.value_counts()
            p, ng, nu = vc.get('positive', 0), vc.get('negative', 0), vc.get('neutral', 0)
            w(f'| {LABEL.get(m, m)} | {name} | {n} | {p} | {p/n*100 if n else 0:.1f}% | {nu} | {nu/n*100 if n else 0:.1f}% | {ng} | {ng/n*100 if n else 0:.1f}% | {p/n if n else 0:.3f} | {1-ng/n if n else 1:.3f} |')
    w('')
w('For the **July slide**: bars = July row; diamonds = **June, from the June deck\'s published sentiment chart** (`chart3.xml`), not from data.  ')
w('For the **August slide**: bars = August row; diamonds = July row of the same tier (this is the "diamonds from BW post-denoise" case — there is no published July chart).\n')

# 5 caveats
w('## 5. Caveats to state in the notes\n')
w('- July and August numbers come from the cleaned raw export, not the Brandwatch dashboard. Unless the same exclusions and retags are '
  'applied in Brandwatch, the dashboard will show higher Capri volume (117/110 raw vs the cleaned figures above).')
w('- May and June (volume line) were built from Brandwatch chart exports before the boolean was cleaned, so the line has a small '
  'methodology step between June and July. Capri noise was ~10% of rows in July and ~20% in August; the step is a few mentions on the competitors and up to ~10 on Capri.')
w('- Competitor sentiment is as exported (no retag review); competitor noise rows (Audi other-model, Peugeot/ID.5 no-model) are excluded.')
w('- Skoda Enyaq Coupe has 0 positive and 0 negative in August under every tier — the bar is 100% neutral by the data.')
open(OUT, 'w').write('\n'.join(L)); print('wrote', OUT, len(L), 'lines')
