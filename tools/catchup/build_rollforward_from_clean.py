"""Standard roll-forward used for Puma Aug-2026: last deck (title, July, June) -> title, August (new), July (carried forward).
Reference implementation. Constants (series names, 5 sentiment categories, pie categories, diamond Y, sheet names) are the PUMA
template (Explorer-style: chart1 volume, chart2 country, chart3 platform, chart4 sentiment) - read them from the last deck first.
See tools/catchup/README.md."""
import os, re, shutil, subprocess, sys, json, copy
from pathlib import Path
from lxml import etree
from openpyxl import Workbook
import argparse
ap = argparse.ArgumentParser(description='Standard roll-forward from the last delivered deck + cleaned data (one new month slide). Puma Aug-2026 reference implementation.')
ap.add_argument('--work', required=True, help='work dir holding <last>_unpacked/, content_pack.json, slide_text.py')
ap.add_argument('--last-unpacked', default='july_unpacked'); ap.add_argument('--last-pptx', required=True)
ap.add_argument('--out', required=True); ap.add_argument('--report-date', required=True)
ap.add_argument('--pptx-scripts', default=os.environ.get('PPTX_SCRIPTS', '/mnt/skills/public/pptx/scripts'))
args = ap.parse_args()
WORK = Path(args.work); sys.path.insert(0, str(WORK))
from slide_text import AUGUST   # the new month's bullets/notes - client content, kept out of git
SRC = WORK / args.last_unpacked; BUILD = WORK / 'build'; U = BUILD / 'unpacked'
OUT = Path(args.out); LAST = Path(args.last_pptx); PPTX_SCRIPTS = Path(args.pptx_scripts)
PACK = json.load(open(WORK / 'content_pack.json')); PUB, AUG = PACK['pub'], PACK['aug']
REPORT_DATE = args.report_date
NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
C, A, P, R = ('{%s}' % NS[k] for k in ('c', 'a', 'p', 'r'))
CT = {'chart': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', 'colors': 'application/vnd.ms-office.chartcolorstyle+xml',
      'style': 'application/vnd.ms-office.chartstyle+xml', 'notes': 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'}
SER = ['Puma', 'Yaris Cross', '2008', 'Kona', 'T-Cross']; CATS = ['Hyundai Kona', 'Peugeot 2008', 'Toyota Yaris Cross', 'VW T-Cross', 'Ford Puma']
PLAT = ['Online News', 'Forums', 'X', 'Blogs', 'Review']; CTRY = ['UK', 'ES', 'IT', 'FR', 'DE']; DIAMOND_Y = [3, 9, 15, 21, 27]
MONTHS = ['May', 'June', 'July', 'August']

def read(p): return Path(p).read_text(encoding='utf-8')
def write(p, s): Path(p).write_text(s, encoding='utf-8')
def xml(p): return etree.parse(str(p), etree.XMLParser(remove_blank_text=False))
def save(t, p): t.write(str(p), xml_declaration=True, encoding='UTF-8', standalone=True)

# ---- 1. structure
if BUILD.exists(): shutil.rmtree(BUILD)
BUILD.mkdir(); shutil.copytree(SRC, U)
r = subprocess.run([sys.executable, str(PPTX_SCRIPTS / 'add_slide.py'), str(U), 'slide2.xml'], capture_output=True, text=True)
if r.returncode: sys.exit(r.stderr)
new = re.search(r'ppt/slides/(slide\d+)\.xml', r.stdout).group(1); print(r.stdout.strip().splitlines()[0])
pres = read(U / 'ppt/presentation.xml'); rels = read(U / 'ppt/_rels/presentation.xml.rels')
def rid_of(slide):
    m = re.search(rf'<Relationship[^>]*Target="slides/{slide}\.xml"[^>]*>', rels).group(0); return re.search(r'Id="(rId\d+)"', m).group(1)
ids = {v: k for k, v in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', pres)}
order = [(ids[rid_of(s)], rid_of(s)) for s in ('slide1', new, 'slide2')]
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', '<p:sldIdLst>' + ''.join(f'<p:sldId id="{i}" r:id="{r}"/>' for i, r in order) + '</p:sldIdLst>', pres, flags=re.S)
pres = re.sub(r'<p14:sldIdLst>.*?</p14:sldIdLst>', '<p14:sldIdLst>' + ''.join(f'<p14:sldId id="{i}"/>' for i, _ in order) + '</p14:sldIdLst>', pres, flags=re.S)
write(U / 'ppt/presentation.xml', pres)

ct = read(U / '[Content_Types].xml')
def add_override(part, ctype):
    global ct
    if f'PartName="{part}"' not in ct: ct = ct.replace('</Types>', f'<Override PartName="{part}" ContentType="{ctype}"/></Types>')
srels_p = U / f'ppt/slides/_rels/{new}.xml.rels'; srels = read(srels_p)
for i in range(1, 5):
    n = 8 + i
    for kind in ('chart', 'colors', 'style'): shutil.copy(U / f'ppt/charts/{kind}{i}.xml', U / f'ppt/charts/{kind}{n}.xml'); add_override(f'/ppt/charts/{kind}{n}.xml', CT[kind])
    crels = read(U / f'ppt/charts/_rels/chart{i}.xml.rels').replace(f'colors{i}.xml', f'colors{n}.xml').replace(f'style{i}.xml', f'style{n}.xml')
    crels = re.sub(r'Target="\.\./embeddings/[^"]+"', f'Target="../embeddings/Microsoft_Excel_Worksheet_c{n}.xlsx"', crels)
    write(U / f'ppt/charts/_rels/chart{n}.xml.rels', crels)
    srels = srels.replace(f'Target="../charts/chart{i}.xml"', f'Target="../charts/chart{n}.xml"')
# notes
shutil.copy(U / 'ppt/notesSlides/notesSlide1.xml', U / 'ppt/notesSlides/notesSlide3.xml')
write(U / 'ppt/notesSlides/_rels/notesSlide3.xml.rels', read(U / 'ppt/notesSlides/_rels/notesSlide1.xml.rels').replace('../slides/slide2.xml', f'../slides/{new}.xml'))
add_override('/ppt/notesSlides/notesSlide3.xml', CT['notes'])
nxt = max(int(x) for x in re.findall(r'Id="rId(\d+)"', srels)) + 1
srels = srels.replace('</Relationships>', f'<Relationship Id="rId{nxt}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide3.xml"/></Relationships>')
write(srels_p, srels); write(U / '[Content_Types].xml', ct)

# ---- 2. charts
def set_cache(cache, values, numeric):
    for pt in cache.findall(C + 'pt'): cache.remove(pt)
    cache.find(C + 'ptCount').set('val', str(len(values)))
    for i, v in enumerate(values):
        pt = etree.SubElement(cache, C + 'pt'); pt.set('idx', str(i)); etree.SubElement(pt, C + 'v').text = f'{v:.17g}' if numeric else str(v)
def ser_by_name(root, name):
    for ser in root.iter(C + 'ser'):
        v = ser.find(C + 'tx'); v = v.find('.//' + C + 'v') if v is not None else None
        if v is not None and v.text == name: return ser
    raise KeyError(name)
def strip_layouts(root):
    n = 0
    for dl in root.iter(C + 'dLbl'):
        if dl.find(C + 'delete') is not None: continue
        lay = dl.find(C + 'layout')
        if lay is not None: dl.remove(lay); n += 1
    return n

vol = {s: PUB['vol'][s][1:4] + [AUG['volume'][s]] for s in SER}; vol['Average'] = [sum(vol[s][j] for s in SER) / 5 for j in range(4)]
print('volume rows:', vol)
t = xml(U / 'ppt/charts/chart9.xml'); root = t.getroot()
for name in SER + ['Average']:
    ser = ser_by_name(root, name); set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), MONTHS, False); set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), vol[name], True)
save(t, U / 'ppt/charts/chart9.xml')
t = xml(U / 'ppt/charts/chart10.xml'); root = t.getroot(); ser = root.find('.//' + C + 'ser')
set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), CTRY, False); set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), [AUG['country'][c] for c in CTRY], True); n10 = strip_layouts(root); save(t, U / 'ppt/charts/chart10.xml')
t = xml(U / 'ppt/charts/chart11.xml'); root = t.getroot(); ser = root.find('.//' + C + 'ser')
set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), PLAT, False); set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), [AUG['platform'][p_] for p_ in PLAT], True); n11 = strip_layouts(root); save(t, U / 'ppt/charts/chart11.xml')
t = xml(U / 'ppt/charts/chart12.xml'); root = t.getroot()
pos = [AUG['sent'][c]['pos_pct'] for c in CATS]; neg = [AUG['sent'][c]['neg_pct'] for c in CATS]; neu = [1 - a - b for a, b in zip(pos, neg)]
for name, vals in (('positive', pos), ('neutral', neu), ('negative', neg)):
    ser = ser_by_name(root, name); set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), vals, True); set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), CATS, False)
for old, newname, xs in (('Positive June', 'Positive July', PUB['pos']), ('Negative June', 'Negative July', [1 - v for v in PUB['neg']])):
    ser = ser_by_name(root, old); ser.find(C + 'tx').find('.//' + C + 'v').text = newname
    set_cache(ser.find(C + 'xVal').find('.//' + C + 'numCache'), xs, True); set_cache(ser.find(C + 'yVal').find('.//' + C + 'numCache'), DIAMOND_Y, True)
for dl in list(root.iter(C + 'dLbl')):
    tx = dl.find(C + 'tx')
    if tx is not None and re.search(r'\d+%', ''.join(tx.itertext())): dl.remove(tx)
save(t, U / 'ppt/charts/chart12.xml'); print('pie layouts stripped', n10, n11)

# ---- 3. embeddings matching <c:f> refs
wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'; ws['A11'] = 'months'
for j, mth in enumerate(MONTHS): ws.cell(11, 2 + j, mth)
for i, s in enumerate(SER + ['Average']):
    ws.cell(12 + i, 1, s)
    for j in range(4): ws.cell(12 + i, 2 + j, f'=AVERAGE({"BCDE"[j]}12:{"BCDE"[j]}16)' if s == 'Average' else vol[s][j])
wb.save(U / 'ppt/embeddings/Microsoft_Excel_Worksheet_c9.xlsx')
for n, cats, vals in ((10, CTRY, [AUG['country'][c] for c in CTRY]), (11, PLAT, [AUG['platform'][p_] for p_ in PLAT])):
    wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'
    for i, (k, v) in enumerate(zip(cats, vals)): ws.cell(12 + i, 1, k); ws.cell(12 + i, 2, v)
    wb.save(U / f'ppt/embeddings/Microsoft_Excel_Worksheet_c{n}.xlsx')
wb = Workbook(); ws = wb.active; ws.title = 'Puma'
for j, c in enumerate(CATS): ws.cell(11, 10 + j, c)
for i, (lab, row) in enumerate(zip(['positive', 'neutral', 'negative'], (pos, neu, neg))):
    ws.cell(12 + i, 9, lab)
    for j, v in enumerate(row): ws.cell(12 + i, 10 + j, v)
ws.cell(17, 9, 'y'); ws.cell(19, 9, 'Positive July'); ws.cell(21, 9, 'Negative July')
for j in range(5): ws.cell(17, 10 + j, DIAMOND_Y[j]); ws.cell(19, 10 + j, PUB['pos'][j]); ws.cell(21, 10 + j, 1 - PUB['neg'][j])
wb.save(U / 'ppt/embeddings/Microsoft_Excel_Worksheet_c12.xlsx')

# ---- 4. slide text
sp = U / f'ppt/slides/{new}.xml'; t = xml(sp); root = t.getroot()
title = next(s for s in root.iter(P + 'sp') if s.find('.//' + P + 'cNvPr').get('name') == 'Title 3'); ts = title.findall('.//' + A + 't'); ts[0].text = AUGUST['subtitle']
for x in ts[1:]: x.text = ''
tb = next(s for s in root.iter(P + 'sp') if s.find('.//' + P + 'cNvPr').get('name') == 'TextBox 1'); body = tb.find(P + 'txBody')
paras = body.findall(A + 'p'); ppr = copy.deepcopy(paras[0].find(A + 'pPr'))
for p_ in paras: body.remove(p_)
srels = read(srels_p); srels = re.sub(r'<Relationship Id="rId\d+" Type="[^"]*hyperlink"[^>]*/>', '', srels)
nxt = max(int(x) for x in re.findall(r'Id="rId(\d+)"', srels)) + 1
def run(text, rid=None):
    r_ = etree.Element(A + 'r'); rpr = etree.SubElement(r_, A + 'rPr'); rpr.set('lang', 'en-GB'); rpr.set('sz', '1200'); rpr.set('dirty', '0')
    if rid: rpr.set('b', '1'); h = etree.SubElement(rpr, A + 'hlinkClick'); h.set(R + 'id', rid)
    te = etree.SubElement(r_, A + 't'); te.text = text
    if text != text.strip(): te.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r_
for runs in AUGUST['bullets']:
    p_ = etree.SubElement(body, A + 'p'); p_.append(copy.deepcopy(ppr))
    for text, url in runs:
        if url:
            rid = f'rId{nxt}'; nxt += 1
            srels = srels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/></Relationships>')
            p_.append(run(text, rid))
        else: p_.append(run(text))
write(srels_p, srels); save(t, sp)
# notes
np_ = U / 'ppt/notesSlides/notesSlide3.xml'; t = xml(np_); root = t.getroot()
body = next(s for s in root.iter(P + 'sp') if s.find(P + 'txBody') is not None and 'Insight 1' in ''.join(s.itertext())).find(P + 'txBody')
for p_ in body.findall(A + 'p'): body.remove(p_)
for bold, text in AUGUST['notes']:
    p_ = etree.SubElement(body, A + 'p'); r_ = etree.SubElement(p_, A + 'r'); rpr = etree.SubElement(r_, A + 'rPr'); rpr.set('lang', 'en-GB'); rpr.set('sz', '1100'); rpr.set('dirty', '0')
    if bold: rpr.set('b', '1')
    etree.SubElement(r_, A + 't').text = text
save(t, np_)
# title slide
write(U / 'ppt/slides/slide1.xml', read(U / 'ppt/slides/slide1.xml').replace('Social Media July 2026', 'Social Media August 2026').replace('17.08.2026', REPORT_DATE))

# ---- 5. clean, pack, validate
r = subprocess.run([sys.executable, str(PPTX_SCRIPTS / 'clean.py'), str(U)], capture_output=True, text=True); print(r.stdout[-300:], r.stderr[-300:])
if OUT.exists(): OUT.unlink()
subprocess.run(['zip', '-Xr', str(OUT), '.'], cwd=U, check=True, capture_output=True); print('packed', OUT.stat().st_size)
r = subprocess.run([sys.executable, str(PPTX_SCRIPTS / 'office/validate.py'), str(OUT), '--original', str(LAST)], capture_output=True, text=True); print(r.stdout[-2500:], r.stderr[-1000:])
