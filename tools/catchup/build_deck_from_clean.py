"""Catch-up build used for Capri Aug-2026 (July slide never delivered): last deck (June) + cleaned data -> title -> August -> July.
Reference implementation - constants below (series names, 7 sentiment categories, label text boxes, axis pin) are the CAPRI template;
read them from the last deck's chart XML before reusing for another carline. See tools/catchup/README.md.
Structural work first (duplicate slides, clone chart parts, notes), then content (charts, bullets, labels, notes, title)."""
import os, re, shutil, subprocess, sys, json, copy
from pathlib import Path
from lxml import etree
from openpyxl import Workbook

import argparse
ap = argparse.ArgumentParser(description='Catch-up build: last delivered deck + cleaned data -> new deck with TWO new month slides')
ap.add_argument('--work', required=True, help='work dir holding <last>_unpacked/, content_pack.json, slide_text.py')
ap.add_argument('--last-unpacked', default='june_unpacked', help='unpacked last delivered deck (dir name inside --work)')
ap.add_argument('--last-pptx', required=True, help='the last delivered .pptx (validate.py baseline)')
ap.add_argument('--out', required=True); ap.add_argument('--report-date', required=True)
ap.add_argument('--pptx-scripts', default=os.environ.get('PPTX_SCRIPTS', '/mnt/skills/public/pptx/scripts'))
args = ap.parse_args()
WORK = Path(args.work); sys.path.insert(0, str(WORK))
from slide_text import JULY, AUGUST   # the two new months' bullets/notes (newer month = AUGUST slot); kept out of git - client content
SRC = WORK / args.last_unpacked
BUILD = WORK / 'build'; U = BUILD / 'unpacked'
OUT = Path(args.out); JUNE_PPTX = Path(args.last_pptx); PPTX_SCRIPTS = Path(args.pptx_scripts)
PACK = json.load(open(WORK / 'content_pack.json'))
REPORT_DATE = args.report_date

NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
C = '{%s}' % NS['c']; A = '{%s}' % NS['a']; P = '{%s}' % NS['p']; R = '{%s}' % NS['r']
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
CT_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml'
CT_STYLE = 'application/vnd.ms-office.chartstyle+xml'
CT_NOTES = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
LINE_SERIES = ['Polestar 2', 'E-3008', 'EV6', 'ID.5', 'Q4 Sportback e-tron', 'Capri', 'Average', 'Enyaq Coupe']
MODELS7 = ['Polestar 2', 'E-3008', 'EV6', 'ID.5', 'Q4 Sportback e-tron', 'Capri', 'Enyaq Coupe']
SENT_CATS = ['Audi Q4 Sportback e-tron', 'Kia EV6', 'Peugeot E 3008', 'Polestar 2', 'Skoda Enyaq Coupe', 'VW ID.5', 'Ford Capri']
PLATFORMS = ['Online News', 'X', 'Forums', 'Blogs', 'Review', 'Tumblr']
COUNTRIES = ['UK', 'IT', 'FR', 'DE', 'ES', 'NL', 'NO']
DIAMOND_Y = [2.1, 6.4499999999999993, 10.799999999999999, 15.049999999999999, 19.299999999999997, 23.529999999999998, 27.799999999999997]
JUNE = PACK['june']; JUL = PACK['months']['7']; AUG = PACK['months']['8']

def read(p): return Path(p).read_text(encoding='utf-8')
def write(p, s): Path(p).write_text(s, encoding='utf-8')
def xml(p): return etree.parse(str(p), etree.XMLParser(remove_blank_text=False))
def save(tree, p): tree.write(str(p), xml_declaration=True, encoding='UTF-8', standalone=True)

# ------------------------------------------------------------------ 1. fresh copy + duplicate slides
if BUILD.exists(): shutil.rmtree(BUILD)
BUILD.mkdir(); shutil.copytree(SRC, U)
def add_slide():
    r = subprocess.run([sys.executable, str(PPTX_SCRIPTS / 'add_slide.py'), str(U), 'slide2.xml'], capture_output=True, text=True)
    if r.returncode: sys.exit(r.stderr)
    m = re.search(r'ppt/slides/(slide\d+)\.xml', r.stdout); print('  ', r.stdout.strip().splitlines()[0]); return m.group(1)
aug_slide = add_slide(); jul_slide = add_slide()   # slide4 (August), slide5 (July)
print('new slides:', aug_slide, jul_slide)

# order: title, August, July ; drop slide2 (June) and slide3 (May)
pres = read(U / 'ppt/presentation.xml')
rels = read(U / 'ppt/_rels/presentation.xml.rels')
def rid_of(slide): return re.search(rf'Id="(rId\d+)"[^>]*Target="slides/{slide}\.xml"', rels).group(1) if re.search(rf'Id="(rId\d+)"[^>]*Target="slides/{slide}\.xml"', rels) else re.search(rf'Target="slides/{slide}\.xml"[^>]*Id="(rId\d+)"', rels).group(1)
ids = dict(re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', pres))
byrid = {v: k for k, v in ids.items()}
order = [(byrid[rid_of('slide1')], rid_of('slide1')), (byrid[rid_of(aug_slide)], rid_of(aug_slide)), (byrid[rid_of(jul_slide)], rid_of(jul_slide))]
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', '<p:sldIdLst>' + ''.join(f'<p:sldId id="{i}" r:id="{r}"/>' for i, r in order) + '</p:sldIdLst>', pres, flags=re.S)
pres = re.sub(r'<p14:sldIdLst>.*?</p14:sldIdLst>', '<p14:sldIdLst>' + ''.join(f'<p14:sldId id="{i}"/>' for i, _ in order) + '</p14:sldIdLst>', pres, flags=re.S)
write(U / 'ppt/presentation.xml', pres)

# ------------------------------------------------------------------ 2. clone chart parts + notes per new slide
ct = read(U / '[Content_Types].xml')
def add_override(part, ctype):
    global ct
    if f'PartName="{part}"' not in ct:
        ct = ct.replace('</Types>', f'<Override PartName="{part}" ContentType="{ctype}"/></Types>')

def clone_charts(slide, base):
    """chart1..4 -> chart{base}..{base+3}; returns {role_index: chart_name}"""
    srels_p = U / f'ppt/slides/_rels/{slide}.xml.rels'; srels = read(srels_p)
    out = {}
    for i in range(1, 5):
        n = base + i - 1
        shutil.copy(U / f'ppt/charts/chart{i}.xml', U / f'ppt/charts/chart{n}.xml')
        shutil.copy(U / f'ppt/charts/colors{i}.xml', U / f'ppt/charts/colors{n}.xml')
        shutil.copy(U / f'ppt/charts/style{i}.xml', U / f'ppt/charts/style{n}.xml')
        crels = read(U / f'ppt/charts/_rels/chart{i}.xml.rels')
        crels = crels.replace(f'colors{i}.xml', f'colors{n}.xml').replace(f'style{i}.xml', f'style{n}.xml')
        crels = re.sub(r'Target="\.\./embeddings/[^"]+"', f'Target="../embeddings/Microsoft_Excel_Worksheet_c{n}.xlsx"', crels)
        write(U / f'ppt/charts/_rels/chart{n}.xml.rels', crels)
        add_override(f'/ppt/charts/chart{n}.xml', CT_CHART); add_override(f'/ppt/charts/colors{n}.xml', CT_COLORS); add_override(f'/ppt/charts/style{n}.xml', CT_STYLE)
        srels = srels.replace(f'Target="../charts/chart{i}.xml"', f'Target="../charts/chart{n}.xml"')
        out[i] = n
    write(srels_p, srels); return out

def add_notes(slide, k):
    shutil.copy(U / 'ppt/notesSlides/notesSlide1.xml', U / f'ppt/notesSlides/notesSlide{k}.xml')
    nrels = read(U / 'ppt/notesSlides/_rels/notesSlide1.xml.rels').replace('../slides/slide2.xml', f'../slides/{slide}.xml')
    write(U / f'ppt/notesSlides/_rels/notesSlide{k}.xml.rels', nrels)
    add_override(f'/ppt/notesSlides/notesSlide{k}.xml', CT_NOTES)
    srels_p = U / f'ppt/slides/_rels/{slide}.xml.rels'; srels = read(srels_p)
    nxt = max(int(x) for x in re.findall(r'Id="rId(\d+)"', srels)) + 1
    srels = srels.replace('</Relationships>', f'<Relationship Id="rId{nxt}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{k}.xml"/></Relationships>')
    write(srels_p, srels)

aug_charts = clone_charts(aug_slide, 9); jul_charts = clone_charts(jul_slide, 13)
add_notes(aug_slide, 3); add_notes(jul_slide, 4)
write(U / '[Content_Types].xml', ct)
print('charts:', aug_charts, jul_charts)

# ------------------------------------------------------------------ 3. chart helpers
def set_cache(cache_el, values, numeric):
    for pt in cache_el.findall(C + 'pt'): cache_el.remove(pt)
    pc = cache_el.find(C + 'ptCount'); pc.set('val', str(len(values)))
    for i, v in enumerate(values):
        pt = etree.SubElement(cache_el, C + 'pt'); pt.set('idx', str(i))
        ve = etree.SubElement(pt, C + 'v'); ve.text = (f'{v:.17g}' if numeric else str(v))

def ser_by_name(root, name):
    for ser in root.iter(C + 'ser'):
        tx = ser.find(C + 'tx'); v = tx.find('.//' + C + 'v') if tx is not None else None
        if v is not None and v.text == name: return ser
    raise KeyError(name)

def strip_manual_layout_in_dlbls(root):
    n = 0
    for dl in root.iter(C + 'dLbl'):
        if dl.find(C + 'delete') is not None: continue
        lay = dl.find(C + 'layout')
        if lay is not None: dl.remove(lay); n += 1
    return n

def update_volume(chart_p, months, vol_by_series):
    t = xml(chart_p); root = t.getroot()
    for name in LINE_SERIES:
        ser = ser_by_name(root, name)
        set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), months, False)
        set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), vol_by_series[name], True)
    # pin value axis to the June rendering scale so the manual line labels keep their mapping
    valax = root.find('.//' + C + 'valAx'); sc = valax.find(C + 'scaling')
    if sc.find(C + 'max') is None:
        mx = etree.Element(C + 'max'); mx.set('val', '600'); sc.insert(1, mx)  # after orientation
    save(t, chart_p)

def update_pie(chart_p, cats, vals):
    t = xml(chart_p); root = t.getroot(); ser = root.find('.//' + C + 'ser')
    set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), cats, False)
    set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), vals, True)
    n = strip_manual_layout_in_dlbls(root); save(t, chart_p); return n

def update_sentiment(chart_p, sent, pre_pos, pre_neg):
    t = xml(chart_p); root = t.getroot()
    pos = [sent[c]['pos_pct'] for c in SENT_CATS]; neg = [sent[c]['neg_pct'] for c in SENT_CATS]; neu = [1 - p - n for p, n in zip(pos, neg)]
    for name, vals in (('positive', pos), ('neutral', neu), ('negative', neg)):
        ser = ser_by_name(root, name); set_cache(ser.find(C + 'val').find('.//' + C + 'numCache'), vals, True)
        set_cache(ser.find(C + 'cat').find('.//' + C + 'strCache'), SENT_CATS, False)
    for name, xs in (('positive pre', pre_pos), ('Negative Pre', [1 - n for n in pre_neg])):
        ser = ser_by_name(root, name)
        set_cache(ser.find(C + 'xVal').find('.//' + C + 'numCache'), xs, True)
        set_cache(ser.find(C + 'yVal').find('.//' + C + 'numCache'), DIAMOND_Y, True)
    # strip any manual data-label text overrides
    for dl in list(root.iter(C + 'dLbl')):
        tx = dl.find(C + 'tx')
        if tx is not None and re.search(r'\d+%', ''.join(tx.itertext())): dl.remove(tx)
    save(t, chart_p)

# ------------------------------------------------------------------ 4. embeddings (sheet names / cells match the chart <c:f> refs)
def build_embeddings(charts, months, vol, plat, ctry, sent, pre_pos, pre_neg):
    wb = Workbook(); ws = wb.active; ws.title = 'Capri'
    ws['A1'] = None
    for j, m in enumerate(months): ws.cell(1, 2 + j, m)
    for i, name in enumerate(LINE_SERIES):
        r = {'Polestar 2': 2, 'E-3008': 3, 'EV6': 4, 'ID.5': 5, 'Q4 Sportback e-tron': 6, 'Capri': 7, 'Enyaq Coupe': 8, 'Average': 9}[name]
        ws.cell(r, 1, name)
        for j in range(4):
            col = 'BCDE'[j]
            ws.cell(r, 2 + j, f'=AVERAGE({col}2:{col}8)' if name == 'Average' else vol[name][j])
    wb.save(U / f'ppt/embeddings/Microsoft_Excel_Worksheet_c{charts[1]}.xlsx')
    wb = Workbook(); ws = wb.active; ws.title = 'Capri-sources'
    for i, p in enumerate(PLATFORMS): ws.cell(1 + i, 1, p); ws.cell(1 + i, 2, plat[p])
    wb.save(U / f'ppt/embeddings/Microsoft_Excel_Worksheet_c{charts[2]}.xlsx')
    wb = Workbook(); ws = wb.active; ws.title = 'Capri'
    for i, cc in enumerate(COUNTRIES): ws.cell(1 + i, 1, cc); ws.cell(1 + i, 2, ctry[cc])
    wb.save(U / f'ppt/embeddings/Microsoft_Excel_Worksheet_c{charts[3]}.xlsx')
    wb = Workbook(); ws = wb.active; ws.title = 'Capri'
    for j, cat in enumerate(SENT_CATS): ws.cell(1, 2 + j, cat)
    labels = ['positive', 'neutral', 'negative', 'positive pre', 'Negative Pre', 'y']
    for i, lab in enumerate(labels): ws.cell(2 + i, 1, lab)
    for j, cat in enumerate(SENT_CATS):
        p, n = sent[cat]['pos_pct'], sent[cat]['neg_pct']
        for i, v in enumerate([p, 1 - p - n, n, pre_pos[j], 1 - pre_neg[j], DIAMOND_Y[j]]): ws.cell(2 + i, 2 + j, v)
    wb.save(U / f'ppt/embeddings/Microsoft_Excel_Worksheet_c{charts[4]}.xlsx')

# ------------------------------------------------------------------ 5. slide content
def label_y(value):  # centre y (EMU) of a line-end label for a given value, fitted on the June deck (Capri 130, Enyaq 31), axis 0-600
    return 3309154 - (value - 31) * 2695.3
def set_label_boxes(root, vol):
    boxes = {'TextBox 25': ('Capri', 173427), 'TextBox 28': ('Average', 246221), 'TextBox 14': ('Enyaq Coupe', 230832)}
    centres = {}
    for sp in root.iter(P + 'sp'):
        nm = sp.find('.//' + P + 'cNvPr').get('name')
        if nm in boxes:
            ser, h = boxes[nm]; centres[nm] = [label_y(vol[ser][3]), h, sp]
    # push apart labels closer than 145k EMU (9pt line height)
    items = sorted(centres.items(), key=lambda kv: kv[1][0])
    for i in range(1, len(items)):
        prev, cur = items[i - 1][1], items[i][1]
        if cur[0] - prev[0] < 145000: cur[0] = prev[0] + 145000
    for nm, (cy, h, sp) in centres.items():
        off = sp.find('.//' + A + 'off'); off.set('y', str(int(cy - h / 2)))

def run(text, url_rid=None):
    r = etree.Element(A + 'r'); rpr = etree.SubElement(r, A + 'rPr'); rpr.set('lang', 'en-GB'); rpr.set('sz', '1200'); rpr.set('dirty', '0')
    if url_rid:
        rpr.set('b', '1'); h = etree.SubElement(rpr, A + 'hlinkClick'); h.set(R + 'id', url_rid)
    t = etree.SubElement(r, A + 't'); t.text = text
    if text != text.strip(): t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r

def set_bullets(root, slide, bullets):
    # find the bullets text box (TextBox 16) and its txBody
    tb = next(sp for sp in root.iter(P + 'sp') if sp.find('.//' + P + 'cNvPr').get('name') == 'TextBox 16')
    body = tb.find(P + 'txBody'); paras = body.findall(A + 'p'); template_ppr = [copy.deepcopy(p.find(A + 'pPr')) for p in paras]
    for p in paras: body.remove(p)
    srels_p = U / f'ppt/slides/_rels/{slide}.xml.rels'; srels = read(srels_p)
    srels = re.sub(r'<Relationship Id="rId\d+" Type="[^"]*hyperlink"[^>]*/>', '', srels)   # drop June's hyperlink
    nxt = max(int(x) for x in re.findall(r'Id="rId(\d+)"', srels)) + 1
    for i, runs in enumerate(bullets):
        p = etree.SubElement(body, A + 'p'); p.append(copy.deepcopy(template_ppr[min(i, len(template_ppr) - 1)]))
        for text, url in runs:
            if url:
                rid = f'rId{nxt}'; nxt += 1
                srels = srels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/></Relationships>')
                p.append(run(text, rid))
            else: p.append(run(text))
    write(srels_p, srels)

def set_title(root, text):
    sp = next(sp for sp in root.iter(P + 'sp') if sp.find('.//' + P + 'cNvPr').get('name') == 'Title 3')
    ts = sp.findall('.//' + A + 't'); ts[0].text = text
    for t in ts[1:]: t.text = ''

def set_notes(k, notes):
    p = U / f'ppt/notesSlides/notesSlide{k}.xml'; t = xml(p); root = t.getroot()
    body = next(sp for sp in root.iter(P + 'sp') if sp.find(P + 'txBody') is not None and 'Country breakdown' in ''.join(sp.itertext())).find(P + 'txBody')
    for para in body.findall(A + 'p'): body.remove(para)
    for bold, text in notes:
        para = etree.SubElement(body, A + 'p'); r = etree.SubElement(para, A + 'r'); rpr = etree.SubElement(r, A + 'rPr')
        rpr.set('lang', 'en-GB'); rpr.set('sz', '1200'); rpr.set('dirty', '0')
        if bold: rpr.set('b', '1')
        etree.SubElement(r, A + 't').text = text
    save(t, p)

def build_month(slide, charts, k, spec, months, vol, plat, ctry, sent, pre_pos, pre_neg):
    update_volume(U / f'ppt/charts/chart{charts[1]}.xml', months, vol)
    n2 = update_pie(U / f'ppt/charts/chart{charts[2]}.xml', PLATFORMS, [plat[p] for p in PLATFORMS])
    n3 = update_pie(U / f'ppt/charts/chart{charts[3]}.xml', COUNTRIES, [ctry[c] for c in COUNTRIES])
    update_sentiment(U / f'ppt/charts/chart{charts[4]}.xml', sent, pre_pos, pre_neg)
    build_embeddings(charts, months, vol, plat, ctry, sent, pre_pos, pre_neg)
    sp = U / f'ppt/slides/{slide}.xml'; t = xml(sp); root = t.getroot()
    set_title(root, spec['subtitle']); set_label_boxes(root, vol); set_bullets(root, slide, spec['bullets']); save(t, sp)
    set_notes(k, spec['notes'])
    print(f'  {slide}: charts {charts} updated; pie layouts stripped {n2}/{n3}')

# volumes: series -> 4 values
def vol4(prev_from_june, jul, aug=None):
    out = {}
    for s in MODELS7:
        seq = (JUNE['vol'][s][1:4] + [JUL['volume'][s]]) if aug is None else (JUNE['vol'][s][2:4] + [JUL['volume'][s], AUG['volume'][s]])
        out[s] = seq
    out['Average'] = [sum(out[s][j] for s in MODELS7) / 7 for j in range(4)]
    return out
VOL_JUL = vol4(True, JUL); VOL_AUG = vol4(True, JUL, AUG)
print('July volume rows:', {k: v for k, v in VOL_JUL.items()}); print('August volume rows:', {k: v for k, v in VOL_AUG.items()})

build_month(jul_slide, jul_charts, 4, JULY, ['April', 'May', 'June', 'July'], VOL_JUL, JUL['platform'], JUL['country'], JUL['sentiment'], JUNE['pos'], JUNE['neg'])
build_month(aug_slide, aug_charts, 3, AUGUST, ['May', 'June', 'July', 'August'], VOL_AUG, AUG['platform'], AUG['country'], AUG['sentiment'],
            [JUL['sentiment'][c]['pos_pct'] for c in SENT_CATS], [JUL['sentiment'][c]['neg_pct'] for c in SENT_CATS])

# ------------------------------------------------------------------ 6. title slide
s1 = read(U / 'ppt/slides/slide1.xml').replace('Social Media June 2026', 'Social Media August 2026').replace('10.07.2026', REPORT_DATE)
write(U / 'ppt/slides/slide1.xml', s1)

# ------------------------------------------------------------------ 7. clean, pack, validate
r = subprocess.run([sys.executable, str(PPTX_SCRIPTS / 'clean.py'), str(U)], capture_output=True, text=True); print(r.stdout[-800:], r.stderr[-400:])
if OUT.exists(): OUT.unlink()
subprocess.run(['zip', '-Xr', str(OUT), '.'], cwd=U, check=True, capture_output=True)
print('packed', OUT, OUT.stat().st_size)
r = subprocess.run([sys.executable, str(PPTX_SCRIPTS / 'office/validate.py'), str(OUT), '--original', str(JUNE_PPTX)], capture_output=True, text=True)
print(r.stdout[-3000:], r.stderr[-1500:])
