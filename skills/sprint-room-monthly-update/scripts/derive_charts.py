#!/usr/bin/env python3
"""
derive_charts.py — Raw-only path: build the four Brandwatch chart-export tables from the
cleaned raw feed + last month's published deck, so the monthly run needs only
`raw.xlsx` + `last_month.pptx` (no manual chart exports).

  volume    : rolling 4 months. Older 3 months CARRIED from last month's published volume
              chart (golden rule — never re-derive prior months); new month = COUNTUNIQUE(URL)
              per model from cleaned raw. Oldest month dropped.
  country   : carline, new month only, unique-URL per country.
  platform  : carline, new month only, unique-URL per Page Type Name.
  sentiment : per model, new month, pos/neu/neg mention counts (cleaned tags).
              (Prior-month diamonds are read from the deck by update_sentiment_chart.)

Outputs Brandwatch-format xlsx (inline strings) that the skill's update_*_chart.py consume.
"""
import argparse, os, re, sys, zipfile
from collections import defaultdict
sys.path.insert(0, os.path.dirname(__file__))
import clean_raw as cr

MONTHS = ["","January","February","March","April","May","June","July","August",
          "September","October","November","December"]
ISO3 = {"GBR":"GBR","ITA":"ITA","FRA":"FRA","DEU":"DEU","ESP":"ESP","NLD":"NLD","NOR":"NOR"}

# ---------- minimal Brandwatch-style xlsx writer (inline strings + numeric <v>) ----------
def _cell(ref, val):
    if isinstance(val,(int,float)):
        return f'<c r="{ref}"><v>{val}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{val}</t></is></c>'

def write_xlsx(path, rows):
    """rows = list of (rownum, [(col,val),...])"""
    body=[]
    for rn,cells in rows:
        body.append(f'<row r="{rn}">'+"".join(_cell(f"{c}{rn}",v) for c,v in cells)+"</row>")
    sheet=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetData>'+"".join(body)+'</sheetData></worksheet>')
    ct=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>')
    rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
    wb=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wbrels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '</Relationships>')
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",ct)
        z.writestr("_rels/.rels",rels)
        z.writestr("xl/workbook.xml",wb)
        z.writestr("xl/_rels/workbook.xml.rels",wbrels)
        z.writestr("xl/worksheets/sheet1.xml",sheet)

# ---------- locate the reported-month slide's charts by role (carline-agnostic) ----------
_MONTHS_IDX={m.lower():i for i,m in enumerate(MONTHS) if m}
def _infer_role(types, series_names, cats):
    sn=[s.lower() for s in series_names]
    if "scatterChart" in types or any(s in ("positive","neutral","negative") for s in sn):
        return "sentiment"
    if "lineChart" in types:
        return "volume"
    if "pieChart" in types:
        j=" ".join(cats).lower()
        if any(w in j for w in ["news","forum","blog","tumblr","review"]) or " x " in f" {j} ":
            return "platform"
        return "country"
    return "unknown"

def reported_charts(pptx):
    """Return (chart_paths_by_role, volume_months, volume_series) for the newest-month slide."""
    z=zipfile.ZipFile(pptx)
    def slide_charts(slide):
        try: rels=z.read(f"ppt/slides/_rels/{slide}.xml.rels").decode()
        except KeyError: return []
        return [f"ppt/charts/{t}" for t in re.findall(r'Target="\.\./charts/(chart\d+\.xml)"',rels)]
    best=None
    for n in z.namelist():
        m=re.match(r"ppt/slides/(slide\d+)\.xml$",n)
        if not m: continue
        if not slide_charts(m.group(1)): continue   # only chart-bearing slides (skip title)
        txt=z.read(n).decode("utf-8","replace")
        mm=re.search(r"\(EU\d\s+([A-Za-z]+)\s+(20\d\d)\)", txt) or re.search(r"([A-Za-z]+)\s+(20\d\d)", txt)
        if not mm: continue
        mo=_MONTHS_IDX.get(mm.group(1).lower())
        if not mo: continue
        key=(int(mm.group(2)),mo)
        if best is None or key>best[0]: best=(key,m.group(1))
    if not best: raise RuntimeError("no month-labelled chart slide found")
    slide=best[1]
    charts=slide_charts(slide)
    roles={}; vmonths=[]; vseries={}
    for cp in charts:
        xml=z.read(cp).decode("utf-8","replace")
        catm=re.search(r'<c:cat>(.*?)</c:cat>',xml,re.S)
        cats=re.findall(r'<c:v>([^<]*)</c:v>',catm.group(1)) if catm else []
        types=re.findall(r'<c:(barChart|lineChart|pieChart|scatterChart|doughnutChart)>',xml)
        snames=re.findall(r'<c:tx>.*?<c:v>([^<]*)</c:v>',xml,re.S)
        role=_infer_role(types,snames,cats)
        roles[role]=cp
        if role=="volume":
            vmonths=cats
            for ser in re.findall(r'<c:ser>(.*?)</c:ser>',xml,re.S):
                nm=re.search(r'<c:tx>.*?<c:v>([^<]*)</c:v>',ser,re.S)
                val=re.search(r'<c:val>(.*?)</c:val>',ser,re.S)
                if nm and val:
                    vseries[nm.group(1)]=[float(v) for v in re.findall(r'<c:v>([^<]*)</c:v>',val.group(1))]
    return roles, vmonths, vseries

def map_deck_to_raw(deck_names, raw_queries):
    """Auto-map a deck volume series name -> raw Query Name using clean_raw model keywords."""
    out={}
    for d in deck_names:
        dl=d.lower(); best=None
        for q in raw_queries:
            for kw in cr.MODEL_KW.get(q,[q.lower()]):
                if kw in dl or dl in kw:
                    best=q; break
            if best: break
        out[d]=best
    return out

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--carline",required=True)
    ap.add_argument("--month",required=True)
    ap.add_argument("--raw",required=True)
    ap.add_argument("--also",nargs="*",default=[])
    ap.add_argument("--last-month",required=True)
    ap.add_argument("--carline-query",default="Ford Capri")
    ap.add_argument("--retag-passing",action="store_true")
    ap.add_argument("--outdir",required=True)
    a=ap.parse_args()
    yr,mm=map(int,a.month.split("-")); prior=f"{yr-1 if mm==1 else yr}-{12 if mm==1 else mm-1:02d}"
    months=(prior,a.month); cq=a.carline_query
    os.makedirs(a.outdir,exist_ok=True)

    rows=cr.load_merged(a.raw,a.also,months)
    cleaned,_=cr.apply_clean(rows,cq,a.retag_passing)

    # unique-URL volume per model per month; sentiment counts per model; carline country/platform
    vol=defaultdict(lambda:defaultdict(set)); sent=defaultdict(lambda:defaultdict(lambda:defaultdict(int)))
    country=defaultdict(set); platform=defaultdict(set)
    for r in cleaned:
        q,mo=r.get("Query Name"),cr.ym(r.get("Date")); url=r.get("Url")
        vol[q][mo].add(url)
        sent[q][mo][(r.get("Sentiment") or "").strip()]+=1
        if q==cq and mo==a.month:
            cc=(r.get("Country Code") or "").strip()
            if cc: country[cc].add(url)
            pt=(r.get("Page Type Name") or "").strip()
            if pt: platform[pt].add(url)

    # ---- VOLUME ----
    roles, dmonths, dser = reported_charts(a.last_month)
    print(f"Detected reported-month chart roles: { {k:v.split('/')[-1] for k,v in roles.items()} }")
    deck2raw = map_deck_to_raw([s for s in dser if s!="Average"], list(vol.keys()))
    new_name=MONTHS[mm]
    carried=[MONTHS[(mm-3-1)%12+1],MONTHS[(mm-2-1)%12+1],MONTHS[(mm-1-1)%12+1]]  # 3 prior calendar months
    vrows=[(11,[("A","months"),("B",carried[0]),("C",carried[1]),("D",carried[2]),("E",new_name)])]
    rn=12; all_vals=[]
    for dname in [s for s in dser if s!="Average"]:
        rawq=deck2raw.get(dname)
        prior3=dser[dname][1:4]  # drop oldest (index0), keep last 3 of the deck's 4
        july=len(vol[rawq][a.month]) if rawq else 0
        vals=[int(prior3[0]),int(prior3[1]),int(prior3[2]),july]
        vrows.append((rn,[("A",dname),("B",vals[0]),("C",vals[1]),("D",vals[2]),("E",vals[3])]))
        all_vals.append(vals)  # deck's Average is the mean of ALL models incl carline
        rn+=1
    avg=[round(sum(c[i] for c in all_vals)/len(all_vals),4) for i in range(4)]
    vrows.append((rn,[("A","Average"),("B",avg[0]),("C",avg[1]),("D",avg[2]),("E",avg[3])]))
    write_xlsx(os.path.join(a.outdir,"chart_volume.xlsx"),vrows)

    # ---- COUNTRY (carline, new month) ----
    crows=[(11,[("A","queries"),("B",a.carline),("C","population")])]
    for i,(cc,urls) in enumerate(sorted(country.items(),key=lambda kv:-len(kv[1]))):
        crows.append((12+i,[("A",cc),("B",len(urls))]))
    write_xlsx(os.path.join(a.outdir,"chart_country.xlsx"),crows)

    # ---- PLATFORM (carline, new month) ----
    prows=[(11,[("A","queries"),("B",a.carline)])]
    for i,(pt,urls) in enumerate(sorted(platform.items(),key=lambda kv:-len(kv[1]))):
        prows.append((12+i,[("A",pt),("B",len(urls))]))
    write_xlsx(os.path.join(a.outdir,"chart_platform.xlsx"),prows)

    # ---- SENTIMENT (all models, new month; counts) ----
    model_order=[deck2raw[s] for s in dser if s!="Average"]
    srows=[(11,[("A","queries")]+[(chr(66+i),m) for i,m in enumerate(model_order)])]
    for rn2,key in [(12,"positive"),(13,"neutral"),(14,"negative")]:
        srows.append((rn2,[("A",key)]+[(chr(66+i),sent[m][a.month].get(key,0)) for i,m in enumerate(model_order)]))
    write_xlsx(os.path.join(a.outdir,"chart_sentiment.xlsx"),srows)

    # ---- report ----
    print(f"Derived 4 chart tables -> {a.outdir}")
    print(f"\nVOLUME (carried {carried} from deck + {new_name} from raw):")
    for rn2,cells in vrows[1:]:
        print("  "+cells[0][1].ljust(24)+str([c[1] for c in cells[1:]]))
    print(f"\nCOUNTRY ({a.carline} {new_name}): "+str({c[0][1]:c[1][1] for _,c in [(r[0],r[1]) for r in crows[1:]]}))
    print(f"PLATFORM ({a.carline} {new_name}): "+str({c[0][1]:c[1][1] for _,c in [(r[0],r[1]) for r in prows[1:]]}))
    print(f"\nSENTIMENT {new_name} (pos/neu/neg counts):")
    for m in model_order:
        s=sent[m][a.month]; tot=s.get('positive',0)+s.get('neutral',0)+s.get('negative',0)
        print(f"  {m:30s} {s.get('positive',0):>3}/{s.get('neutral',0):>3}/{s.get('negative',0):>3}"
              f"  ({100*s.get('positive',0)/tot:.0f}%/{100*s.get('neutral',0)/tot:.0f}%/{100*s.get('negative',0)/tot:.0f}%)")

if __name__=="__main__":
    main()
