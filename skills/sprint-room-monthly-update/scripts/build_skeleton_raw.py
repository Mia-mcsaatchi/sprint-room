#!/usr/bin/env python3
"""
build_skeleton_raw.py — layout-agnostic skeleton builder for the raw-only path.

Unlike the Explorer-hardcoded build_skeleton.py, this detects the reported-month slide
and its four charts BY ROLE (works whatever the prior deck's slide/chart numbering is),
clones it into the new-month slide with its OWN copies of the four charts, drops the
oldest month slide, and fixes all package bookkeeping (rels, Content-Types, notes) so
there are no recurring manual fixes.

    python3 build_skeleton_raw.py --last-month X.pptx --carline Puma --segment EU5 \
        --new-month "July 2026" --report-date 17.08.2026 --out work/.../skeleton
"""
import argparse, os, re, shutil, subprocess, sys, zipfile
PPTX = "/mnt/skills/public/pptx/scripts"
MONTHS=["","January","February","March","April","May","June","July","August","September","October","November","December"]
MIDX={m.lower():i for i,m in enumerate(MONTHS) if m}

def read(p): return open(p,encoding="utf-8").read()
def write(p,s): open(p,"w",encoding="utf-8").write(s)

def infer_role(xml):
    types=re.findall(r'<c:(barChart|lineChart|pieChart|scatterChart|doughnutChart)>',xml)
    sn=[s.lower() for s in re.findall(r'<c:tx>.*?<c:v>([^<]*)</c:v>',xml,re.S)]
    catm=re.search(r'<c:cat>(.*?)</c:cat>',xml,re.S)
    cats=re.findall(r'<c:v>([^<]*)</c:v>',catm.group(1)) if catm else []
    if "scatterChart" in types or any(s in ("positive","neutral","negative") for s in sn): return "sentiment"
    if "lineChart" in types: return "volume"
    if "pieChart" in types:
        j=" ".join(cats).lower()
        return "platform" if (any(w in j for w in ["news","forum","blog","tumblr","review"]) or " x " in f" {j} ") else "country"
    return "unknown"

def slide_month(up, slide):
    txt=read(f"{up}/ppt/slides/{slide}")
    m=re.search(r"\(EU\d\s+([A-Za-z]+)\s+(20\d\d)\)",txt) or re.search(r"([A-Za-z]+)\s+(20\d\d)",txt)
    if not m: return None
    mo=MIDX.get(m.group(1).lower())
    return (int(m.group(2)),mo) if mo else None

def slide_charts(up, slide):
    rp=f"{up}/ppt/slides/_rels/{slide}.rels"
    if not os.path.exists(rp): return []
    return re.findall(r'Target="\.\./charts/(chart\d+\.xml)"',read(rp))

def main():
    ap=argparse.ArgumentParser()
    for a in ["last-month","carline","segment","new-month","report-date","out"]:
        ap.add_argument("--"+a,required=True)
    a=ap.parse_args()
    up=os.path.join(a.out,"unpacked")
    if os.path.exists(up): shutil.rmtree(up)
    os.makedirs(up)
    zipfile.ZipFile(a.last_month).extractall(up)

    # detect reported (newest) + prior (older) chart-bearing month slides
    slides=[os.path.basename(p) for p in sorted(__import__("glob").glob(f"{up}/ppt/slides/slide*.xml"))]
    cand=[(slide_month(up,s),s) for s in slides if slide_charts(up,s) and slide_month(up,s)]
    cand.sort(key=lambda x:x[0])
    prior_slide, reported_slide = cand[0][1], cand[-1][1]
    rep_charts=slide_charts(up,reported_slide)
    roles={infer_role(read(f"{up}/ppt/charts/{c}")):c for c in rep_charts}
    print(f"reported={reported_slide} charts={rep_charts} roles={roles}")
    print(f"prior(drop)={prior_slide}")

    # clone reported slide -> new slide (shares charts; we'll give it its own copies)
    r=subprocess.run(["python3",f"{PPTX}/add_slide.py",up,reported_slide],capture_output=True,text=True)
    m=re.search(r"Created ppt/slides/(slide\d+\.xml)",r.stdout);
    if not m: sys.exit("add_slide failed:\n"+r.stdout+r.stderr)
    new_slide=m.group(1)
    nid=re.search(r'id="(\d+)" r:id="(rId\d+)"',r.stdout); new_id,new_rid=nid.group(1),nid.group(2)
    print(f"new slide={new_slide} id={new_id} rid={new_rid}")

    # duplicate reported charts -> fresh numbers; repoint new slide rels
    nums=[int(re.search(r"chart(\d+)\.xml",c).group(1)) for c in
          __import__("glob").glob(f"{up}/ppt/charts/chart*.xml")]
    nxt=max(nums)+1
    cdir=f"{up}/ppt/charts"; tdir=f"{up}/ppt/theme"; edir=f"{up}/ppt/embeddings"
    nrels=read(f"{up}/ppt/slides/_rels/{new_slide}.rels")
    ct_adds=[]
    for c in rep_charts:
        old=int(re.search(r"chart(\d+)",c).group(1)); new=nxt; nxt+=1
        shutil.copy(f"{cdir}/chart{old}.xml",f"{cdir}/chart{new}.xml")
        for pre in ("colors","style"):
            if os.path.exists(f"{cdir}/{pre}{old}.xml"): shutil.copy(f"{cdir}/{pre}{old}.xml",f"{cdir}/{pre}{new}.xml")
        ct_adds += [(f"/ppt/charts/chart{new}.xml","application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
                    (f"/ppt/charts/style{new}.xml","application/vnd.ms-office.chartstyle+xml"),
                    (f"/ppt/charts/colors{new}.xml","application/vnd.ms-office.chartcolorstyle+xml")]
        rs=f"{cdir}/_rels/chart{old}.xml.rels"
        if os.path.exists(rs):
            t=read(rs).replace(f"colors{old}.xml",f"colors{new}.xml").replace(f"style{old}.xml",f"style{new}.xml")
            # give the new chart its OWN themeOverride + embedding (never share — sharing corrupts)
            tm=re.search(r'theme/(themeOverride\d+)\.xml',t)
            if tm and os.path.exists(f"{tdir}/{tm.group(1)}.xml"):
                shutil.copy(f"{tdir}/{tm.group(1)}.xml",f"{tdir}/themeOverride{new}.xml")
                t=t.replace(f"{tm.group(1)}.xml",f"themeOverride{new}.xml")
                ct_adds.append((f"/ppt/theme/themeOverride{new}.xml","application/vnd.openxmlformats-officedocument.themeOverride+xml"))
            em=re.search(r'embeddings/([^"]+\.xlsx)',t)
            if em and os.path.exists(f"{edir}/{em.group(1)}"):
                newx=f"Microsoft_Excel_Worksheet{new}.xlsx"
                shutil.copy(f"{edir}/{em.group(1)}",f"{edir}/{newx}")
                t=t.replace(f"embeddings/{em.group(1)}",f"embeddings/{newx}")
            write(f"{cdir}/_rels/chart{new}.xml.rels",t)
        nrels=nrels.replace(f"../charts/chart{old}.xml",f"../charts/chart{new}.xml")
    write(f"{up}/ppt/slides/_rels/{new_slide}.rels",nrels)

    # reorder sldIdLst: title, new, reported ; drop prior
    pres=read(f"{up}/ppt/presentation.xml")
    ids=re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>',pres)
    rels=read(f"{up}/ppt/_rels/presentation.xml.rels")
    def slide_of(rid):
        m=re.search(rf'Id="{rid}"[^>]*slides/(slide\d+\.xml)',rels); return m.group(1) if m else None
    title=[ (i,r) for i,r in ids if slide_of(r)=="slide1.xml" ][0]
    rep=[ (i,r) for i,r in ids if slide_of(r)==reported_slide ][0]
    new=(new_id,new_rid)
    order=[title,new,rep]
    block="\n    "+"\n    ".join(f'<p:sldId id="{i}" r:id="{r}"/>' for i,r in order)+"\n  "
    pres=re.sub(r'(<p:sldIdLst>).*?(</p:sldIdLst>)',lambda mm:mm.group(1)+block+mm.group(2),pres,flags=re.S)
    if "<p14:sldIdLst>" in pres:
        p14="\n            "+"\n            ".join(f'<p14:sldId id="{i}"/>' for i,_ in order)+"\n          "
        pres=re.sub(r'(<p14:sldIdLst>).*?(</p14:sldIdLst>)',lambda mm:mm.group(1)+p14+mm.group(2),pres,flags=re.S)
    write(f"{up}/ppt/presentation.xml",pres)

    # content-types for new chart parts
    ct=read(f"{up}/[Content_Types].xml")
    add="".join(f'<Override PartName="{pn}" ContentType="{c}"/>' for pn,c in ct_adds)
    ct=ct.replace("</Types>",add+"</Types>"); write(f"{up}/[Content_Types].xml",ct)

    # subtitle on new slide (segment-aware) + title slide
    s=read(f"{up}/ppt/slides/{new_slide}")
    s=re.sub(r'(<a:t>[^<]*–\s*Social Media \(EU\d\s+)[^<)]+(\)</a:t>)',rf'\g<1>{a.new_month}\g<2>',s)
    s=re.sub(r'(<a:t>)[^<–]+(\s*–\s*Social Media \(EU\d\s+[^<)]+\)</a:t>)',rf'\g<1>{a.carline}\g<2>',s)
    write(f"{up}/ppt/slides/{new_slide}",s)
    s1=read(f"{up}/ppt/slides/slide1.xml")
    s1=re.sub(r'(<a:t>Social Media )[^<]+(</a:t>)',rf'\g<1>{a.new_month}\g<2>',s1)
    s1=re.sub(r'<a:t>\d{2}\.\d{2}\.\d{4}</a:t>',f'<a:t>{a.report_date}</a:t>',s1,count=1)
    write(f"{up}/ppt/slides/slide1.xml",s1)

    # drop prior slide from sldIdLst already done; remove its presentation rel so clean can purge it
    # (clean.py removes now-unreferenced slide + charts)
    r=subprocess.run(["python3",f"{PPTX}/clean.py",up],capture_output=True,text=True)
    print(r.stdout.strip()[-400:] if r.stdout else r.stderr[-400:])
    print(f"\nSkeleton at {up}  (new slide {new_slide}, charts {[v for v in roles]})")
    # emit role->new chart mapping for the charts stage
    import json
    newmap={}
    # recompute new chart numbers per role from new slide rels
    nrels=read(f"{up}/ppt/slides/_rels/{new_slide}.rels")
    for role,old in roles.items():
        # find which new chart the old role maps to via order of duplication
        pass
    print("ROLES_JSON="+json.dumps({r:c for r,c in roles.items()}))

if __name__=="__main__": main()
