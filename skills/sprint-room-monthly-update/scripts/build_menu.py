#!/usr/bin/env python3
"""build_menu.py — evidence bullet-menu from CLEANED raw (raw-only path). Hard-gate artifact."""
import argparse, os, re, sys
from collections import defaultdict, Counter
sys.path.insert(0, os.path.dirname(__file__))
import clean_raw as cr

def norm(t): return re.sub(r"\s+"," ",re.sub(r"[^a-z0-9\s]"," ",(t or "").lower())).strip()[:80]

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--carline",required=True); ap.add_argument("--month",required=True)
    ap.add_argument("--raw",required=True); ap.add_argument("--also",nargs="*",default=[])
    ap.add_argument("--carline-query",default="Ford Capri"); ap.add_argument("--out",required=True)
    a=ap.parse_args()
    yr,mm=map(int,a.month.split("-")); prior=f"{yr-1 if mm==1 else yr}-{12 if mm==1 else mm-1:02d}"
    months=(prior,a.month); cq=a.carline_query
    rows=cr.load_merged(a.raw,a.also,months); cleaned,_=cr.apply_clean(rows,cq)

    vol=defaultdict(lambda:defaultdict(set)); sent=defaultdict(lambda:defaultdict(Counter))
    byq_mo=defaultdict(list)
    for r in cleaned:
        q,mo=r.get("Query Name"),cr.ym(r.get("Date"))
        vol[q][mo].add(r.get("Url")); sent[q][mo][(r.get("Sentiment") or "").strip()]+=1
        byq_mo[(q,mo)].append(r)
    L=[f"# {a.carline} {a.month} — Bullet menu (cleaned raw)  ⟵ HARD GATE\n"]

    L.append("## Volume MoM (unique URL, cleaned)\n\n| Model | Jun | Jul | Δ | Δ% |\n|---|---|---|---|---|")
    vr=[]
    for q in vol:
        p,n=len(vol[q][prior]),len(vol[q][a.month]); vr.append((q,p,n,n-p,(n-p)/p*100 if p else 0))
    for q,p,n,d,dp in sorted(vr,key=lambda x:-abs(x[3])):
        L.append(f"| {q} | {p} | {n} | {d:+d} | {dp:+.0f}% |")

    L.append("\n## Sentiment % MoM (share of model's mentions)\n\n| Model | Pos Jul | Pos Jun | Neg Jul | Neg Jun |\n|---|---|---|---|---|")
    for q in vol:
        cn,cp=sent[q][a.month],sent[q][prior]; tn,tp=sum(cn.values()),sum(cp.values())
        if not tn or not tp: continue
        L.append(f"| {q} | {100*cn['positive']/tn:.0f}% ({cn['positive']}) | {100*cp['positive']/tp:.0f}% | "
                 f"{100*cn['negative']/tn:.0f}% ({cn['negative']}) | {100*cp['negative']/tp:.0f}% |")

    L.append("\n## Top story clusters — news+blog, ≥5 mentions (≥30 = driver, <30 = cluster)\n")
    for q in vol:
        sub=[r for r in byq_mo[(q,a.month)] if r.get("Page Type") in ("news","blog")]
        cl=Counter(norm(r.get("Title")) for r in sub)
        sig=[(k,c) for k,c in cl.most_common() if c>=5 and k]
        if not sig: continue
        L.append(f"### {q} ({len(sub)} news+blog)\n")
        for k,c in sig:
            ex=[r for r in sub if norm(r.get("Title"))==k]
            smix=Counter((r.get('Sentiment') or '') for r in ex)
            cmix=Counter((r.get('Country') or '') for r in ex).most_common(3)
            L.append(f"- **[{c}]** {ex[0].get('Title')[:110]}")
            L.append(f"  - {dict(smix)} | {cmix} | {ex[0].get('Url')}")
        L.append("")

    L.append(f"## {a.carline} sentiment by platform (Jul) — Insight-3 notes\n")
    cn=byq_mo[(cq,a.month)]
    for s in ("positive","negative","neutral"):
        pt=Counter(r.get("Page Type") for r in cn if (r.get("Sentiment") or "")==s)
        L.append(f"- **{s}** ({sum(pt.values())}): {dict(pt)}")

    L.append(f"\n## {a.carline} country breakdown (Jul vs Jun)\n\n| Country | Jul | Jun | Δ |\n|---|---|---|---|")
    cp=byq_mo[(cq,prior)]
    for c in ("United Kingdom","Spain","Germany","France","Italy","Netherlands","Norway"):
        n=len({r.get('Url') for r in cn if r.get('Country')==c})
        p=len({r.get('Url') for r in cp if r.get('Country')==c})
        L.append(f"| {c} | {n} | {p} | {n-p:+d} |")

    L.append("\n---\n\n## DRAFT BULLETS — confirm/edit, then I build\n")
    L.append("_(placeholders below are filled from the locked numbers above; edit freely)_\n")
    L.append("### bullet1 (volume)\n> …\n\n### bullet2 (drivers)\n> …\n\n### bullet3 (sentiment)\n> …\n")
    L.append("### hyperlinks\n\n| linked phrase | URL |\n|---|---|\n| … | … |\n")

    os.makedirs(os.path.dirname(a.out),exist_ok=True)
    open(a.out,"w").write("\n".join(L))
    print(f"Wrote {a.out}\n"); print("\n".join(L[:40]))

if __name__=="__main__": main()
