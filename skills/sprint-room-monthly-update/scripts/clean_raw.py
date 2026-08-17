#!/usr/bin/env python3
"""
clean_raw.py — Post-export cleaning for the Sprint Room raw feed (raw-only path).

Applies the analyst-approved policy (£100 rule):
  DROP   = false-brand match (e.g. Capri = island/design/person) OR off-topic non-automotive
  RETAG  = genuine automotive but the tracked model is incidental (another tracked model /
           passing multi-brand) currently tagged positive/negative -> neutral (row KEPT)
  KEEP   = model-focused genuine content
  REVIEW = signals conflict / low confidence -> flagged, never auto-dropped

Merges multiple raw exports (some Brandwatch pulls drop a query) by adding only queries
missing from the primary file. Importable (load_merged / classify / apply_clean) and CLI.

CLI:
    python3 clean_raw.py --carline Capri --month 2026-07 \
        --raw primary.xlsx [--also extra1.xlsx ...] --out menu.md
"""
import argparse, json, os, re, sys
from collections import Counter, defaultdict
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "..", "tools"))
try:
    from bw_read import records
except ImportError:  # dev fallback
    sys.path.insert(0, "/root/sprint-room/tools"); from bw_read import records

# ---- keyword policy (shared base; Capri place-name noise is the main real case) ----
# ---- per-carline false-brand / off-topic keyword policy (false terms only apply to the carline) ----
# The carline name is a common word (Capri=island, Puma=sportswear/animal) so these disambiguate noise.
FALSE_BY_CARLINE = {
    "Ford Capri": [
        "isola di capri","capri island","isla de capri","anacapri"," amalfi","sorrento",
        " napoli"," naples","yacht","resort","moulay","khadija","lalla","living space",
        "interior design","interiordaily","baleine","whale","cuba","logistics platform",
        "fashionunited","michel cymes","nathalie marquay","aviazionecivile","fco-otp",
        "radici del","trustpilot","umbrellas banned","tour groups","group limits","gutteridge",
    ],
    "Ford Puma": [
        "sneaker","trainers","footwear"," nike"," adidas","puma se","sportswear","rihanna",
        "tracksuit","cleats","running shoe","cougar","mountain lion","wildcat","big cat",
        "helicopter","eurocopter","mtu powerpack","puma helicopter",
    ],
    "Explorer": [
        "internet explorer","file explorer","windows explorer","explorer.exe","microsoft edge",
        "dora the explorer","dora","boy scout","scout troop","polar explorer","space explorer",
        "ocean explorer","national geographic","natgeo","explorer academy","ie11",
    ],
    "Kuga": [
        # "Kuga" is largely Ford-specific; guard only against the few known homonyms
        "yuto kuga","kuga anime","kuga surname",
    ],
}
# subset that is non-automotive regardless of brand
OFFTOPIC_BY_CARLINE = {
    "Ford Capri": ["baleine","whale","cuba","logistics platform","fashionunited",
                   "aviaz","fco-otp","radici del","gutteridge","umbrellas banned"],
    "Ford Puma": ["helicopter","eurocopter","mtu powerpack","cougar","mountain lion","big cat"],
    "Explorer": ["internet explorer","file explorer","windows explorer","explorer.exe",
                 "dora the explorer","national geographic","natgeo","polar explorer","space explorer"],
}
# genuine-content overrides (never drop on a stray keyword)
GENUINE_BY_CARLINE = {
    "Ford Capri": ["ford adapta","capri 100% eléctric","capri 100% electric","nuevo capri","ford capri rwd"],
    "Ford Puma": ["ford puma","puma st","puma gen-e"],
    "Explorer": ["ford explorer","explorer ev","explorer bev"],
    "Kuga": ["ford kuga"],
}
# per-MODEL homonym/off-topic terms, applied to that model as a COMPETITOR too (not just when it
# is the carline). Kept tight so genuine car content is never dropped; the strong brand+model
# phrase (STRONG_KW) always overrides. Competitor queries are otherwise well-tuned upstream.
MODEL_FALSE = {
    "Hyundai Tucson": ["tucson arizona","tucson, az","university of arizona","tucson police",
                       "tucson unified","city of tucson","tucson weather","crossword"],
    "Hyundai Kona": ["kona coffee","kona brew","ironman","triathlon","kona ice","big island",
                     "kona hawaii"],
}
# analyst-approved specific calls, per carline (stable across re-exports)
FORCE_KEEP_BY_CARLINE  = {"Ford Capri": ["nani roma","más de 13.000","mas de 13.000","ford explorer y capri"]}
FORCE_RETAG_BY_CARLINE = {"Ford Capri": ["i drove 6 brand new cars","modelljahr 2027","alle radici","fco-otp","mazda mx-5","mazda mx5"]}
# WEAK token = what the boolean matched (appears even in false-brand hits, e.g. "capri" island)
MODEL_KW = {
    "Ford Capri": ["capri","ford capri"],
    "Peugeot E3008": ["3008","e-3008","e 3008"],
    "Polestar 2": ["polestar"],
    "Audi Q4 Sportback e-tron": ["q4 sportback","audi q4"],
    "Kia EV6": ["ev6","ev 6"],
    "VW ID.5": ["id.5","id5"],
    "Skoda Enyaq Coupe": ["enyaq"],
    "Ford Puma": ["puma"],
    "Peugeot 2008": ["2008","e-2008"],
    "Toyota Yaris Cross": ["yaris cross","yaris"],
    "Hyundai Kona": ["kona"],
    "VW T-Cross": ["t-cross","t cross"],
    "Explorer": ["explorer"],
    "Skoda Enyaq": ["enyaq"],
    "Megane E-Tech": ["megane","e-tech"],
    "VW id4": ["id4","id.4","id 4"],
    "Volvo XC40 Recharge/ EX40": ["xc40","ex40"],
    "Kuga": ["kuga"],
    "Dacia Bigster": ["bigster"],
    "Jaecoo 7": ["jaecoo"],
    "Hyundai Tucson": ["tucson"],
    "VW Tiguan": ["tiguan"],
    "Peugeot 3008": ["3008"],
}
# STRONG signal = disambiguating brand+model phrase; presence means it really IS the car
STRONG_KW = {
    "Ford Capri": ["ford capri"],
    "Peugeot E3008": ["peugeot"],
    "Polestar 2": ["polestar 2"],
    "Audi Q4 Sportback e-tron": ["audi q4","q4 sportback"],
    "Kia EV6": ["kia ev6"],
    "VW ID.5": ["vw id.5","volkswagen id.5","id.5"],
    "Skoda Enyaq Coupe": ["skoda enyaq","enyaq coupe"],
    "Ford Puma": ["ford puma","puma st","puma gen"],
    "Peugeot 2008": ["peugeot 2008","peugeot e-2008"],
    "Toyota Yaris Cross": ["toyota yaris"],
    "Hyundai Kona": ["hyundai kona"],
    "VW T-Cross": ["vw t-cross","volkswagen t-cross","t-cross"],
    "Explorer": ["ford explorer"],
    "Skoda Enyaq": ["skoda enyaq"],
    "Megane E-Tech": ["megane"],
    "VW id4": ["vw id.4","volkswagen id.4","id.4","id4"],
    "Volvo XC40 Recharge/ EX40": ["volvo xc40","xc40"],
    "Kuga": ["ford kuga"],
    "Dacia Bigster": ["dacia bigster"],
    "Jaecoo 7": ["jaecoo"],
    "Hyundai Tucson": ["hyundai tucson"],
    "VW Tiguan": ["vw tiguan","volkswagen tiguan","tiguan"],
    "Peugeot 3008": ["peugeot 3008"],
}
OTHER_FORD = ["ford focus","mustang","fiesta","ford capri"]

def ym(d): return d[:7] if d and re.match(r"\d{4}-\d{2}", d) else None

def load_merged(primary, extras, months):
    rows = [r for r in records(primary) if ym(r.get("Date")) in months]
    present = {r.get("Query Name") for r in rows}
    for ex in extras or []:
        for r in records(ex):
            if ym(r.get("Date")) in months and r.get("Query Name") not in present:
                rows.append(r)
    return rows

def classify(r, carline_q="Ford Capri", retag_passing=False):
    q     = r.get("Query Name") or ""
    title = (r.get("Title") or "").strip()
    full  = (r.get("Full Text") or "").strip()
    sent  = (r.get("Sentiment") or "").strip()
    domain = (r.get("Domain") or "").strip()
    txt, tl = (title + " " + full + " " + domain).lower(), title.lower()
    own = MODEL_KW.get(q, [q.lower()])
    strong = STRONG_KW.get(q, [q.lower()])
    own_in_title = any(k in tl for k in own)
    strong_in_txt = any(k in txt for k in strong)
    false_kws  = FALSE_BY_CARLINE.get(carline_q, [])
    offtopic_kws = OFFTOPIC_BY_CARLINE.get(carline_q, [])
    genuine_kws = GENUINE_BY_CARLINE.get(carline_q, [])
    force_keep  = FORCE_KEEP_BY_CARLINE.get(carline_q, [])
    force_retag = FORCE_RETAG_BY_CARLINE.get(carline_q, [])

    if any(k in txt for k in force_retag):
        return ("RETAG", f"approved: passing -> neutral") if sent in ("positive","negative") else ("KEEP","approved passing (neutral)")
    if any(k in txt for k in force_keep) or any(k in txt for k in genuine_kws):
        return "KEEP", "approved/genuine model content"

    # false-brand / off-topic keywords are name-specific to the carline (Capri island, Puma shoes);
    # competitors are clean, so only apply them to the carline to avoid false hits on genuine posts.
    false_hit = offtopic_hit = None
    if q == carline_q:
        false_hit    = next((k for k in false_kws if k in txt), None)
        offtopic_hit = next((k for k in offtopic_kws if k in txt), None)
    elif q in MODEL_FALSE:                      # competitor homonym/off-topic guard
        false_hit = next((k for k in MODEL_FALSE[q] if k in txt), None)
    # other tracked model (or other Ford) is the subject in the title, own model absent
    other = [m for m,kws in MODEL_KW.items() if m!=q and any(k in tl for k in kws)]
    other += [f for f in OTHER_FORD if f in tl]

    if (false_hit or offtopic_hit) and not strong_in_txt:
        return "DROP", f"false/off-topic: '{false_hit or offtopic_hit}'"
    if (false_hit or offtopic_hit) and strong_in_txt:
        return "REVIEW", f"false signal '{false_hit or offtopic_hit}' but model present"
    if other and not own_in_title:
        return ("RETAG", f"other-model piece ({other[0]}); {sent}->neutral") if sent in ("positive","negative") \
               else ("KEEP", "passing (other-model) already neutral")
    if own_in_title or strong_in_txt:
        return "KEEP", "model-focused"
    # only vet sentiment tags for the reported carline, and only for news/blog (social rarely names
    # the brand in the "title"); competitor raw tags stand (Brandwatch authority)
    if q == carline_q and sent in ("positive","negative") and (r.get("Page Type") in ("news","blog")):
        if retag_passing:
            return "RETAG", f"passing news/blog (no {carline_q} in title); {sent}->neutral"
        return "REVIEW", f"news/blog with no {carline_q} in title; {sent} tag needs a look"
    return "KEEP", "neutral passing"

def apply_clean(rows, carline_q="Ford Capri", retag_passing=False):
    """Return (cleaned_rows, actions) with DROPs removed and RETAGs set to neutral."""
    cleaned, actions = [], []
    for r in rows:
        action, reason = classify(r, carline_q, retag_passing)
        actions.append((r, action, reason))
        if action == "DROP":
            continue
        if action == "RETAG":
            r = dict(r); r["Sentiment"] = "neutral"
        cleaned.append(r)
    return cleaned, actions

def _vol_sent(rows):
    vol = defaultdict(lambda: defaultdict(set)); sent = defaultdict(lambda: defaultdict(Counter))
    for r in rows:
        q, mo = r.get("Query Name"), ym(r.get("Date"))
        vol[q][mo].add(r.get("Url"))
        sent[q][mo][(r.get("Sentiment") or "").strip()] += 1
    return vol, sent

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carline", required=True)
    ap.add_argument("--month", required=True, help="new month YYYY-MM")
    ap.add_argument("--raw", required=True)
    ap.add_argument("--also", nargs="*", default=[])
    ap.add_argument("--carline-query", default=None, help="raw Query Name for the carline (default: 'Ford <carline>')")
    ap.add_argument("--retag-passing", action="store_true", help="promote carline news/blog passing pos/neg to neutral (else REVIEW)")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    yr, mm = map(int, a.month.split("-"))
    prior = f"{yr-1 if mm==1 else yr}-{12 if mm==1 else mm-1:02d}"
    months = (prior, a.month)
    cq = a.carline_query or f"Ford {a.carline}"

    rows = load_merged(a.raw, a.also, months)
    cleaned, actions = apply_clean(rows, cq, a.retag_passing)

    buckets = defaultdict(list)
    for r, act, reason in actions:
        buckets[act].append((r, reason))
    print(f"Merged rows: {len(rows)} -> cleaned: {len(cleaned)}")
    for k in ("DROP","RETAG","REVIEW","KEEP"):
        print(f"  {k:7s}: {len(buckets[k])}")

    v0,s0 = _vol_sent(rows); v1,s1 = _vol_sent(cleaned)
    print(f"\nVolume (unique URL) {a.month}  raw -> cleaned:")
    for q in sorted(v1):
        print(f"  {q:30s} {len(v0[q][a.month]):>4} -> {len(v1[q][a.month]):>4}")
    print(f"\n{cq} sentiment raw -> cleaned:")
    for mo in (prior, a.month):
        c0,c1 = s0[cq][mo], s1[cq][mo]
        print(f"  {mo}: raw pos{c0.get('positive',0)}/neu{c0.get('neutral',0)}/neg{c0.get('negative',0)}"
              f"  -> clean pos{c1.get('positive',0)}/neu{c1.get('neutral',0)}/neg{c1.get('negative',0)}")

    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    def line(r, reason):
        t=(r.get("Title") or "").strip()[:70]
        return (f"- [{ym(r.get('Date'))} | {r.get('Query Name')} | {r.get('Sentiment')} | "
                f"{r.get('Page Type')}] {t} | {r.get('Domain')}  \n  ↳ {reason}  \n  {r.get('Url')}")
    with open(a.out,"w") as f:
        f.write(f"# {a.carline} {a.month} — Cleaning menu (all models)\n\n")
        f.write(f"Merged {len(rows)} → cleaned {len(cleaned)} — "
                f"DROP {len(buckets['DROP'])} · RETAG {len(buckets['RETAG'])} · "
                f"REVIEW {len(buckets['REVIEW'])} · KEEP {len(buckets['KEEP'])}\n")
        for k,title in [("DROP","Proposed DROP (removed)"),
                        ("RETAG","Proposed RETAG → neutral (kept)"),
                        ("REVIEW","Needs your call (not touched)")]:
            f.write(f"\n## {title} — {len(buckets[k])}\n\n")
            for r,reason in buckets[k]:
                f.write(line(r,reason)+"\n")
    print(f"\nMenu: {a.out}")

if __name__ == "__main__":
    main()
