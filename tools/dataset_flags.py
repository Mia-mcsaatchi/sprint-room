#!/usr/bin/env python3
"""
dataset_flags.py — Pre-build sweep of a raw Brandwatch export: "is there anything worth
flagging before I build the slide?"

Run this BEFORE `menu`. It does no judgement; it lays out, deterministically, everything the
analyst needs to make the Phase 1 / Phase 3 calls (boolean noise, sentiment retags, real vs
passing-mention drivers, competitor movers) in one markdown file:

  1. Volume MoM per model (rows and unique URLs — they should match; if not, say so)
  2. Platform (Page Type Name) MoM per model
  3. Sentiment counts + % MoM per model (as exported — BEFORE retag)
  4. Carline country breakdown (all / news), MoM
  5. Story clusters per model (news+blog by normalised title), both months
  6. Carline daily volume — the peaks to click into in Brandwatch
  7. Every carline positive + negative row for the new month, with text — the retag list
  8. Noise scan — carline rows matching the carline's `noise_patterns` from config
     (e.g. Capri the island, the classic 1970s car, GTA 6's "Capri" city). Each is a
     candidate boolean exclusion, NOT an automatic drop (£100 rule — the analyst decides).
  9. Carline news rows split into "title names the carline" vs "passing mention"

Usage:
    python3 dataset_flags.py --carline Capri --raw raw.xlsx --new-month 8 --prior-month 7 \
        --out dataset_flags.md [--config monthly-run/config/carlines.json] [--min-cluster 3]
"""
import argparse, json, os, re, sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bw_read import load_dataframe  # noqa: E402

MONTHS = ["", "January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


def norm_title(t):
    t = re.sub(r"[^\w\s]", " ", str(t).lower())
    return re.sub(r"\s+", " ", t).strip()[:80]


def pct(n, d):
    return f"{n / d * 100:.1f}%" if d else "–"


def snip(s, n=150):
    return re.sub(r"\s+", " ", str(s))[:n]


def resolve_carline_query(df, carline):
    for q in df["Query Name"].unique():
        if carline.lower() in str(q).lower():
            return q
    sys.exit(f"Carline {carline!r} not found in Query Name values: {sorted(df['Query Name'].unique())}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carline", required=True)
    ap.add_argument("--raw", required=True)
    ap.add_argument("--new-month", type=int, required=True)
    ap.add_argument("--prior-month", type=int, required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                     "..", "monthly-run", "config", "carlines.json"))
    ap.add_argument("--min-cluster", type=int, default=3)
    a = ap.parse_args()

    df = load_dataframe(a.raw)
    if df.empty:
        sys.exit("No data rows found — is this a Bulk Mentions export with 'Query Name' in column B?")
    ptn_col = "Page Type Name" if "Page Type Name" in df.columns else "Page Type"
    for c in ("Full Text", "Title", "Country", "Domain", "Url", "Sentiment", "Page Type", ptn_col):
        if c not in df.columns:
            df[c] = ""
    df["txt"] = (df["Title"] + " " + df["Full Text"]).str.replace(r"\s+", " ", regex=True)
    df["norm"] = df["Title"].map(norm_title)

    noise_patterns = []
    try:
        cfg = json.load(open(a.config))
        noise_patterns = cfg.get(a.carline, {}).get("noise_patterns", [])
    except (OSError, json.JSONDecodeError):
        pass

    NM, PM = a.new_month, a.prior_month
    new, prior = df[df.Month == NM], df[df.Month == PM]
    cq = resolve_carline_query(df, a.carline)
    cn, cp = new[new["Query Name"] == cq], prior[prior["Query Name"] == cq]
    models = list(df["Query Name"].unique())
    L = []
    w = L.append

    w(f"# {a.carline} — dataset flags, {MONTHS[NM]} vs {MONTHS[PM]}\n")
    w(f"Source: `{os.path.basename(a.raw)}` — {len(df)} rows, {len(models)} queries: "
      f"{', '.join(sorted(models))}.  ")
    stray = df[~df.Month.isin([NM, PM])]
    if len(stray):
        w(f"⚠️ {len(stray)} rows fall outside months {PM}/{NM} (after UTC→local bucketing): "
          f"{stray.groupby('Month').size().to_dict()}.  ")
    dup = df.duplicated(["Query Name", "Url"]).sum()
    w(f"Duplicate (query, URL) rows: **{dup}** — volume = COUNTUNIQUE(URL), so rows ≠ URLs if this is > 0.\n")

    # 1. volume
    w("## 1. Volume MoM (unique URLs)\n")
    w("| Model | Prior | New | Δ | Δ% | rows≠URLs? |\n|---|---|---|---|---|---|")
    rows = []
    for q in models:
        p, n = prior[prior["Query Name"] == q], new[new["Query Name"] == q]
        pu, nu = p["Url"].nunique(), n["Url"].nunique()
        rows.append((q, pu, nu, nu - pu, (nu - pu) / pu * 100 if pu else 0.0,
                     "" if (len(p), len(n)) == (pu, nu) else f"rows {len(p)}/{len(n)}"))
    rows.sort(key=lambda r: r[3], reverse=True)
    for q, pu, nu, d, dp, flag in rows:
        star = " ⚠️" if abs(dp) >= 30 else ""
        w(f"| {'**' + q + '**' if q == cq else q} | {pu} | {nu} | {d:+d} | {dp:+.1f}%{star} | {flag} |")
    w("\n⚠️ = ±30% or more MoM — a candidate 'major change' for bullet 1 / notes.\n")

    # 2. platform
    w("## 2. Platform MoM (rows by Page Type Name)\n")
    w("| Model | Platform | Prior | New | Δ |\n|---|---|---|---|---|")
    for q in models:
        sub = df[df["Query Name"] == q]
        pv = sub[sub.Month == PM][ptn_col].value_counts()
        nv = sub[sub.Month == NM][ptn_col].value_counts()
        for plat in sorted(set(pv.index) | set(nv.index), key=lambda k: -(nv.get(k, 0) + pv.get(k, 0))):
            w(f"| {q} | {plat} | {pv.get(plat, 0)} | {nv.get(plat, 0)} | {nv.get(plat, 0) - pv.get(plat, 0):+d} |")
    w("")

    # 3. sentiment
    w("## 3. Sentiment MoM — AS EXPORTED (before any retag)\n")
    w("| Model | Month | n | pos | pos% | neg | neg% | neu |\n|---|---|---|---|---|---|---|---|")
    for q in models:
        for label, sub in ((MONTHS[PM], prior[prior["Query Name"] == q]), (MONTHS[NM], new[new["Query Name"] == q])):
            vc = sub["Sentiment"].value_counts()
            n = len(sub)
            w(f"| {q} | {label} | {n} | {vc.get('positive', 0)} | {pct(vc.get('positive', 0), n)} | "
              f"{vc.get('negative', 0)} | {pct(vc.get('negative', 0), n)} | {vc.get('neutral', 0)} |")
    zero = [q for q in models if len(new[(new["Query Name"] == q) & (new.Sentiment != "neutral")]) == 0]
    if zero:
        w(f"\n⚠️ Models with **zero** positive and negative mentions in {MONTHS[NM]}: {', '.join(zero)} "
          f"— the sentiment bar will be 100% neutral; that is the data, not a chart bug.")
    w("")

    # 4. country
    w(f"## 4. {a.carline} country breakdown (all / news)\n")
    w("| Country | Prior all | New all | Δ | Prior news | New news | Δ |\n|---|---|---|---|---|---|---|")
    ctry = Counter(cn["Country"]) + Counter(cp["Country"])
    for c, _ in ctry.most_common(12):
        pa, na = (cp["Country"] == c).sum(), (cn["Country"] == c).sum()
        pn = ((cp["Country"] == c) & (cp["Page Type"] == "news")).sum()
        nn = ((cn["Country"] == c) & (cn["Page Type"] == "news")).sum()
        w(f"| {c or '(blank)'} | {pa} | {na} | {na - pa:+d} | {pn} | {nn} | {nn - pn:+d} |")
    w("\nBreak a story down by country in the notes if it has > ~30 mentions in one country.\n")

    # 5. clusters
    w(f"## 5. Story clusters (news + blog, ≥{a.min_cluster} same-title mentions)\n")
    w("Read the Full Text before calling a cluster noise (£100 rule). A cluster whose title names a "
      "different model is usually a syndicated piece that *mentions* this model = passing mention → notes only.\n")
    for q in models:
        for label, sub in ((MONTHS[NM], new), (MONTHS[PM], prior)):
            s = sub[(sub["Query Name"] == q) & sub["Page Type"].isin(["news", "blog"])]
            vc = s["norm"].value_counts()
            sig = vc[vc >= a.min_cluster]
            if len(sig) == 0:
                continue
            w(f"### {q} — {label} ({len(s)} news+blog rows)\n")
            for nrm, c in sig.items():
                g = s[s["norm"] == nrm]
                r0 = g.iloc[0]
                w(f"- **[{c}]** {snip(r0['Title'], 110)}  ")
                w(f"  sentiment {dict(g['Sentiment'].value_counts())} · countries "
                  f"{dict(g['Country'].value_counts().head(3))} · e.g. <{r0['Url']}>  ")
                w(f"  _{snip(r0['Full Text'], 160)}_")
            w("")

    # 6. daily peaks
    w(f"## 6. {a.carline} daily volume, {MONTHS[NM]} (click into the peaks in Brandwatch)\n")
    daily = cn.groupby(cn["Date"].dt.day).size()
    if len(daily):
        thr = max(daily.mean() * 2, 5)
        w("| Day | Mentions | Top domains |\n|---|---|---|")
        for day, n in daily.items():
            top = dict(cn[cn["Date"].dt.day == day]["Domain"].value_counts().head(3))
            w(f"| {day:02d}{' ⚠️' if n >= thr else ''} | {n} | {top} |")
    w("")

    # 7. retag list
    w(f"## 7. {a.carline} — every positive / negative row in {MONTHS[NM]} (retag review)\n")
    w("Rules: passing / multi-model / other-model items → **neutral**; clearly wrong polarity → fix; "
      "delete only blatant irrelevance. Every change must be made in Brandwatch too.\n")
    for sent in ("negative", "positive"):
        s = cn[cn["Sentiment"] == sent].sort_values("Date")
        w(f"### {sent} ({len(s)})\n")
        for _, r in s.iterrows():
            w(f"- {r['Date'].date()} · {r['Page Type']} · {r['Country']} · {r['Domain']} · <{r['Url']}>  ")
            w(f"  {snip(r['Title'], 100)} — _{snip(r['Full Text'], 200)}_")
        w("")

    # 8. noise
    w(f"## 8. Noise scan — {a.carline} rows matching config `noise_patterns`\n")
    if not noise_patterns:
        w("_No `noise_patterns` configured for this carline in carlines.json — add the phrases you "
          "exclude in Brandwatch so this section works next cycle._\n")
    else:
        pat = re.compile("|".join(noise_patterns), re.I)
        hits = df[(df["Query Name"] == cq) & df["txt"].str.contains(pat)]
        by = hits.groupby(["Month", "Page Type"]).size().to_dict()
        w(f"{len(hits)} rows matched ({by}). Candidate boolean exclusions — check each, don't drop blind.\n")
        for _, r in hits.sort_values(["Month", "Date"]).iterrows():
            m = pat.search(r["txt"])
            w(f"- M{int(r['Month'])} {r['Date'].date()} · {r['Page Type']} · {r['Sentiment']} · {r['Domain']} · "
              f"matched `{m.group(0) if m else ''}` · <{r['Url']}>  ")
            w(f"  _{snip(r['txt'], 140)}_")
        clean_n = cn[~cn.index.isin(hits.index)]
        clean_p = cp[~cp.index.isin(hits.index)]
        w(f"\nIf all matches are excluded: {a.carline} {MONTHS[PM]} {len(cp)} → **{len(clean_p)}**, "
          f"{MONTHS[NM]} {len(cn)} → **{len(clean_n)}** "
          f"(Δ {len(clean_n) - len(clean_p):+d}, {pct(len(clean_n) - len(clean_p), len(clean_p))}).\n")

    # 9. focused vs passing news
    w(f"## 9. {a.carline} news+blog in {MONTHS[NM]}: title names the carline vs passing mention\n")
    nb = cn[cn["Page Type"].isin(["news", "blog"])]
    focused = nb[nb["Title"].str.contains(a.carline, case=False)]
    w(f"{len(nb)} news+blog rows → **{len(focused)} carline-titled**, {len(nb) - len(focused)} passing/other-title.\n")
    w("### Carline-titled\n")
    for _, r in focused.sort_values("Date").iterrows():
        w(f"- {r['Date'].date()} · {r['Country']} · {r['Sentiment']} · [{snip(r['Title'], 100)}]({r['Url']})")
    w("\n### Other-title (check: passing mention → notes, or a real driver with a different headline)\n")
    for nrm, c in nb[~nb.index.isin(focused.index)]["norm"].value_counts().items():
        g = nb[nb["norm"] == nrm].iloc[0]
        w(f"- [{c}] {snip(g['Title'], 100)} · {g['Domain']} — _{snip(g['Full Text'], 120)}_")
    w("")

    with open(a.out, "w") as f:
        f.write("\n".join(L))
    print(f"Wrote {a.out}  ({len(L)} lines)")


if __name__ == "__main__":
    main()
