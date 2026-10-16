#!/usr/bin/env python3
"""
Extract drivers and produce a bullet menu draft for analyst review.

For each competitor, clusters news/blog titles, dedupes, and identifies the top
stories driving volume. Computes sentiment % by page type. Surfaces MoM deltas.

Output is a markdown file the analyst reviews before bullets are written into the deck.

Usage:
    python3 extract_drivers.py --carline Capri --raw raw.xlsx \\
        --new-month 4 --prior-month 3 --out menu.md
"""
import argparse, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from collections import defaultdict

def load_raw(xlsx_path):
    from raw_loader import load_raw as _load
    df = _load(xlsx_path)
    if 'Full Text' not in df.columns:
        df['Full Text'] = ''
    return df

def norm_title(t):
    t = re.sub(r'[^a-z0-9\s]', ' ', t.lower())
    return re.sub(r'\s+', ' ', t).strip()[:80]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--carline', required=True)
    ap.add_argument('--raw', required=True)
    ap.add_argument('--new-month', type=int, required=True)
    ap.add_argument('--prior-month', type=int, required=True)
    ap.add_argument('--out', required=True)
    args = ap.parse_args()

    df = load_raw(args.raw)
    import pandas as pd

    new = df[df['Month']==args.new_month]
    prior = df[df['Month']==args.prior_month]

    # Resolve carline query name (user gave us short name, find the actual query)
    queries = df['Query Name'].unique()
    carline_query = None
    for q in queries:
        if args.carline.lower() in str(q).lower():
            carline_query = q
            break
    if not carline_query:
        print(f"Could not find query for carline '{args.carline}' in: {sorted(queries)}")
        return

    out_lines = []
    out_lines.append(f"# {args.carline} — Driver Menu, Month {args.new_month} vs {args.prior_month}\n")

    # === Volume MoM ===
    out_lines.append("## Volume MoM (sorted by Δ)\n")
    out_lines.append("| Model | Prior | New | Δ | Δ% |")
    out_lines.append("|---|---|---|---|---|")
    rows = []
    for q in queries:
        n = len(new[new['Query Name']==q])
        p = len(prior[prior['Query Name']==q])
        d = n - p
        dpct = d/p*100 if p else 0
        rows.append((q, p, n, d, dpct))
    rows.sort(key=lambda r: abs(r[3]), reverse=True)
    for q,p,n,d,dpct in rows:
        out_lines.append(f"| {q} | {p} | {n} | {d:+d} | {dpct:+.1f}% |")
    out_lines.append("")

    # === Sentiment MoM ===
    out_lines.append("## Sentiment % MoM\n")
    out_lines.append("| Model | Pos new | Pos prior | Δpp | Neg new | Neg prior | Δpp |")
    out_lines.append("|---|---|---|---|---|---|---|")
    for q in queries:
        np = new[new['Query Name']==q]
        pp = prior[prior['Query Name']==q]
        if len(np)==0 or len(pp)==0: continue
        np_p = (np['Sentiment']=='positive').mean()*100
        np_n = (np['Sentiment']=='negative').mean()*100
        pp_p = (pp['Sentiment']=='positive').mean()*100
        pp_n = (pp['Sentiment']=='negative').mean()*100
        out_lines.append(f"| {q} | {np_p:.1f}% | {pp_p:.1f}% | {np_p-pp_p:+.1f} | {np_n:.1f}% | {pp_n:.1f}% | {np_n-pp_n:+.1f} |")
    out_lines.append("")

    # === Top drivers per model (news+blog clusters with ≥5 mentions) ===
    out_lines.append("## Top story clusters (news + blog, ≥5 mentions)\n")
    out_lines.append("Use ≥10 as the threshold for putting on slide; 5–9 is notes-only.\n")
    for q in queries:
        sub = new[(new['Query Name']==q) & (new['Page Type'].isin(['news','blog']))]
        if len(sub)==0: continue
        sub = sub.copy()
        sub['norm'] = sub['Title'].apply(norm_title)
        clusters = sub['norm'].value_counts()
        sig = clusters[clusters>=5]
        if len(sig)==0: continue
        out_lines.append(f"### {q} ({len(sub)} news+blog rows)\n")
        for n, c in sig.items():
            s = sub[sub['norm']==n].iloc[0]
            sentiment_mix = sub[sub['norm']==n]['Sentiment'].value_counts().to_dict()
            country_mix = sub[sub['norm']==n]['Country'].value_counts().head(3).to_dict()
            full_text_sample = (sub[sub['norm']==n]['Full Text'].iloc[0][:200] if 'Full Text' in sub.columns else '').strip()
            out_lines.append(f"- **[{c} mentions]** {s['Title'][:120]}")
            out_lines.append(f"  - URL: `{s['Url']}`")
            out_lines.append(f"  - Sentiment mix: {sentiment_mix}")
            out_lines.append(f"  - Top countries: {country_mix}")
            if full_text_sample:
                out_lines.append(f"  - Full text excerpt: *{full_text_sample}*")
        out_lines.append("")

    # === Per-page-type breakdown for carline (Insight 3 source data) ===
    out_lines.append(f"## {args.carline} sentiment by page type (Insight 3 notes)\n")
    cn = new[new['Query Name']==carline_query]
    for sent in ['positive','negative','neutral']:
        pt = cn[cn['Sentiment']==sent]['Page Type'].value_counts().to_dict()
        out_lines.append(f"- **{sent}** ({len(cn[cn['Sentiment']==sent])} total): {pt}")
    out_lines.append("")

    # === Country breakdown for carline (notes block) ===
    out_lines.append(f"## {args.carline} country breakdown (notes)\n")
    out_lines.append("| Country | New (all) | Δ (all) | New (news) | Δ (news) |")
    out_lines.append("|---|---|---|---|---|")
    cp = prior[prior['Query Name']==carline_query]
    for country in ['Germany','United Kingdom','Spain','France','Italy','Netherlands','Norway']:
        n_all = len(cn[cn['Country']==country])
        p_all = len(cp[cp['Country']==country])
        n_news = len(cn[(cn['Country']==country) & (cn['Page Type']=='news')])
        p_news = len(cp[(cp['Country']==country) & (cp['Page Type']=='news')])
        out_lines.append(f"| {country} | {n_all} | {n_all-p_all:+d} | {n_news} | {n_news-p_news:+d} |")
    out_lines.append("")

    # === Empty bullet menu skeleton for analyst to fill ===
    out_lines.append("---\n")
    out_lines.append("## DRAFT BULLET MENU — analyst to confirm/edit\n")
    out_lines.append("Use the data above to fill these in. Each named driver MUST have a mention count.\n"
                     "Keep the `## bullet1/2/3` and `## hyperlinks` headings exactly as-is - write_bullets.py parses them.\n")
    out_lines.append("## bullet1\n")
    out_lines.append("> Mentions volume {direction} for {scope} from {prior month}, with Ford " + args.carline + " {Δ}, alongside ...\n")
    out_lines.append("## bullet2\n")
    out_lines.append("> " + args.carline + "'s {direction} was driven by news ({Δ}). Top driver: [STORY] ({count} mentions, {region}). Secondary driver: ...\n")
    out_lines.append("## bullet3\n")
    out_lines.append(f"> {args.carline} sentiment ...\n")
    out_lines.append("## hyperlinks\n")
    out_lines.append("| linked phrase | URL |")
    out_lines.append("|---|---|")
    out_lines.append("| Top story name | https://... |")
    out_lines.append("")
    out_lines.append("---\n")
    out_lines.append("After analyst confirms, save this file in place and proceed to Stage 3.\n")

    with open(args.out, 'w') as f:
        f.write('\n'.join(out_lines))
    print(f"Wrote {args.out}")

if __name__ == '__main__':
    main()
