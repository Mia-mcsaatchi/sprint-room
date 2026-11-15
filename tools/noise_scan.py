#!/usr/bin/env python3
"""
noise_scan.py - Phase 1 helper: surface likely boolean noise in a carline query.

Carline names are ordinary words (Explorer, Puma, Capri, Kuga is the exception), so
queries pick up Triumph "Rally Explorer" bikes, "MBR Explorer" spacecraft, Puma
trainers, Capri holidays... This scans the raw export for the carline's query and:

  1. splits each month into rows that mention a brand term (default: "ford") and
     rows that don't;
  2. for the no-brand rows, finds the words immediately around the carline name
     (e.g. "rally explorer", "explorer buds") and lists the contexts that repeat
     and never occur in brand rows - these are candidate NOT-terms for the
     Brandwatch boolean;
  3. writes every no-brand row to a CSV for the analyst to mark keep / exclude.

It FLAGS, it never excludes (the GBP 100 rule): rows like "zappi for our Explorer" are
genuine owner chatter without the word Ford. Exclusions are made in Brandwatch
(Phase 1), then the chart exports are re-pulled.

Usage:
    python3 tools/noise_scan.py RAW.xlsx --carline Explorer --out noise.csv
    python3 tools/noise_scan.py RAW.xlsx --carline Puma --brand-terms ford "puma gen-e"
"""
import argparse, csv, os, re, sys
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "skills", "sprint-room-monthly-update", "scripts"))
from _bw_raw import load_raw_df  # noqa: E402

WORD = r"[\wÀ-ÿ\-\.]+"


def contexts(text, term):
    """Bigram contexts around each occurrence of term: 'prev term' and 'term next'."""
    out = set()
    for m in re.finditer(rf"({WORD})?\W{{0,3}}\b{term}\b\W{{0,3}}({WORD})?", text, re.I):
        prev, nxt = (m.group(1) or "").lower(), (m.group(2) or "").lower()
        if prev:
            out.add(f"{prev} {term}")
        if nxt:
            out.add(f"{term} {nxt}")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("raw")
    ap.add_argument("--carline", required=True, help="carline name; also the term searched for")
    ap.add_argument("--brand-terms", nargs="+", default=["ford"],
                    help="a row containing any of these counts as on-brand (default: ford)")
    ap.add_argument("--min-repeat", type=int, default=2,
                    help="min no-brand rows for a context to be suggested (default 2)")
    ap.add_argument("--out", help="CSV of every no-brand row for analyst review")
    a = ap.parse_args()

    df = load_raw_df(a.raw)
    q = next((x for x in df["Query Name"].unique() if a.carline.lower() in x.lower()), None)
    if q is None:
        sys.exit(f"No query matching {a.carline!r}; got {sorted(df['Query Name'].unique())}")
    sub = df[df["Query Name"] == q].copy()
    term = re.escape(a.carline.lower())
    sub["text"] = (sub["Title"] + " \n " + sub["Full Text"]).str.lower()
    brand_re = "|".join(re.escape(b.lower()) for b in a.brand_terms)
    sub["on_brand"] = sub["text"].str.contains(brand_re, regex=True)
    sub["ctx"] = sub["text"].apply(lambda t: contexts(t, term))

    print(f"Query: {q}   brand terms: {a.brand_terms}")
    print(f"{'Month':>5} {'total':>6} {'brand':>6} {'no-brand':>9} {'no-brand %':>10}")
    for m, g in sub.groupby("Month"):
        nb = int((~g["on_brand"]).sum())
        print(f"{m:>5} {len(g):>6} {int(g['on_brand'].sum()):>6} {nb:>9} {nb/len(g)*100:>9.1f}%")

    brand_ctx = Counter(c for cs in sub.loc[sub["on_brand"], "ctx"] for c in cs)
    nb_rows = sub[~sub["on_brand"]]
    ctx_rows = defaultdict(list)
    for i, cs in nb_rows["ctx"].items():
        for c in cs:
            ctx_rows[c].append(i)
    # generic glue words make useless NOT-terms
    stop = {"the", "der", "die", "das", "den", "dem", "el", "la", "le", "les", "las", "los", "il",
            "de", "del", "di", "du", "und", "and", "y", "e", "et", "a", "of", "in", "en", "het",
            "een", "mit", "with", "con", "per", "for", "für", "my", "mein", "meinem", "our", "new",
            "neue", "neuen", "nuevo", "nouveau", "s", "is", "ist", "es", "un", "una", "ein", "eine"}
    cands = []
    for c, idx in ctx_rows.items():
        other = c.replace(a.carline.lower(), "").strip()
        if len(idx) >= a.min_repeat and brand_ctx[c] == 0 and other not in stop and not other.isdigit():
            months = Counter(nb_rows.loc[idx, "Month"])
            cands.append((len(idx), c, dict(sorted(months.items())),
                          nb_rows.loc[idx[0], "Title"][:80]))
    cands.sort(reverse=True)
    print(f"\nCandidate NOT-term contexts (>= {a.min_repeat} no-brand rows, 0 brand rows):")
    for n, c, months, title in cands[:40]:
        print(f"  {n:>4}  \"{c}\"  by month {months}  e.g. {title}")

    if a.out:
        cand_set = {c for _, c, _, _ in cands}
        with open(a.out, "w", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["Month", "Page Type", "Country", "Sentiment", "Domain", "Title", "Url",
                        "matched_contexts", "suggested", "analyst_decision"])
            for _, r in nb_rows.sort_values(["Month", "Domain"]).iterrows():
                hit = sorted(r["ctx"] & cand_set)
                w.writerow([r["Month"], r["Page Type"], r["Country"], r["Sentiment"], r.get("Domain", ""),
                            r["Title"][:200], r["Url"], "; ".join(sorted(r["ctx"]))[:200],
                            "repeated non-brand context: " + ", ".join(hit) if hit else "review", ""])
        print(f"\nWrote {len(nb_rows)} no-brand rows to {a.out} (fill analyst_decision: keep/exclude)")


if __name__ == "__main__":
    main()
