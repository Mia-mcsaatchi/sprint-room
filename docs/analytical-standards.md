# Analytical Standards

The judgement rules that separate a real driver from noise. These are the decisions that make
the report credible; automation supplies evidence, but these calls are the analyst's.

## Passing mentions ≠ drivers

Coverage where the carline appears only incidentally (mentioned alongside several other brands,
or where a different model is the focus) is a **passing mention**. It belongs in **speaker notes
only**, never in bullets. Only genuinely carline-focused content earns driver attribution.

All passing mentions are tagged **neutral** — and Val checks that the retag shows up in both the
chart *and* Brandwatch.

## Fading prior-month content is not a "driver"

A story that is declining from last month is **not** a new driver. Only *new* clusters warrant
driver attribution. Declining stories are explained as **normalisation**, not called out as
drivers.

## The £100 rule

Don't hard-exclude a row unless you'd bet **£100** it's noise. Ambiguous rows stay in the file
and are flagged for analyst review rather than silently dropped.

> Cautionary example: the Explorer/Kuga "recall PR" cluster was almost written off as Kuga
> boolean noise. Reading the Full Text column showed it was a Stoll & Sauer law-firm press
> release about **recall 25S79** (high-pressure fuel-injection fire risk) affecting Explorer +
> Capri — genuine carline content. Read the text before dismissing a top cluster.

## Conservative retagging

- Obvious news mis-tags (passing/corporate mentions) → **neutral only**. Don't over-neutralise.
- If an item is clearly positive but tagged negative (or vice versa), correct it.
- Delete only blatantly irrelevant items.
- The **BW-cleaned export is authoritative** for sentiment when available — it supersedes any
  rules-based pass.

## Clusters

A **cluster** is a driver with volume **< 30**. When no negative driver reaches 30 mentions for
a model, report it with Val's exact phrase: *"no clusters of issues were identified as drivers"*.

## Phase 1 cleaning

Brandwatch export **filenames are unreliable**. Always validate the file by reading the
`Query Name` column distribution, not by trusting the filename. (`tools/bw_read.py --querydist`.)

Sort noise into three tiers, never one:

| Tier | What goes in it | Action in Brandwatch |
|---|---|---|
| **DELETE** | Certain noise — you'd bet £100 (e.g. Xiaomi "Explorer Edition", the French verb *explorer*, mining "junior explorers", other Renault E-Tech models in the Megane query) | Delete, and add a guarded exclusion |
| **REVIEW** | Brand present but model name not visible (often truncated Full Text) | Check by eye; most are genuine |
| **RETAG NEUTRAL** | Real content where the carline is only a passing reference (e.g. Skoda Peaq launch citing the Enyaq) | Retag neutral, never delete |

Give each DELETE row a short exact **search snippet** so it can be found from the Mention List
search box, plus its URL.

## Query hygiene

The Sept 2026 Explorer review found the query was losing genuine coverage, not just collecting
noise. Rules learned:

- **Never exclude single everyday words.** German `war` ("was"), French `occasion`, Spanish
  `a la venta` / `renting`, Italian `incidente` / `agenti`, English `dead` / `died` / `stocks`
  each appeared in ~0 surviving mentions — proof they were silently removing whole languages'
  worth of posts. Exclude **phrases** or `NEAR`-scoped pairs instead.
- **Guard every new noise exclusion** with `AND NOT ford` (or the model's brand) so a genuine
  article that happens to contain the noise term survives.
- **Test before saving:** run the proposed exclusions over the last raw export and confirm they hit
  0 genuine mentions. After saving, re-export and compare old vs new by URL (lost genuine / new noise).
- **Inclusion can be too narrow too.** Owner posts often lack an "electric" word near the model
  name; terms like `kWh`, `RWD`, `MEB`, `Akku`, `elbil`, `bereik` recover them.
- **All six queries share the exclusion block.** Change them together, in one month.
- **Editing a query back-fills earlier months** in Brandwatch. Keep published figures for prior
  months in the deck (carry-forward) and quote MoM on a like-for-like basis; say so in the notes.
- Back up the live query to [`../monthly-run/config/queries/`](../monthly-run/config/queries/)
  before editing — Brandwatch keeps no history.

## Carry-forward

In a new month's build, the prior month's data is left entirely untouched and carried forward
from the *published* prior-month deck. Never re-derive last month's numbers from raw data.
