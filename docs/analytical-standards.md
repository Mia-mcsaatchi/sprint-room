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

## Carry-forward

In a new month's build, the prior month's data is left entirely untouched and carried forward
from the *published* prior-month deck. Never re-derive last month's numbers from raw data.
