# Running the monthly Sprint Room with Claude

How to run one carline for one month in Claude Code: clean the Brandwatch data, write the
insights, and roll last month's deck forward. Claude does the analysis and the build; you approve
the judgement calls and make the changes in Brandwatch.

Offline copies of this guide (PDF / HTML) can be generated from this file; the content is the same.

## Set up the session

| Setting | Value | Why |
|---|---|---|
| Repository | `mia-mcsaatchi/sprint-room` | Connect it when you start a Claude Code session. You need access to the repo on GitHub. |
| Model | Opus 5.5 | Pick it in the model menu before your first message. |
| Effort | High | Noise screening and insight writing need the extra reasoning. |
| Permission mode | Auto, or Accept edits | Auto runs long steps in the background without prompts. Accept edits asks before running commands. |

Client exports and decks never go into the repo, because `.gitignore` keeps them out
(`monthly-run/inputs/`, `monthly-run/work/`). Download every deliverable before you close the
session: files in the session's scratch space are deleted when it ends.

## Each month, per carline

### 1. Load the repo and confirm the job

Claude reads the runbook, standards and template rules, then says what it will do this month and
which files it needs.

```text
Load the repo and summarise what you need to do in this session. We're doing the [CARLINE] Sprint Room for [MONTH YEAR]. List the files you need from me.
```

### 2. Export the raw data from Brandwatch (you)

- Use Data Download, in **Excel** format, not CSV.
- Include **all six queries** (the carline and its competitors) and cover the **reporting month
  plus the previous month**, so month-on-month changes can be calculated.
- Upload the file to the session. Don't worry about the file name: Brandwatch names files after
  whichever query comes first, and Claude identifies queries from the data itself.

### 3. Clean the data and pull out the noise

Claude sorts every irrelevant mention into a deletion list you can act on in Brandwatch, and
drafts boolean exclusions for the query.

```text
Here's the raw export for [MONTH]. Clean this month's data and extract the noise. Give me a workbook with three tabs:
- DELETE: certain noise (the £100 rule)
- REVIEW: unsure, for me to check
- RETAG NEUTRAL: passing mentions, which must be retagged, not deleted
On every row include the URL, the reason, and a "Search snippet" column: a short exact phrase from the mention I can paste into the Brandwatch Mention List search box to find it and delete it.
Also draft boolean exclusions per query to stop this noise coming back. Guard them so they can't remove genuine content, and test them against the export.
```

Then you, in Brandwatch:

- Paste each **Search snippet** into the Mention List search box, then delete the mention, or
  retag it neutral for the RETAG tab.
- Paste the exclusions into the queries, in one batch, so Brandwatch only has to refresh once.
- Re-export, and ask Claude to **compare the before and after** pulls: anything genuine lost, any
  new noise let in.

### 4. Screen for insights: the Summary box (box 4)

Claude works out the volume, platform, country and sentiment changes, finds the stories behind
them, and proposes the three Summary bullets with mention counts and links.

```text
Screen the cleaned data for insights for the Summary box (box 4), following the repo's editorial standards. Give me a bullet menu with counts and links to approve first.
If nothing significant happened, write a short summary for the box and skip the rest.
If a story is big enough to deserve its own slide, tell me why and propose it. Build it per the repo only after I approve.
```

> **Approval step.** Nothing goes into the deck until you've approved the bullet menu. Check the
> stories, counts and links, and that passing mentions sit in the notes, not the bullets.

### 5. Build the deck from last month's slide

- Upload **last month's Sprint Room deck** and give Claude the report date.
- The template is **fixed**: Claude changes only the numbers, the chart data, the Summary bullets
  and the notes, never the layout or styling.
- Earlier months on the volume chart are copied **from last month's deck**. Brandwatch sometimes
  restates past months, for example after a query edit, and the deck must match what the client
  has already seen.

```text
Here's last month's Sprint Room deck. Build [MONTH YEAR] from it with the approved menu. Report date [DD.MM.YYYY].
Follow the template exactly: only update the numbers, the chart data, the Summary bullets and the notes. Copy earlier months from this deck as published. If Brandwatch's figures for those months don't match, flag it and explain it in the notes.
```

Then you, in PowerPoint, check that:

- the sentiment chart shows **6 diamonds** (the previous month);
- the title-slide date and the slide subtitle month are right;
- the links open, and the charts' "Edit Data" opens with the new figures.

### 6. QA and deliver (you)

- Every sentiment change must be made in **Brandwatch too**, not just in the chart. Val checks this.
- Upload the deck to the Trello card and email Val (CC Mike, Seb, Mia).

## Pitfalls we've hit

- **Over-broad exclusions remove real coverage.** Single everyday words in a NOT block, such as
  German `war` ("was"), French `occasion` or Spanish `a la venta`, quietly delete genuine posts.
  Ask Claude to check a query before you trust a low volume. See
  [`analytical-standards.md`](analytical-standards.md#query-hygiene).
- **Competitors share the exclusion block.** If you change it for one query, change it for all
  six in the same month, or share of voice stops being comparable.
- **Changing a query re-runs past months.** When you edit a query, Brandwatch back-fills earlier
  months, so they no longer match the published deck. Keep the published numbers on the chart,
  quote month-on-month changes on the same basis, and explain the change in the notes.
- **Passing mentions are not noise.** Retag them neutral and keep them out of the bullets. Delete
  only what you'd bet £100 on.
- **The Summary box grows to fit its text.** Keep the bullets about as long as last month's and put
  the overflow in the notes, or the box runs off the slide.
- **Preview quirks.** In Claude's LibreOffice preview the sentiment chart shows only 1–2 diamonds,
  and the line-chart end labels can overlap. Both display correctly in PowerPoint. If the preview
  won't open at all, ask Claude to install LibreOffice Impress in the session (see
  [`tooling-patterns.md`](tooling-patterns.md#rendering-in-a-fresh-cloud-session)).

## Files Claude will ask for

| File | From | Used for |
|---|---|---|
| Raw Bulk Mentions export (.xlsx) | Brandwatch, all 6 queries, reporting + previous month | Cleaning, volumes, sentiment, drivers |
| Re-export after cleaning | Brandwatch, same scope | Before/after check and final figures |
| Last month's Sprint Room deck (.pptx) | Trello card | Template, plus earlier months as published |
| Report date | You | Title slide |

## Current state (handover, October 2026)

- **Explorer September 2026** was built from the June 2026 deck (July and August decks were not
  delivered). June on the volume chart is carried from the published June deck; July to September
  come from the refined query.
- The **Explorer Brandwatch query was refined in September 2026** — see
  [`../monthly-run/config/queries/`](../monthly-run/config/queries/). The **competitor queries were
  not changed**; aligning their exclusions is the first job for October.
- Sentiment for September was used as exported (Phase 3 retag not yet applied in Brandwatch).

Source of truth: this repo (`monthly-run/RUNBOOK.md`, `docs/`). When something new trips you up,
ask Claude to add it to the relevant doc so the next person doesn't relearn it.
