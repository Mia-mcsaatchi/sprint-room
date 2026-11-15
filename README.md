# Sprint Room — Ford LENS Social Media Reporting

Single source of truth for the monthly **Sprint Room** social-media reporting programme
(Ford carlines: **Capri, Explorer, Kuga, Puma**). This repo consolidates the skills,
process, analytical standards, calculations, and slide/chart configs so the programme
can be handed over or migrated without knowledge living only in someone's head.

> **Confidential / internal.** Contains McCann + Ford LENS process detail and Brandwatch
> data structure. Keep this repository **private**.

## What's here

| Path | What it is |
|------|-----------|
| [`monthly-run/`](monthly-run/) | **Run this monthly.** Orchestrator (`run.py`) + [`RUNBOOK.md`](monthly-run/RUNBOOK.md) + per-carline [`config`](monthly-run/config/carlines.json) that drives a full cycle |
| [`docs/`](docs/) | The knowledge base — process, standards, calculations, slide configs, tooling patterns, data dictionary, glossary |
| [`skills/`](skills/) | The automation **skills** verbatim (SKILL.md + scripts + references) — the engine `monthly-run` calls |
| [`process/`](process/) | Source-of-truth documents: the 6-phase process doc + the automation architecture PDF |
| [`samples/`](samples/) | Small fixtures showing **table structure** (not datasets): the four chart-data exports + a structure-only raw-export header |
| [`tools/`](tools/) | Reusable utilities: `bw_read.py` (reliable Brandwatch parser), `rollforward_from_spec.py` (roll a delivered deck forward one month from a JSON spec, reusing the oldest slide's parts) |
| [`MIGRATION.md`](MIGRATION.md) | How to stand this up on a new machine / onboard a new analyst |

## Running a month (the point of this repo)

```bash
cd monthly-run
python3 run.py newmonth --carline Kuga --month 2026-07   # scaffold + list the 6 files to drop in
python3 run.py doctor   --carline Kuga --month 2026-07   # confirm inputs
python3 run.py validate --carline Kuga --month 2026-07
python3 run.py menu     --carline Kuga --month 2026-07   # ← approve the bullet menu (hard gate)
python3 run.py skeleton --carline Kuga --month 2026-07 --report-date 17.08.2026
python3 run.py fixskeleton --carline Kuga --month 2026-07
python3 run.py roles    --carline Kuga --month 2026-07   # config vs. real chart order
python3 run.py charts   --carline Kuga --month 2026-07
python3 run.py bullets  --carline Kuga --month 2026-07
python3 run.py embed    --carline Kuga --month 2026-07
python3 run.py pack     --carline Kuga --month 2026-07   # → the delivered .pptx
```

Full step-by-step (incl. the manual Brandwatch/Sheets phases) is in
[`monthly-run/RUNBOOK.md`](monthly-run/RUNBOOK.md). Client exports live under
`monthly-run/inputs/` and are git-ignored — **only table structure is committed, never the
monthly dataset.**

## The 30-second mental model

1. **Analyst** cleans Brandwatch queries + retags sentiment (Phases 1 & 3 — human judgement, not automated).
2. **Deterministic Python** aggregates volume, sentiment %, country/platform, and MoM deltas — *all quantitative facts are locked before any LLM is involved*.
3. **The skill** rebuilds the four charts, embeds editable backing xlsx, writes evidenced bullets with inline hyperlinks, and drops the oldest comparison slide.
4. **Analyst** reviews the bullet menu (hard gate), approves, and QAs the final deck before delivery to Val (CC Mike, Seb, Mia).

Read [`docs/workflow-phases.md`](docs/workflow-phases.md) first, then
[`docs/calculations.md`](docs/calculations.md) and [`docs/slide-configs.md`](docs/slide-configs.md).

## Carlines & competitor sets

| Carline | Segment | Competitor set |
|---------|---------|----------------|
| **Kuga** | EU5 | Dacia Bigster, Jaecoo 7, Hyundai Tucson, Peugeot 3008, VW Tiguan |
| **Puma** | EU5 | Peugeot 2008, Toyota Yaris Cross, Hyundai Kona, VW T-Cross |
| **Explorer** | EU7 | Skoda Enyaq, VW ID.4, Renault Megane E-Tech, Peugeot E-3008, Volvo XC40 Recharge/EX40 |
| **Capri** | EU7 | Peugeot E-3008, VW ID.5, Skoda Enyaq Coupe, Polestar 2, Audi Q4 Sportback e-tron, Kia EV6 |

> Competitor counts differ (Puma tracks five, not six). Never assume the Explorer set —
> read the actual competitor names from the deck's `chart3.xml` `<c:cat>` strings each cycle.

## Skills index

| Skill | Purpose | In scope for this repo |
|-------|---------|------------------------|
| `sprint-room-monthly-update` | Roll a carline deck forward one month (charts, bullets, backing xlsx) | ✅ core |
| `voc-report` | Convert weekly EU Voice-of-Customer text report → clean `.docx` | ✅ related |
| `amazon-influencer:influencer-slicer-qa` | QA/repair the Amazon Influencer Slicer `.xlsm` | ⚠️ **separate project** — not included here by default |

The Amazon Influencer slicer skill belongs to a different client/workstream. It's left out
so this repo stays a clean Sprint Room handover. It can be added here or given its own repo —
see [`skills/README.md`](skills/README.md).
