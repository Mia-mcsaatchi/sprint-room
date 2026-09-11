# Monthly Run — Runbook

The turnkey sequence for one carline, one month. The orchestrator ([`run.py`](run.py)) calls
the skill's scripts with the right arguments and per-carline config
([`config/carlines.json`](config/carlines.json)); this runbook is the checklist around it,
including the manual Brandwatch/Sheets steps that stay human.

> Every command takes `--carline <Capri|Explorer|Kuga|Puma> --month <YYYY-MM>`.
> Add `--dry-run` to any stage to see the exact command without running it.
> Point at the live skill instead of the repo copy with
> `export SKILL_DIR=/mnt/skills/user/sprint-room-monthly-update`.

## 0. Before the automated part (manual, in Brandwatch / Sheets)

These are analyst judgement and are **not** automated (see
[`../docs/workflow-phases.md`](../docs/workflow-phases.md)):

- **Phase 1** — clean the boolean queries (peak-check, add exclusions), all four carlines.
- **Phase 3** — retag sentiment in the Mention List (passing mentions → neutral; fix wrong
  tags), and make the same change in Brandwatch, not just the chart.
- Export the **four chart files** + one **raw export** covering both months.

## 1. Scaffold the cycle

```bash
cd monthly-run
python3 run.py newmonth --carline Kuga --month 2026-07
```

Drop the six exports into the folder it prints, renamed exactly:

```
inputs/Kuga/2026-07/
  last_month.pptx      raw.xlsx
  chart_volume.xlsx    chart_country.xlsx
  chart_platform.xlsx  chart_sentiment.xlsx
```

## 2. Check inputs & validate

```bash
python3 run.py doctor   --carline Kuga --month 2026-07   # all six present?
python3 run.py validate --carline Kuga --month 2026-07   # runs the skill validator
```

`validate` will flag the classic problems: the 4th volume column being the new month
(UTC/GMT quirk), missing competitors in the sentiment export, and a platform export that came
out **unfiltered** (`all`/`all`) — re-export that one filtered to the carline if flagged.

## 2b. Dataset flags sweep (needs only `raw.xlsx`)

```bash
python3 run.py flags --carline Kuga --month 2026-07
```

Writes `work/Kuga/2026-07/dataset_flags.md`: volume / platform / sentiment MoM per model, the
carline's country split, story clusters for both months, the carline's daily peaks (what to click
into in Brandwatch), **every carline positive/negative row with its text** (the Phase 3 retag
list), a **noise scan** against `noise_patterns` in `config/carlines.json` (candidate boolean
exclusions — reviewed, never auto-dropped), and the carline's news split into carline-titled vs
passing mention. Read it before `menu`; anything you exclude or retag here has to happen in
Brandwatch **before** you export the four chart files, because it changes all of them.

## 3. Bullet menu  ← HARD GATE

```bash
python3 run.py menu --carline Kuga --month 2026-07
```

Open `work/Kuga/2026-07/menu.md`. Check clusters, counts, and sample URLs against the
[analytical](../docs/analytical-standards.md) and [editorial](../docs/editorial-standards.md)
standards (£100 rule; passing mentions → notes; "no clusters of issues were identified as
drivers" when negative driver volume < 30; carline named first; counts alongside %). Edit and
**approve** before going further. Nothing is built into the deck until you do.

## 4. Skeleton + manual fixes

```bash
python3 run.py skeleton    --carline Kuga --month 2026-07 --report-date 17.08.2026
python3 run.py fixskeleton --carline Kuga --month 2026-07     # prints the manual checklist
```

`fixskeleton` lists the recurring `build_skeleton.py` bugs to fix by hand (orphaned relationship,
stale Content-Types overrides, the dropped notesSlide, and — for EU5 Kuga/Puma — the hardcoded
`EU7` subtitle). These are surfaced, not auto-applied, on purpose. Work through them in the
unpacked skeleton.

## 5. Confirm chart mapping, then charts

```bash
python3 run.py roles  --carline Kuga --month 2026-07     # config vs. actual skeleton
python3 run.py charts --carline Kuga --month 2026-07
```

`roles` reads each chart's categories and tells you whether `config/carlines.json` matches
reality. For **Puma** (and any carline marked `verify_chart_roles`), `charts` refuses to run
until you've confirmed the mapping and passed `--i-verified-roles`:

```bash
python3 run.py charts --carline Puma --month 2026-07 --i-verified-roles
```

## 6. Bullets, backing xlsx, pack

```bash
python3 run.py bullets --carline Kuga --month 2026-07     # needs the approved menu.md
python3 run.py embed   --carline Kuga --month 2026-07     # patch chart mapping first (it reminds you)
python3 run.py pack    --carline Kuga --month 2026-07     # → work/.../Kuga_Social_Media_July_2026_Sprint_Room.pptx
```

## 7. QA & deliver (manual — Phase 6)

Render to PDF, eyeball the four charts (ignore the documented LibreOffice-only quirks in
[`../docs/tooling-patterns.md`](../docs/tooling-patterns.md)), check the subtitle month and
title-slide date, tidy the semiprocess sheet, upload to the Trello card, and email **Val**
(CC Mike, Seb, Mia).

---

### One-carline cheat sheet

```bash
cd monthly-run
python3 run.py newmonth   --carline C --month YYYY-MM
python3 run.py doctor     --carline C --month YYYY-MM
python3 run.py validate   --carline C --month YYYY-MM
python3 run.py flags      --carline C --month YYYY-MM      # read dataset_flags.md; retag/exclude in BW
python3 run.py menu       --carline C --month YYYY-MM      # approve menu.md
python3 run.py skeleton   --carline C --month YYYY-MM --report-date DD.MM.YYYY
python3 run.py fixskeleton --carline C --month YYYY-MM
python3 run.py roles      --carline C --month YYYY-MM
python3 run.py charts     --carline C --month YYYY-MM
python3 run.py bullets    --carline C --month YYYY-MM
python3 run.py embed      --carline C --month YYYY-MM
python3 run.py pack       --carline C --month YYYY-MM
```

Run the block four times (Capri, Explorer, Kuga, Puma) each cycle.
