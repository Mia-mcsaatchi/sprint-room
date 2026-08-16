# Migration & Onboarding

How to pick this programme up on a new machine, or hand it to a new analyst.

## 1. Accounts & access you need

- **Brandwatch** — the "Ford Sprint Room" dashboards (ignore Focus, Mustang Mach-E, Puma Gen-E).
- **Trello** — the current month's card (decks live here).
- **Google Sheets** — the per-carline *semiprocess* workbooks (raw data + helper formulas + pivots).
- **The previous month's `.pptx`** for the carline you're building.

## 2. Local environment

Python 3 with:

```bash
pip install lxml openpyxl
```

- `lxml` — all OOXML chart/slide manipulation and Brandwatch parsing.
- `openpyxl` — writing backing xlsx / fixtures only. **Do not** use it to *read* Brandwatch
  exports (see [`docs/tooling-patterns.md`](docs/tooling-patterns.md) — they use inline strings).
- LibreOffice (`soffice`) — only for PDF rendering during QA. Its render quirks are documented
  and are **not** bugs in your deck.

Quick sanity check that parsing works:

```bash
python3 tools/bw_read.py samples/raw_export_SAMPLE_schema_only.xlsx --querydist
python3 tools/bw_read.py samples/chart-exports/chart_volume_monthsxqueries.xlsx --head 12
```

## 3. Where the skill lives

The skill is designed to be dropped into a Claude skills directory
(`/mnt/skills/user/sprint-room-monthly-update/`). This repo keeps a verbatim copy under
[`skills/`](skills/) as the versioned source of truth. When you update the skill in the
working environment, commit the change back here so the two never drift.

## 4. First build on a new setup — recommended order

1. Read [`docs/workflow-phases.md`](docs/workflow-phases.md) end to end.
2. Do Phases 1 & 3 in Brandwatch by hand (query cleaning + sentiment retag). These are
   **analyst judgement** and are deliberately not automated.
3. Export the four chart-data files + one raw-data export covering both months.
4. Run the skill from Stage 0. It confirms scope, validates inputs, and stops if a file is
   missing or the platform export came out unfiltered (`all`/`all`).
5. Approve the **Stage 2 bullet menu** — this is the hard gate. Nothing gets built into the
   deck until you've signed off the clusters, counts, and URLs.
6. Let it run Stages 3–7 (skeleton → charts → bullets → backing xlsx → pack → PDF QA).

## 5. Per-cycle manual fixes you will still need

`build_skeleton.py` has known bugs that recur every cycle. Do not assume they are fixed —
[`docs/tooling-patterns.md`](docs/tooling-patterns.md#known-build_skeletonpy-bugs) lists all
four with the manual remedy. The big ones: orphaned relationship/Content-Types entries for
the dropped slide, the missing `notesSlide` relationship on the new slide, and the hardcoded
`EU7` subtitle regex that breaks for Kuga/Puma (EU5).

## 6. Keeping this repo alive

- Commit the skill back here whenever it changes in the working environment.
- When a per-cycle "gotcha" bites you, add it to the relevant doc in the same PR — the value
  of this repo is that hard-won lessons don't get re-learned.
- Don't commit full raw Brandwatch exports (they're large and contain scraped third-party
  content). `.gitignore` already blocks them; use small fixtures like the one in `samples/`.
