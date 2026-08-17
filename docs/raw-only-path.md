# The raw-only path (raw + last month's deck only)

The original monthly run needs six inputs (last month's deck + a raw export + **four**
Brandwatch chart exports). The four chart exports are **derivable** from the raw feed plus
last month's published deck, so this path shrinks the monthly inputs to just:

```
monthly-run/inputs/<CARLINE>/<YYYY-MM>/
  raw.xlsx          # the Bulk Mentions export (may be split across pulls — see "Merging")
  last_month.pptx   # the previous month's published deck
```

Everything quantitative is still computed deterministically **before** any language model
(the golden rule holds). The chart exports are reconstructed, not hand-made.

## Why this exists / when to use it

Use it whenever you have the raw feed and last month's deck but not the four chart exports —
or to skip the manual chart-export step entirely. The carline's **prior months and prior
sentiment diamonds are still carried from the published deck**, never re-derived (analytical
standard: match what the client saw).

## The pipeline

```bash
cd monthly-run
python3 run.py newmonth   --carline Puma --month 2026-07     # scaffold; drop raw.xlsx + last_month.pptx in
python3 run.py clean      --carline Puma --month 2026-07 [--retag-passing]   # → work/.../cleaning_menu.md  (SIGN-OFF)
python3 run.py derive     --carline Puma --month 2026-07 [--retag-passing]   # → the 4 chart_*.xlsx in inputs/
python3 run.py rawmenu    --carline Puma --month 2026-07     # → work/.../menu.md   ← HARD GATE (approve bullets)
#   save approved bullets as work/.../menu_confirmed.md (write_bullets format: ## bullet1 …)
python3 run.py skeletonraw --carline Puma --month 2026-07 --report-date DD.MM.YYYY
python3 run.py chartsraw  --carline Puma --month 2026-07     # updates the new slide's 4 charts (role-detected)
python3 run.py embedraw   --carline Puma --month 2026-07     # write new-month values into each chart's own backing xlsx (Edit-Data linked)
python3 run.py bullets    --carline Puma --month 2026-07     # uses menu_confirmed.md, auto-finds the new slide
#   author work/.../notes.md (overflow: country/platform breakdowns, competitor caveats)
python3 run.py notes      --carline Puma --month 2026-07     # attach a speaker-notes slide to the new slide
python3 run.py pack       --carline Puma --month 2026-07     # zip + validate + best-effort PDF/PNG QA
```

Add `--also other_raw.xlsx` to any of `clean`/`derive`/`rawmenu` to **merge** a second export
(see below). Add `--dry-run` to preview any stage's command.

### What each new stage does

| Stage | Script | Output |
|-------|--------|--------|
| `clean` | `clean_raw.py` | Drops false-brand/off-topic rows, retags passing/other-model → neutral, flags the rest REVIEW. Writes a line-by-line `cleaning_menu.md` for sign-off. **Nothing is dropped silently.** |
| `derive` | `derive_charts.py` | Reconstructs `chart_volume/country/platform/sentiment.xlsx` in Brandwatch export format from the **cleaned** raw + last month's deck. |
| `rawmenu` | `build_menu.py` | Volume/sentiment MoM, story clusters, country/platform breakdowns from cleaned raw → the bullet-menu **hard gate**. |
| `skeletonraw` | `build_skeleton_raw.py` | Layout-agnostic skeleton: clones the reported-month slide **and its four charts** (own colors/style/themeOverride/embedding copies), drops the oldest month, fixes all package bookkeeping. |
| `chartsraw` | `update_*_chart.py` | Updates the new slide's four charts, located **by role** (not by fixed chart number). |
| `embedraw` | `link_embeddings.py` | Writes the new-month values into each chart's **own** backing workbook by reading its `<c:f>` refs + cache, so "Edit Data in Excel" matches the chart; the volume Average is written as a live `=AVERAGE()` formula. |
| `notes` | `write_notes.py` | Clones a template notes-slide, swaps in `work/.../notes.md` (overflow per editorial-standards.md), and wires it to the new slide. |

## Cleaning policy (the £100 rule, per carline)

- **DROP** — false-brand matches (Capri = the island/design/person; Puma = sportswear/animal/
  helicopter) and clearly off-topic non-automotive items. False terms are **carline-specific**
  and applied **only to the carline** (competitors are clean; Phase-1 boolean cleaning is theirs).
- **RETAG → neutral** — genuine automotive but the carline is incidental (another model is the
  subject / passing multi-brand), when currently tagged positive/negative.
- **REVIEW** — anything uncertain; surfaced for the analyst, never auto-dropped.
- `--retag-passing` promotes carline **news/blog** passing pos/neg to neutral instead of REVIEW
  (use when the analyst has pre-approved treating them as passing).

Keyword lists live in `clean_raw.py` (`FALSE_BY_CARLINE`, `MODEL_KW`, `STRONG_KW`,
`FORCE_KEEP/RETAG_BY_CARLINE`). Add new carlines there.

## Gotchas discovered building this (do not re-learn them)

1. **A single export can silently drop a query.** The July Capri pull came back missing Polestar 2;
   a re-pull was missing VW ID.5. Neither was complete. `load_merged` reconstructs the full model
   set by taking the primary export plus **only the queries missing** from it (`--also`). Always
   confirm every configured model is present before building.
2. **Chart roles differ per carline AND the deck's chart numbering is not stable.** Never hardcode
   `chart3=sentiment`. Explorer sentiment=chart3, Capri=chart4, Puma=chart16. All role detection
   (derive, chartsraw, the sentiment prior-diamond reader) classifies **by chart type + series
   names + categories**, and finds the reported-month slide by its subtitle month.
3. **Diamond (prior-month) series names vary.** Explorer/Capri use `positive pre`/`Negative Pre`;
   Puma labels them by month (`Positive May`). `update_sentiment_chart.py` matches by polarity and
   renames the month to the new prior month (`--prior-month-name`).
4. **Never string-replace chart values.** `update_volume_chart.py` now patches numCache
   **positionally** — replacing by old value silently corrupts when a new value equals another
   month's old value (e.g. April 1005 reused as May), which rotated the whole line.
5. **The volume Average is the mean of ALL models (incl. the carline)** in these decks
   (June Average 141 = mean of all 5/7), not competitors-only. Carried-month averages must match
   the deck exactly, so `derive` computes it that way and `update_volume` prefers the export's row.
6. **The EU5/EU7 subtitle is carline-specific** — `build_skeleton_raw` reads the segment from
   `config/carlines.json` and does not hardcode `EU7`.
7. **Cloned charts must own their parts.** `skeletonraw` duplicates each chart's colors, style,
   themeOverride **and embedding** and repoints the rels — sharing an embedding between two charts
   corrupts the file.
8. **Tooling parity:** the current `pptx` skill has **no `pack.py`** (a .pptx is a plain zip;
   validation is `office/validate.py`) and no `unpack.py`. `pack_and_qa.py` packs by zip and
   validates with `--original` so the template's own schema quirks (e.g. `authors.xml`) aren't
   read as new errors.
9. **LibreOffice PDF rendering may be unavailable** in a given sandbox (import filter errors on
   every deck). The pack step's PDF/PNG render is best-effort and never fails the build; QA in
   PowerPoint. The deck is still schema-valid.
10. **"Only report if there's something worth it" (Val).** A quiet month (e.g. Capri July — volume
    down, no news drivers, sub-30 fragmented negativity) may warrant **no slide at all**. Run
    through the menu, then decide.

## What this path does NOT do

- Phases 1/3/4/6 (boolean cleaning, Brandwatch retag, Sheets pivots, Trello/email) remain manual.
- The old Explorer-specific `embed_backing_xlsx.py` is superseded by `embedraw`/`link_embeddings.py`
  for this path (layout-agnostic, reads each chart's own `<c:f>` refs).
