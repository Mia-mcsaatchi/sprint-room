---
name: sprint-room-monthly-update
description: Build the new month's Sprint Room slide for a Ford carline (Capri, Explorer, Kuga, or Puma). Use whenever the user wants to produce a monthly Sprint Room update, replicate last month's deck with new data, or roll a Ford LENS social media report forward by one month. Triggers on phrases like "do the [Capri/Kuga/Puma] sprint room for [month]", "make this month's deck", "replicate the March slide for April", or any request involving a Brandwatch chart export + last month's pptx where the goal is a new dated slide. The skill rebuilds the four charts (volume line, country pie, platform pie, sentiment combo bar+diamonds), writes evidenced bullets with inline hyperlinks to sample articles, embeds editable backing xlsx files per chart, and removes the older comparison slide.
---

# Sprint Room Monthly Update

This skill replicates the workflow used to build the April 2026 Explorer Sprint Room slide. It applies identically to Capri, Kuga, and Puma — the carline is just a variable. The output is a `.pptx` with three slides: title → new month (just built) → previous month (carried forward untouched for comparison).

The skill assumes the analyst will provide: last month's deck, this month's raw Brandwatch export covering both months, this month's chart-data exports (volume×month, country, platform, sentiment), and confirmation of the reporting carline.

The workflow has six stages. Run them in order. Each stage has gotchas documented from the original Explorer build — don't skip them.

## Stage 0 — Confirm scope before touching anything

Before any tool calls, confirm with the user:

1. **Carline** (Capri / Kuga / Puma — Explorer was the original, the others are this skill's purpose)
2. **Reporting month** (and current date, which goes on the title slide)
3. **Files received**: last month's `.pptx`, one raw-data xlsx (any filename; what matters is the `Query Name` column containing all 6 competitors), and the four chart exports
4. **Sentiment data state**: warn the user that if they have retagged sentiment between the chart export and the raw data, the chart export is authoritative for the chart bars but the raw data drives the bullet narratives — confirm which version they want used for each

If any file is missing, stop and ask. Don't proceed on assumptions.

Once confirmed, list the files into `/home/claude/inputs/` and rename them to predictable names (`last_month.pptx`, `raw.xlsx`, `chart_volume.xlsx`, `chart_country.xlsx`, `chart_platform.xlsx`, `chart_sentiment.xlsx`).

## Stage 1 — Validate the data

Before generating anything, sanity-check the inputs. This catches data problems before they become embedded in the deck.

Use `scripts/validate_inputs.py` (read it before running):

```bash
python3 scripts/validate_inputs.py --carline <Capri|Kuga|Puma> --raw raw.xlsx \
    --vol chart_volume.xlsx --country chart_country.xlsx \
    --platform chart_platform.xlsx --sent chart_sentiment.xlsx \
    --last-month last_month.pptx
```

The script reports:
- Row counts per query per month from the raw data
- Whether the 4th column of the volume chart is actually the new month (Brandwatch labels it with `YYYY-MM-DD 23:00:00` which is a UTC vs GMT quirk — that 4th column IS the new month)
- Whether all 6 competitors appear in the chart_sentiment export
- Whether the platform export is unfiltered ("all"/"all" dimensions) — if so, it's NOT the carline-specific slice and the user needs to re-export
- Sentiment % per competitor for both the new month and the prior month (computed from raw data)

If anything fails, report to the user and stop.

## Stage 2 — Build the bullet menu (BEFORE editing the deck)

This is the most analyst-judgement-heavy stage. The skill cannot write good bullets on its own — but it can supply the evidence and let the analyst confirm.

Run `scripts/extract_drivers.py` to produce a markdown bullet menu draft:

```bash
python3 scripts/extract_drivers.py --carline <Capri|Kuga|Puma> --raw raw.xlsx \
    --new-month <N> --prior-month <N-1> --out menu.md
```

The script does Layer 3+4 from the architecture doc: news+blog story clustering by normalized title, sentiment-by-page-type breakdown, country breakdowns, MoM deltas for volume and sentiment, and a top-driver list per model with `mention_count`, `representative_url`, and `representative_title`.

**Key gotchas from the Explorer build:**

- **Don't dismiss top clusters as "boolean noise" without reading the Full Text column.** The Explorer/Kuga recall PR was originally written off as Kuga noise — actually it's a Stoll & Sauer law-firm press release about recall 25S79, which affects Explorer + Capri (high-pressure fuel injection fire risk). Genuine carline content. If you wouldn't bet £100 that a cluster is noise, treat it as a real driver and label it accordingly.
- **Quantify every driver in the bullet.** Each named driver should carry its mention count. If a candidate driver has <5 mentions, drop it from the slide and move it to notes.
- **Hyperlink the key claims inline.** Top-3 named stories per bullet get a real URL. Pick the cleanest domain (e.g. `finanznachrichten.de` over a regional `24hamburg.de`).
- **Use the standard phrase** "no clusters of issues were identified as drivers" when negative-driver volume <30 for any model. Val expects this exact phrasing.

Show the analyst the bullet menu and the underlying driver table. Ask them to confirm or edit before proceeding to Stage 3. **Do not** auto-write bullets and slap them in the deck — the analyst's domain judgment is mandatory here.

## Stage 3 — Set up the deck skeleton

Read `references/pptx_structure.md` for the OOXML layout you're about to manipulate. Then run:

```bash
python3 scripts/build_skeleton.py --last-month last_month.pptx --carline <X> \
    --new-month "April 2026" --report-date "10.05.2026" --out skeleton/
```

The script:
1. Unpacks the deck to `skeleton/unpacked/`
2. Duplicates slide 2 (last month's slide) → becomes the new month's slide
3. Reorders slides: title → new month → last month (Feb slide and earlier are removed)
4. Duplicates charts 1–4 (last month's charts) → charts 9–12 (new month's charts) so editing the new doesn't disturb the old
5. Repoints the new slide's rels at charts 9–12
6. Updates the title slide month + date
7. Updates the new slide's subtitle (and fixes year typos — original Explorer deck said "2025" instead of "2026")
8. Registers chart9–12 + styles + colors in `[Content_Types].xml`

**Gotcha**: the script must pack with strict validation **on** (`pack.py` without `--validate false`). If it fails, the duplicated chart's `<c:ext>` may have an empty self-closing element — read `references/pptx_structure.md` for the fix.

## Stage 4 — Replace chart data

For each of the four new charts, run the appropriate script. Each script reads the chart's existing structure and replaces values atomically — never replaces by global string substitution which could double-apply.

```bash
python3 scripts/update_volume_chart.py skeleton/unpacked/ppt/charts/chart9.xml --vol chart_volume.xlsx
python3 scripts/update_country_chart.py skeleton/unpacked/ppt/charts/chart10.xml --country chart_country.xlsx
python3 scripts/update_sentiment_chart.py skeleton/unpacked/ppt/charts/chart11.xml \
    --new-month-sent chart_sentiment.xlsx --prior-from-last-month last_month.pptx
python3 scripts/update_platform_chart.py skeleton/unpacked/ppt/charts/chart12.xml --platform chart_platform.xlsx
```

**Sentiment chart specifics:**
- The 100% stacked bars show **the new month** (April in the original build).
- The diamonds show **the prior month** (March in the original).
- "positive pre" X-values = prior-month positive %s.
- "Negative Pre" X-values = `1 - prior-month negative %` (because the bar is stacked from the right; the diamond sits at the boundary between neutral and negative blocks).
- **Authority for the diamond values**: the *prior month's published chart* (extracted from `last_month.pptx`'s `chart3.xml`) — NOT the current raw data. Last month's published values are what the client saw; using newly-retagged numbers would create a mismatch. The script handles this; don't override.
- **Strip any manual `<a:t>X%</a:t>` data-label overrides on the new chart.** The original Explorer deck had stale labels frozen from a previous month that displayed wrong values. Auto-labels are reliable; manual overrides aren't. The script strips them by default.

**Country and platform pie specifics:**
- Both charts have per-data-point manual layout positions tuned for last month's slice sizes. After updating values, the script removes `<c:manualLayout>` entries on data labels so they auto-position around the new slices. Without this, labels overlap when slice sizes change.

## Stage 5 — Write the bullets into the deck

Use `scripts/write_bullets.py` with the analyst-confirmed menu from Stage 2:

```bash
python3 scripts/write_bullets.py skeleton/unpacked/ppt/slides/slide4.xml \
    --menu menu.md
```

The script reads the menu (markdown format with sections `## bullet1`, `## bullet2`, `## bullet3`, and a `## hyperlinks` table) and writes the three bullet paragraphs with inline `<a:hlinkClick>` elements pointing at new `<Relationship>` entries it adds to `slide4.xml.rels`.

**Hyperlink convention** (from the original Explorer March deck):
- Linked phrase is bold (`b="1"`).
- No explicit color or underline rPr — PowerPoint picks up the hyperlink theme color and underlines automatically when `<a:hlinkClick>` is present.
- The relationship ID is unique per slide and added to `slide4.xml.rels` with `Type=".../hyperlink"`, `Target=<url>`, `TargetMode="External"`.

The script does **not** write the notes slide. Put the notes in a `## notes` section of the menu (country breakdowns all/news-only and Insight 3 sentiment-by-page-type come from the `extract_drivers.py` output, plus competitor detail and closing commentary) and paste them into the new slide's notes by hand:
- Country breakdowns (all and news-only), from the raw data
- Insight 3 sentiment-source-by-page-type, from the raw data using the new sentiment tags
- Closing commentary paragraph, written from the menu

## Stage 6 — Embed editable backing xlsx files

This stage is what turns external-URL chart links into clickable "Edit Data in Excel" charts.

```bash
python3 scripts/embed_backing_xlsx.py skeleton/unpacked/
```

The script generates **four separate** xlsx files (one per chart, not shared — sharing one xlsx between charts is what corrupted the file in the original Explorer build) and writes them to `ppt/embeddings/`. It:
1. Creates `Microsoft_Excel_Worksheet.xlsx` (volume), `..._1.xlsx` (country), `..._2.xlsx` (sentiment), `..._3.xlsx` (platform).
2. Each xlsx has the correct **sheet name** that the chart's `<c:f>` cell formulas reference — typically `'Explorer July'` for charts 9/10/11 and `'Explorer'` for chart 12 (these names are arbitrary, inherited from the original template; they don't need to match the carline).
3. Cell layout matches each chart's formula refs exactly:
   - Volume: A11=`months`, B11:E11 = month labels, A12:A18 = competitor names + `Average` (using `=AVERAGE(B12:B17)` so it auto-recomputes if the analyst edits competitor values), B12:E18 = values.
   - Country: A22:B28 = label/value pairs.
   - Sentiment: I12:I17 = sentiment category labels, J11:O11 = competitor names, J12:O14 = pos/neu/neg %s, J15:O15 = prior-month positive %, J16:O16 = `1 − prior-month negative %`, J17:O17 = Y coordinates [3, 8.75, 14.6, 20.5, 26.3, 32.1].
   - Platform: A12:B17 = label/value pairs (Online News, X, Forums, Blogs, Review, Tumblr).
4. Repoints each chart's `rId4` from the old Egnyte external URL to `../embeddings/Microsoft_Excel_WorksheetN.xlsx` with relationship Type=`.../relationships/package`.
5. Registers each xlsx as an `<Override>` in `[Content_Types].xml` with ContentType `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.

**Critical pitfalls — these caused the "PowerPoint found a problem" error in the Explorer build:**
- **Each chart needs its own xlsx file.** Don't share one between charts.
- **No duplicate `<Override>` entries in `[Content_Types].xml`.** Use a regex that handles both leading-whitespace and no-leading-whitespace variants when removing prior entries.
- **Don't change the `<c:externalData r:id="rId4">` element** inside the chart XML. Only change the relationship's `Type` and `Target`. The chart's pointer is unchanged.

## Stage 7 — Pack, render, QA

```bash
python3 /mnt/skills/public/pptx/scripts/clean.py skeleton/unpacked/
python3 /mnt/skills/public/pptx/scripts/office/pack.py skeleton/unpacked/ output.pptx
```

**Pack with validation ON** (no `--validate false`). If it fails, fix the underlying issue rather than skipping validation. Common causes from the Explorer build: empty `<c:ext .../>` children in duplicated chart XMLs (open the file, add `<c:dummy/>` inside or remove the element), or unresolved relationships.

Render to PDF for visual QA:

```bash
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

**Known LibreOffice rendering quirks that are NOT bugs in your file:**
- Sentiment chart shows only 1–2 diamonds instead of 6. PowerPoint renders all 6 correctly.
- Line chart's right-edge labels stack on top of each other for competitors near the same Y value. This is in the original deck too.
- Bullet 3 may clip at the bottom edge. PowerPoint's autofit handles this.

Don't chase these in LibreOffice — they go away in PowerPoint. Do chase any data anomalies (wrong percentages, missing categories, weird labels with old values).

Copy to outputs and present:

```bash
cp output.pptx /mnt/user-data/outputs/<Carline>_Social_Media_<Month>_Sprint_Room.pptx
```

## What the skill does NOT do (out of scope)

- Phase 1 of the process doc (boolean cleaning in Brandwatch) — analyst job
- Phase 4 (semiprocess Google Sheets pivot tables) — analyst job
- Phase 6 sub-steps: uploading to Trello card, sending email — analyst job
- Sentiment retagging — analyst job in Brandwatch
- Cross-carline rollups or multi-carline single decks — this skill is per-carline

## When the input deck isn't from Explorer

The four other Ford carlines (Capri, Kuga, Puma) follow the same template — same chart layouts, same slide structure, same notes structure. The only variables are:
- The subtitle text (e.g. "Capri – Social Media (EU7 April 2026)" instead of "Explorer – Social Media...")
- The carline name in bullet text
- The list of competitors (cross-check with the user — Capri's competitor set may differ from Explorer's)
- The country/platform pies still pull the same way

When the user gives you a non-Explorer last-month deck, do NOT assume Explorer competitor list; read the actual `<c:cat>` strings from `chart3.xml` to get the real competitor names.

## References

- `references/pptx_structure.md` — OOXML chart/slide layout details
- `references/bullet_writing.md` — How to write good Sprint Room bullets (style, quantification, hyperlink rules, Val's standard phrases)
- `references/embed_xlsx_pitfalls.md` — Detailed notes on the file-corruption bugs and how to avoid them
