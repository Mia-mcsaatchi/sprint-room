# Tooling Patterns

Reliable, hard-won technical patterns. Skipping these re-introduces bugs that have already been
paid for once.

## Reading Brandwatch exports

BW xlsx exports **do not** parse with openpyxl and **do not** convert reliably via LibreOffice
CSV. Use direct `zipfile` + `lxml iterparse` on `xl/worksheets/sheet1.xml`:

- Text lives in **inline strings** (`<is><t>` elements), not sharedStrings (which are empty).
  openpyxl returns blanks / "0 is out of range".
- Column letters: `re.match(r'[A-Z]+', cell_ref).group()`.
- Numeric cells: read `<c:v>`.
- **Header detection:** scan for the row where **column B == "Query Name"**. Rows above it are a
  metadata block (Report / Brand / From / To / Label ...). Data starts on the next row.
  Do **not** hardcode a header row number, and do **not** trust the filename.
- **Big files (80MB+):** stream with `iterparse` and call `row.clear()` after each row to avoid
  memory exhaustion.

A ready-made implementation is in [`../tools/bw_read.py`](../tools/bw_read.py):

```bash
python3 tools/bw_read.py FILE.xlsx --querydist   # validate contents by Query Name
python3 tools/bw_read.py FILE.xlsx --head 20
```

## OOXML deck manipulation

Unpack via `unzip`, manipulate the XML directly (lxml for charts, regex for slide text and
rels), repack via `zip -Xr` from **inside** the unpacked directory, then validate.

- **Chart data:** patch `<c:numCache>` / `<c:strCache>` positionally with lxml — find the block
  inside each `<c:ser>`, remove existing `<c:pt>`, inject new `<c:pt idx="N">`. Never string-
  replace values (silently corrupts on overlap).
- **Pie charts:** after updating values, strip stale hardcoded `%` data labels and remove
  `<c:manualLayout>` positions so labels auto-place.
- **Pack with validation ON.** If packing fails, fix the cause — don't disable validation. A
  common cause is an empty self-closing `<c:ext .../>` in a duplicated chart (add a child or
  remove the element).

## Known `build_skeleton.py` bugs — manual fix EVERY cycle

These recur; assume they are present until proven otherwise:

1. **Orphaned relationship not removed.** The `<Relationship>` for the deleted older slide stays
   in `ppt/_rels/presentation.xml.rels` — remove it manually.
2. **Stale Content-Types overrides.** `[Content_Types].xml` overrides for the deleted chart files
   and deleted slide XML aren't cleaned — remove them manually.
3. **Dropped notesSlide relationship.** The duplicated new slide loses its `notesSlide`
   relationship. Manually: create the notesSlide XML, its `.rels` file, register it in
   `[Content_Types].xml`, and re-insert the `rId2` relationship into the new slide's rels.
4. **Hardcoded `EU7` in the subtitle regex.** Fails for EU5 carlines (**Kuga, Puma**) — apply the
   subtitle update manually for those.
5. **Duplicated charts still share the old embedding and theme override.** It copies
   `chartN.xml`, `colorsN`, `styleN` and the rels, but the new chart's rels keep pointing at the
   *old* `Microsoft_Excel_Worksheet*.xlsx` and `themeOverrideN.xml`. Editing the new chart's data
   then rewrites last month's workbook (the "never share one xlsx" trap). Copy the embedding and
   theme override per new chart, repoint the rels, and register the theme overrides in
   `[Content_Types].xml`.
6. **Hardcoded `/mnt/skills/...` paths** for `unpack.py` / `add_slide.py`. In Claude Code cloud
   sessions the pptx skill lives elsewhere; unpack with `zipfile` and duplicate parts directly.

The Sept 2026 Explorer build did all of the above by hand in one script (duplicate slide + charts
+ embeddings + theme overrides + notes, retarget hyperlinks, refill caches and workbooks), then ran
the pptx skill's `clean.py` and `validate.py --original last_month.pptx`. Ask Claude to follow the
same approach rather than `build_skeleton.py` as-is.

## Template facts (Explorer deck, as of June 2026)

- Reporting slide charts: **chart1 = volume (line)**, **chart2 = country (pie)**,
  **chart3 = platform (pie)**, **chart4 = sentiment (bar + scatter diamonds)**.
- Workbook sheets: `'Explorer July'` (volume, country, sentiment), `'Explorer'` (platform).
  Volume `A1:E8` (row 1 months, rows 2–7 models, row 8 `=AVERAGE`); sentiment `A1:G7`
  (row 1 models; rows 2–4 pos/neu/neg; row 5 prior positive; row 6 `1 − prior negative`;
  row 7 Y = 3, 8.75, 14.6, 20.5, 26.3, 32.1); platform `A1:B6`.
- **Country pie quirk:** the chart formula points at `A1:B7` but the template workbook holds the
  values at `A22:B28`. Write the new values to both so "Edit Data" and the chart agree.
- Fixed orders — volume: Enyaq, XC40 Recharge/ EX40, E-3008, Megane E-Tech, ID.4, Explorer;
  sentiment: Megane E-Tech, Peugeot E3008, Skoda Enyaq, Volvo XC40 Recharge/ EX40, VW id4,
  Explorer; country codes UK, ES, DE, FR, IT, NL, NO; platforms Online News, X, Forums, Blogs,
  Review, Tumblr.
- The Summary text box (`TextBox 16`) is `spAutoFit` and already reaches the slide bottom — keep
  bullets about as long as last month's; overflow goes to the notes.
- Notes follow `Insight 2` (country all/news, competitor platform breakdown) and `Insight 3`
  (positive/negative by content source) headings.

## Driver extraction

`scripts/extract_drivers.py` reads the raw export via LibreOffice CSV conversion, which is
unreliable for Brandwatch files (see above). Load the export with `tools/bw_read.py` instead and
apply the same clustering (normalised title, news + blog first).

## Backing xlsx pitfalls (caused "PowerPoint found a problem" in the Explorer build)

- Each chart needs its **own** xlsx — don't share one between charts.
- No duplicate `<Override>` entries in `[Content_Types].xml` (handle both whitespace variants
  when de-duping). If the deck uses `<Default Extension="xlsx">`, skip Overrides entirely.
- Don't modify `<c:externalData r:id="rId4">` inside the chart XML — only the relationship's
  `Type` and `Target`.

See the skill's own reference notes for the fullest detail:
- [`../skills/sprint-room-monthly-update/references/pptx_structure.md`](../skills/sprint-room-monthly-update/references/pptx_structure.md)
- [`../skills/sprint-room-monthly-update/references/embed_xlsx_pitfalls.md`](../skills/sprint-room-monthly-update/references/embed_xlsx_pitfalls.md)
- [`../skills/sprint-room-monthly-update/references/bullet_writing.md`](../skills/sprint-room-monthly-update/references/bullet_writing.md)

## LibreOffice render quirks that are NOT bugs in your deck

When rendering to PDF for QA, ignore these (PowerPoint renders correctly):
- Sentiment chart shows only 1–2 diamonds instead of 6.
- Line-chart right-edge labels stack for competitors near the same Y value (present in the
  original deck too).
- Bullet 3 clips at the bottom edge (PowerPoint autofit handles it).

Chase real data anomalies (wrong %, missing categories, labels showing old values) — not these.

## Rendering in a fresh cloud session

Claude Code cloud containers may ship only `libreoffice-core`, which opens nothing ("source file
could not be loaded", even for a .txt). Install Impress once per session:

```bash
apt-get install -y --no-install-recommends libreoffice-impress
```

Then render with the pptx skill's `scripts/office/soffice.py --headless --convert-to pdf` and
`pdftoppm`. Hyperlinks can render in different colours in LibreOffice even when their XML is
identical — PowerPoint applies the theme hyperlink colour to all of them.
