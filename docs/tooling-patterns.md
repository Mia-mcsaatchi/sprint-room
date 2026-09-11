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

For pandas work use `bw_read.load_dataframe(path)` — it returns a DataFrame keyed by header
label with `Date` parsed and `Month` bucketed on the local-time month boundary (a 23:xx UTC
mention on the last day of the month belongs to the next month, which is what Brandwatch's own
`23:00:00` column label reflects). `extract_drivers.py`, `validate_inputs.py` and
`tools/dataset_flags.py` all load through it.

- **LibreOffice CSV conversion can fail outright**, not just unreliably: on the Capri Aug-2026
  export `soffice --convert-to csv` returned "source file could not be loaded" and produced no
  file. The skill scripts keep it only as a fallback when `bw_read` cannot be imported.
- **Column letters move between exports.** The Capri Aug-2026 file had 177 columns (not 186);
  `Full Text` was `AF`, `Reach (new)` `BX`, `Page Type Name` `DW`. Never address a column by
  letter — look it up by header name.

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
