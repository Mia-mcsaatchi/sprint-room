# Slide & Chart Configs

Deck structure, per-carline chart ordering, and the backing-workbook layout. These differ by
carline and are the single biggest source of "why did the deck corrupt" errors — verify against
the actual deck each cycle.

## Deck structure

Three slides after a build: **title → new month → previous month**. The skeleton builder clones
last month's slide into the new-month slide and drops any slide older than "previous month".

Four charts per reporting slide:
1. Volume — rolling line (4 months × competitors + Average)
2. Country — pie (current month)
3. Platform — pie (current month)
4. Sentiment — combo (100% stacked bars = current month; diamonds = prior month)

The skeleton builder duplicates last month's charts 1–4 into charts **9–12** so editing the new
month doesn't disturb the carried-forward slide.

## Per-carline chart order (⚠️ differs from the skill's Explorer default)

The skill's Explorer template maps `chart9=volume, chart10=country, chart11=sentiment,
chart12=platform`. Other carlines do **not** follow this. Known mappings:

| Carline | Chart 10 | Chart 11 | Chart 12 | Notes |
|---------|----------|----------|----------|-------|
| **Explorer** (skill default) | Country | Sentiment | Platform | volume = chart9; 5 competitors |
| **Kuga** | **Country** | **Platform** | **Sentiment** | EU5 |
| **Capri** | **Platform** | **Country** | **Sentiment** | EU7; 6 competitors |
| **Puma** | order swapped from Explorer template (`chart3=platform`, `chart4=sentiment`) | | | EU5; **5 competitors, not 6** |

> **The `embed_backing_xlsx.py` chart→xlsx mapping must be patched per carline before running.**
> Don't assume the Explorer order. Confirm each chart's role by reading its `<c:f>` cell
> references / category strings in the actual deck.

## Backing workbook (editable "Edit Data in Excel") embedding

Turns external-URL chart links into clickable editable charts. Each chart gets its **own** xlsx
under `ppt/embeddings/` — **never share one xlsx between charts** (sharing corrupted the file in
the original Explorer build).

Cell layouts each xlsx must match (the chart's `<c:f>` formulas reference these exactly):

- **Volume:** `A11=months`, `B11:E11` = month labels, `A12:A18` = competitor names + `Average`
  (`=AVERAGE(B12:B17)`), `B12:E18` = values.
- **Country:** `A22:B28` = label / value pairs.
- **Sentiment:** `I12:I17` = category labels; `J11:O11` = competitor names; `J12:O14` =
  pos/neu/neg %; `J15:O15` = prior-month positive %; `J16:O16` = `1 − prior-month negative %`;
  `J17:O17` = Y coords `[3, 8.75, 14.6, 20.5, 26.3, 32.1]`.
- **Platform:** `A12:B17` = label / value pairs (Online News, X, Forums, Blogs, Review, Tumblr).

Sheet names inside the xlsx are inherited from the template (e.g. `'Explorer July'` for charts
9/10/11, `'Explorer'` for chart 12) — arbitrary but **must match** the chart's `<c:f>` refs; they
don't need to match the carline.

### Kuga's actual layout (differs from the above — confirmed Sep 2026 cycle)

The Kuga deck already embeds one xlsx per chart and registers xlsx via `<Default Extension="xlsx">`,
so the safest path is to **copy each chart's existing workbook and edit its cells**, not regenerate
it with `embed_backing_xlsx.py` (whose cell map is Explorer's):

| Chart | Sheet | Cells |
|-------|-------|-------|
| Volume | `'Kuga - 2'` | `A11:E11` months, `A12:A17` series (Tiguan, 3008, Tucson, Jaecoo 7, Kuga, Bigster), `A18` Average (`=AVERAGE(B12:B17)`, **includes Kuga**) |
| Country | `Sheet1` | `A12:B16` (UK, ES, DE, IT, FR) |
| Platform | `Sheet1` | `A12:B17` (Forums, Online News, X, Review, Blogs, Tumblr) |
| Sentiment | `Kuga` | `K11:P11` models (Tucson, Jaecoo 7, Bigster, 3008, Tiguan, Kuga); `J12:J14` pos/neu/neg rows; `K17:P17` Y = `[2.6, 7.6, 12.6, 17.6, 22.6, 27.6]`; `K19:P19` `Positive Pre`; `K21:P21` `Negative Pre` (`1 − prior neg`) |

Embedding names: charts 1–4 use `Microsoft_Excel_Worksheet{,1,2,3}.xlsx`, charts 5–8 use `…4–7`.
When the older slide is dropped, its charts 5–8 and workbooks 4–7 go with it; the new charts 9–12
take workbooks `…8–11` and their own `themeOverride9–12.xml` copies.

### Content-Types registration — the corruption trap

- If the deck already registers xlsx via a **`<Default Extension="xlsx">`** entry in
  `[Content_Types].xml`, do **not** add per-file `<Override>` entries — doing so risks a
  duplicate-Override corruption.
- If it does not, register each xlsx as an `<Override>` with ContentType
  `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, and use a regex that
  removes prior entries in both leading-whitespace and no-whitespace variants so you never
  create duplicates.
- **Don't touch** the `<c:externalData r:id="rId4">` element inside the chart XML. Only change
  the *relationship's* `Type` (to `.../relationships/package`) and `Target` (to
  `../embeddings/Microsoft_Excel_WorksheetN.xlsx`).

## Chart data patching technique

Patch `numCache` / `strCache` positionally with lxml: find the `<c:numCache>` / `<c:strCache>`
block inside each `<c:ser>`, remove existing `<c:pt>` entries, inject new `<c:pt idx="N">`
elements by index. **Never** use string-replace on chart values — it silently corrupts when old
and new values overlap.

After updating pie values, strip stale hardcoded percentage data labels and remove
`<c:manualLayout>` label positions so labels auto-place around the new slices.
