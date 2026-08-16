# PPTX Structure Reference

Sprint Room decks follow a consistent OOXML layout. This doc captures what each part holds and what to watch out for.

## Slide layout

```
ppt/
├── slides/
│   ├── slide1.xml       — Title slide ("Explorer Team Sprint Room / Social Media <Month>")
│   ├── slide2.xml       — Last month's reporting slide (becomes "comparison" after roll-forward)
│   ├── slide3.xml       — Two months ago slide (deleted in roll-forward)
│   └── slide4.xml       — NEW: this month's reporting slide (created during roll-forward)
├── slides/_rels/
│   └── slide<N>.xml.rels — chart refs (rId5–8 for chart1–4, hyperlink rels rId3, rId4, etc.)
├── charts/
│   ├── chart1.xml–chart4.xml    — last month's 4 charts
│   ├── chart5.xml–chart8.xml    — two-months-ago charts (deleted)
│   └── chart9.xml–chart12.xml   — NEW: this month's charts (duplicated from chart1–4)
├── charts/_rels/
│   └── chart<N>.xml.rels        — each chart has rId1=style, rId2=colors, rId3=themeOverride, rId4=oleObject/package
├── embeddings/                  — embedded xlsx files for charts
├── notesSlides/
│   ├── notesSlide1.xml          — notes for slide2 (last month)
│   ├── notesSlide2.xml          — notes for slide3 (two months ago, deleted)
│   └── notesSlide3.xml          — NEW: notes for slide4 (this month)
└── presentation.xml             — slide order in <p:sldIdLst>
```

## The four charts on each reporting slide

| Chart # (last month) | Chart # (this month) | Type | Source data shape |
|---|---|---|---|
| chart1 | chart9 | Volume line chart (4 months × 6 competitors + Average) | sheet 'Explorer July' A11:E18 |
| chart2 | chart10 | Country pie (7 categories, single month) | sheet 'Explorer July' A22:B28 |
| chart3 | chart11 | Sentiment combo (stacked bar + diamond scatter) | sheet 'Explorer July' I11:O17 |
| chart4 | chart12 | Platform pie (6 platforms, single month) | sheet 'Explorer' A12:B17 |

The sheet names `Explorer July` and `Explorer` are arbitrary — they come from the original template. The chart's `<c:f>` cell formulas reference these names. When embedding new backing xlsx, the **sheet names must match what the formulas expect**, not the carline being reported on.

## Sentiment chart anatomy (the trickiest one)

Series order inside chart11.xml:
1. `positive` — stacked bar, 6 values, the new month's pos %
2. `neutral` — stacked bar, 6 values, the new month's neu %
3. `negative` — stacked bar, 6 values, the new month's neg %
4. `positive pre` — XY scatter, 6 (x,y) pairs, the prior month's pos % as X
5. `Negative Pre` — XY scatter, 6 (x,y) pairs, **`1 − prior-month neg %`** as X

The diamonds use 6 Y values `[3, 8.75, 14.6, 20.5, 26.3, 32.1]` — these are pre-tuned row positions on the chart's Y axis, not data values. Don't change them.

Categories are inverted on the visible axis (`<c:crosses val="max"/>` on the value axis flips the bars). The first category in data order (Megane E-Tech) renders at the **bottom** of the chart; the last (Explorer) renders at the **top**.

Manual `<a:t>X%</a:t>` data-label overrides may exist on specific data points — typically idx=0 (top of chart = Explorer in original template) and idx=5 (bottom = Megane). These can become stale (show old months' values). Always strip them when updating the chart so labels auto-compute from data.

## Volume line chart anatomy

7 series total: 6 competitors + Average. Each series has:
- One `<c:cat>` block with 4 month labels (e.g. January / February / March / April)
- One `<c:val>` block with 4 numeric values
- Per-data-point label overrides on idx=0,1,2 (hidden via `<c:delete val="1"/>`) and idx=3 (visible at end of line with manualLayout positioning)

The idx=3 manualLayout positions are tuned for the last month's line endpoints. When values change, labels may stack. **Don't try to retune these positions** — same issue exists in original deck and PowerPoint handles it adequately. Stripping the manualLayout for line charts is risky (labels disappear into other elements).

## Pie chart anatomy

Both country and platform pies have 1 series with 6–7 data points and per-data-point label positions via `<c:manualLayout>`. When slice sizes change between months (e.g. DE went from 23% in March to 31% in April), labels positioned for the old slice land in the middle of the new wedge and overlap.

**Fix:** strip `<c:layout><c:manualLayout>...</c:manualLayout></c:layout>` from each `<c:dLbl>` — labels auto-position around the new slices.

Leave any `<c:dLbl>` that has `<c:delete val="1"/>` alone — those are intentionally hidden labels (e.g. for empty `Review` or `Tumblr` platforms).

## Hyperlink relationships

Each hyperlink in slide4.xml needs:
1. A `<Relationship>` entry in `slide4.xml.rels` with:
   - Id = unique (e.g. `rId3`, `rId4`, `rId9`, `rId10`)
   - Type = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`
   - Target = the full URL
   - TargetMode = `External`
2. An `<a:hlinkClick xmlns:r="..." r:id="<rId>"/>` element inside the linked run's `<a:rPr>`

PowerPoint applies the hyperlink theme color and underline automatically. Don't manually set color or underline on the run; the theme handles it.

## Pack validation

The deck must pack cleanly with strict OOXML validation. Common reasons it doesn't:

| Symptom | Cause | Fix |
|---|---|---|
| Empty `<c:ext .../>` rejected | Duplicated chart XML has self-closing extension marker | Open chart XML, locate offending `<c:ext>` self-close, replace with `<c:ext ...><c:dummy/></c:ext>` or remove element entirely |
| Duplicate Override in Content_Types | Previous regex left old entries behind | Re-run cleanup with `<Override PartName="..."[^>]*/>` regex that handles all whitespace variants |
| Reference to non-existent part | Chart rel points to a chart XML that was deleted | Either restore chart XML or fix the rel target |
| Multiple charts share one embedding | Non-standard structure | Build one xlsx per chart |
