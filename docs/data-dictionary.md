# Data Dictionary

Structure of the Brandwatch exports the pipeline consumes. Read
[`tooling-patterns.md`](tooling-patterns.md) for *how* to parse them (they need inline-string
handling, not openpyxl).

## Raw "Bulk Mentions Download" export

One row per mention. **~180 columns** — 186 (A → GD) in the schema-only sample, 177 in the Capri
Aug-2026 export. The column *set* is stable but the *letters* are not (Full Text has been seen at
both `AI` and `AF`), so always look columns up by header name; the letters below are the sample's.
Structure:

- **Rows 1–5:** metadata block — `Report:`, `Brand:` (the tracked models across columns B..),
  `From:`, `To:`, `Label:`.
- **Header row:** the row where **column B == "Query Name"** (varies by file — detect it, don't
  hardcode). In the sample this is row 6, data from row 7.
- **Data rows:** below the header.

Key columns used downstream:

| Col | Field | Used for |
|-----|-------|----------|
| A | Query Id | model identity |
| B | Query Name | model label — **validate the export by this column's distribution** |
| C | Date | month bucketing |
| D | Title | story clustering (normalised title) |
| E | Url | volume = `COUNTUNIQUE(URL)`; hyperlink sources |
| F | Domain | picking the cleanest source domain |
| G | Sentiment | sentiment %; retag target |
| H | Page Type | news/blog/forum filtering for clustering |
| I | Language | translation need |
| J / M | Country Code / Country | country pie + country breakdowns |
| AI | Full Text | **read before dismissing a cluster as noise** (£100 rule) |
| CA | Reach (new) | source-evidence ranking |
| EI | Page Type Name | platform pie (Online News, Forums, ...) |

The remaining ~170 columns are platform-specific engagement metrics (Facebook/X/Instagram/
Reddit/TikTok/YouTube/LinkedIn/Bluesky/Threads counts), author/geo fields, and redaction/
workflow flags — mostly empty for news rows. A small schema-only sample preserving all 186
columns is at [`../samples/raw_export_SAMPLE_schema_only.xlsx`](../samples/raw_export_SAMPLE_schema_only.xlsx).

## Chart-data exports

Small "Chart Data Export" files. Metadata block in rows 1–9 (`Dimension 1`, `Dimension 2`,
`Aggregate`, dates, `Project Id`, `Location`, `Query Id`), then the data table from row 11.

| Export | Dimensions | Data table |
|--------|-----------|-----------|
| **Volume** (`chart_volume_monthsxqueries`) | months × queries | r11 = month labels (`A=months`, B..E), r12+ = model → 4 monthly volumes. **4th column = new month** (UTC/GMT `23:00:00` label). |
| **Country** (`chart_country_queriesxcountries`) | queries × countries | r11 header (`queries`, carline, `population`), r12+ = country → volume (+ population). |
| **Sentiment** (`chart_sentiment_queriesxsentiment`) | queries × sentiment | r11 = model names, r12 = positive, r13 = neutral, r14 = negative (counts). |
| **Platform** (`chart_platform_...`) | **should be** carline × pageTypes | r11 header, r12+ = platform → volume. |

### ⚠️ The platform-export gotcha

If the platform export's dimensions come out as **`all` / `all`** it is *unfiltered* — the rows
will be boolean-component names (e.g. `Melon_two`, `qsrfi_b3`) with 0 volume, **not** real
platforms. Re-export filtered to the carline. The bad case is captured for reference at
[`../samples/chart-exports/chart_platform_UNFILTERED_all-all_BAD-EXAMPLE.xlsx`](../samples/chart-exports/chart_platform_UNFILTERED_all-all_BAD-EXAMPLE.xlsx);
`scripts/validate_inputs.py` detects it.
