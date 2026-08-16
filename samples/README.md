# Samples

Small fixtures for reference and for testing scripts. **No full raw exports** live here — they're
large and contain scraped third-party content (see the repo `.gitignore`).

## `chart-exports/` — real, tiny "Chart Data Export" files

From an Explorer (EU7) April-2026 cycle:

| File | What it demonstrates |
|------|----------------------|
| `chart_volume_monthsxqueries.xlsx` | 4-month rolling volume; note the 4th column's `23:00:00` UTC/GMT label = the new month |
| `chart_country_queriesxcountries.xlsx` | carline × country with a population column |
| `chart_sentiment_queriesxsentiment.xlsx` | model × pos/neu/neg counts (March = prior month for an April build) |
| `chart_platform_UNFILTERED_all-all_BAD-EXAMPLE.xlsx` | **the failure case** — `all`/`all` dimensions, rows are boolean components with 0 volume. This is what a wrongly-unfiltered platform export looks like; re-export filtered to the carline |

## `raw_export_SAMPLE_schema_only.xlsx`

**Table structure only** — the metadata block + the full **186-column** header, with **no data
rows**. Raw data changes every month and isn't the point; the column layout is. Column meanings
and example values are documented in
[`../docs/data-dictionary.md`](../docs/data-dictionary.md).

> Real BW exports use inline strings and need [`../tools/bw_read.py`](../tools/bw_read.py) to
> read (openpyxl returns blanks) — see [`../docs/tooling-patterns.md`](../docs/tooling-patterns.md).

Try:

```bash
python3 ../tools/bw_read.py chart-exports/chart_volume_monthsxqueries.xlsx --head 14
```
