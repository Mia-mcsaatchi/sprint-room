# Calculations

Every quantitative fact is computed deterministically in Python **before** any language model
is involved. The LLM's role is language only — never arithmetic. This is the core stability
guarantee of the whole pipeline.

---

## 1. Volume (line chart — CT1 / Chart 1)

- **Metric:** `COUNTUNIQUE(URL)` per model per month.
- **Window:** rolling **4 months**; the oldest month is dropped each cycle.
- **Average row:** mean of the competitor rows, written as a live formula
  (`=AVERAGE(B12:B17)` style) so it recomputes if an analyst edits a value. **Never overwrite
  the Average row when shifting months left** — it contains formulae.

### The UTC/GMT 4th-column quirk (important)

Brandwatch labels the newest month's column with a timestamp like `YYYY-MM-DD 23:00:00`
(e.g. `2026-03-31 23:00:00.0`). That is a UTC-vs-GMT artefact — **that 4th column IS the new
reporting month** (April in the example), not the last day of the prior month. Do not "correct"
it.

*(Confirmed in the sample volume export: columns are Jan / Feb / Mar / `2026-03-31 23:00` = Apr.)*

---

## 2. Sentiment (combo bar + diamonds — CT3 / sentiment chart)

Two layers on one chart:

- **Bars = current month.** 100% stacked horizontal bar: negative / neutral / positive as
  a share of that model's mentions for the month.
- **Diamonds = prior month.** An XY scatter series overlaid on the bars, one diamond per model.

### Diamond geometry — get this exactly right

For each model, the diamonds are positioned as:

| Series | X value | Why |
|--------|---------|-----|
| **positive pre** | prior-month **positive %** | plots directly |
| **Negative Pre** | **`1 − prior-month negative %`** | the bar is stacked from the right, so the negative diamond sits at the neutral/negative *boundary*, not at the raw negative rate |

- **Y coordinates** (six row positions): `[3, 8.75, 14.6, 20.5, 26.3, 32.1]`.
- **Marker:** diamond; colour-coded (green = negative series, red = positive series). Not
  achievable in Google Sheets — requires Excel + python-pptx / direct OOXML.

### Authority for the diamond (prior-month) values

The prior month's diamonds come from the **prior month's *published* chart** (extract from
`last_month.pptx`'s `chart3.xml`) — **NOT** re-derived from current raw data. Those are the
numbers the client actually saw; re-deriving would create a mismatch. Frozen as delivered.

### Series-name casing

Match the chart's existing series names exactly. In the **Kuga** deck they are capitalised:
`'Positive Pre'` / `'Negative Pre'`. Casing mismatches silently break the lookup map.

### Data labels

Strip any manual `<a:t>X%</a:t>` data-label overrides before/after updating. Old decks carried
stale labels frozen from a previous month showing wrong values. Auto-labels are reliable;
manual overrides are not.

---

## 3. Country pie (CT2a)

- Carline × Country, **current month only**, plotted as share %.
- After updating values, **remove `<c:manualLayout>` entries** on the data labels so they
  auto-position around the new slice sizes. Without this, labels overlap when slices change.

---

## 4. Platform pie (CT2b)

- Carline × Platform, **current month only** (typical categories: Online News, X, Forums,
  Blogs, Review, Tumblr).
- **Must be the carline-specific slice.** If the export's dimensions come out as `all` / `all`,
  it is the *unfiltered* pull and is wrong — re-export filtered to the carline.
  (See `samples/chart-exports/chart_platform_UNFILTERED_all-all_BAD-EXAMPLE.xlsx` for what the
  bad export looks like: rows are boolean-component names with 0 volume.)
- Same `<c:manualLayout>` label-stripping as the country pie.

---

## 5. Drivers, deltas & thresholds (insight layer — I1–I4)

Computed deterministically, then handed to the analyst as an evidence menu.

- **Volume Δ:** MoM change in count and %; top platform driver; top country driver.
- **Sentiment Δ:** neg/pos %-point shift MoM; which model moved most in each direction.
- **Story clustering:** news + blog only (highest signal); Twitter/forum sampled (max ~50 rows
  per bucket). Cluster by normalised title. Top 3 topics per model × sentiment bucket, each
  with `count`, `example_url`, `summary`.

### Thresholds & rules

| Rule | Value |
|------|-------|
| **Cluster threshold** | a driver with volume **< 30** is a "cluster". When negative-driver volume `< 30` for a model, use Val's exact phrase: *"no clusters of issues were identified as drivers"*. |
| **Driver-on-slide minimum** | a candidate driver with **< 5** mentions is dropped from the slide and moved to notes. |
| **£100 rule** | never hard-exclude a row unless you'd bet **£100** it's noise. Ambiguous rows stay in the file, flagged for analyst review. |
| **Country-breakdown heuristic** | break a story down by country (in the notes) if it has **> ~30** mentions in any single country. |

### Sentiment base

Sentiment percentages are always a share of **that model's total mentions for the month**.

---

## Golden rule

> All quantitative facts are locked before any LLM is called. The LLM structures language
> around numbers Python already computed and topics already clustered — it is never asked to
> invent or to do arithmetic. Zero hallucination risk on counts.
