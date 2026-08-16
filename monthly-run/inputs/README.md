# inputs/

Per-cycle Brandwatch exports go here, under `inputs/<CARLINE>/<YYYY-MM>/`. Create the folder
with `python3 run.py newmonth --carline <C> --month <YYYY-MM>`.

**Contents are git-ignored** — these are client data and change every month. Only this README
is tracked. Expected filenames per cycle:

```
last_month.pptx      the previous month's delivered deck (carry-forward source)
raw.xlsx             raw Bulk Mentions export covering BOTH months
chart_volume.xlsx    volume × month (rolling 4)
chart_country.xlsx   carline × country
chart_platform.xlsx  carline × platform  (must be carline-filtered, NOT all/all)
chart_sentiment.xlsx carline × sentiment (prior month, for the diamonds)
```

See [`../RUNBOOK.md`](../RUNBOOK.md) for the full sequence and
[`../../docs/data-dictionary.md`](../../docs/data-dictionary.md) for the export formats.
