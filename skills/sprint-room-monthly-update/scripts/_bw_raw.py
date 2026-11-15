#!/usr/bin/env python3
"""
Shared loader for the raw Brandwatch "Bulk Mentions Download" export.

Why not LibreOffice/openpyxl: BW exports store text as inline strings and put a
metadata block (Report/Brand/From/To/Label) above the header, so a CSV round-trip
either produces nothing or the wrong header. We stream sheet1.xml with lxml and
detect the header as the row whose column B == "Query Name" (same approach as
tools/bw_read.py, kept here so the skill is self-contained).

Month bucketing: the Date column is UTC, but reporting months are UK months
(the export's From: is e.g. "Aug 31 23:00 UTC" = 1 Sep 00:00 BST). We convert to
Europe/London before taking the month, so a mention at 23:30 UTC on the last day
of a BST month lands in the next month - matching Brandwatch's own charts.
"""
import re, zipfile
from lxml import etree

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REPORT_TZ = "Europe/London"


def _sheet(z):
    names = [n for n in z.namelist() if re.search(r"worksheets/sheet\d+\.xml$", n)]
    if not names:
        raise RuntimeError("No worksheet XML found in archive")
    return next((n for n in names if n.endswith("sheet1.xml")), sorted(names)[0])


def _val(c):
    if c.get("t") == "inlineStr":
        t = c.find(f"{NS}is/{NS}t")
        return t.text if t is not None else ""
    v = c.find(f"{NS}v")
    return v.text if v is not None else None


def iter_records(path):
    """Yield data rows below the 'Query Name' header as dicts keyed by header label."""
    hdr = None
    with zipfile.ZipFile(path) as z, z.open(_sheet(z)) as fh:
        for _, el in etree.iterparse(fh, events=("end",), tag=f"{NS}row"):
            row = {}
            for c in el.iter(f"{NS}c"):
                ref = c.get("r")
                if ref:
                    row[re.match(r"[A-Z]+", ref).group()] = _val(c)
            el.clear()
            if hdr is None:
                if (row.get("B") or "").strip() == "Query Name":
                    hdr = row
                continue
            yield {hdr.get(col, col): v for col, v in row.items()}
    if hdr is None:
        raise RuntimeError(f"Header row (column B == 'Query Name') not found in {path}")


def load_raw_df(path):
    """Raw export -> DataFrame with Date (UTC), LocalDate, Month, Title, Full Text cleaned."""
    import pandas as pd
    df = pd.DataFrame(list(iter_records(path)))
    if "Query Name" not in df.columns:
        raise RuntimeError("raw export has no 'Query Name' column")
    df["Query Name"] = df["Query Name"].fillna("").astype(str).str.strip()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["LocalDate"] = df["Date"].dt.tz_localize("UTC").dt.tz_convert(REPORT_TZ)
    df["Month"] = df["LocalDate"].dt.month
    for col in ("Title", "Full Text", "Url", "Sentiment", "Page Type", "Country"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)
    return df
