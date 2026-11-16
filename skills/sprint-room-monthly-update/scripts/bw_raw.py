#!/usr/bin/env python3
"""
bw_raw.py - Load a Brandwatch "Bulk Mentions Download" xlsx into a pandas DataFrame.

Self-contained so the skill works wherever it is deployed (it does not depend on the
repo's tools/ folder). Same approach as tools/bw_read.py - see docs/tooling-patterns.md:

  * BW exports do not convert via LibreOffice ("source file could not be loaded") and keep
    text in inline strings, so read xl/worksheets/sheet1.xml directly with lxml iterparse.
  * The header is the row where column B == "Query Name" (rows above it are the
    Report/Brand/From/To/Label metadata block). Never hardcode the row number.

Adds a `Month` column bucketed in Europe/London time. BW's `Date` column is UTC, while the
export window and the dashboard charts use the project's UK timezone (e.g. "From: Mon Aug 31
23:00:00 UTC" = 1 Sep 00:00 BST), so a mention at 23:30 UTC on the last day of a month belongs
to the NEXT month.
"""
import re, zipfile

import pandas as pd
from lxml import etree

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _sheet_name(z):
    for n in z.namelist():
        if re.search(r"worksheets/sheet1\.xml$", n):
            return n
    cands = sorted(n for n in z.namelist() if "worksheets/sheet" in n)
    if not cands:
        raise RuntimeError("No worksheet XML found in archive")
    return cands[0]


def _shared_strings(z):
    try:
        data = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = etree.fromstring(data)
    return ["".join(si.itertext()) for si in root.iter(f"{NS}si")]


def _cell_value(c, sst):
    t = c.get("t")
    if t == "inlineStr":
        isel = c.find(f"{NS}is")
        return "".join(isel.itertext()) if isel is not None else ""
    v = c.find(f"{NS}v")
    if v is None:
        return None
    if t == "s" and sst:
        return sst[int(v.text)]
    return v.text


def iter_rows(path):
    """Yield each row as {column_letter: value}. Streams; safe for 80MB+ exports."""
    with zipfile.ZipFile(path) as z:
        sst = _shared_strings(z)
        with z.open(_sheet_name(z)) as fh:
            for _, el in etree.iterparse(fh, events=("end",), tag=f"{NS}row"):
                row = {}
                for c in el.iter(f"{NS}c"):
                    ref = c.get("r")
                    if ref:
                        row[re.match(r"[A-Z]+", ref).group()] = _cell_value(c, sst)
                yield row
                el.clear()


def load_raw(path, tz="Europe/London"):
    """Return the raw export as a DataFrame with Date (tz-aware) and Month columns."""
    hdr, recs = None, []
    for row in iter_rows(path):
        if hdr is None:
            if (row.get("B") or "").strip() == "Query Name":
                hdr = row
            continue
        recs.append({hdr.get(col, col): val for col, val in row.items()})
    if hdr is None:
        raise RuntimeError(f"{path}: header row (column B == 'Query Name') not found - "
                           "is this a Bulk Mentions Download?")
    df = pd.DataFrame(recs, columns=list(hdr.values()))
    df = df[df["Query Name"].fillna("").str.strip() != ""].copy()
    df["Query Name"] = df["Query Name"].str.strip()
    d = pd.to_datetime(df["Date"].astype(str).str.replace(r"\.0$", "", regex=True),
                       errors="coerce", utc=True)
    df["Date"] = d.dt.tz_convert(tz) if tz else d
    df["Month"] = df["Date"].dt.month
    for col in ("Title", "Full Text", "Url", "Sentiment", "Page Type", "Country"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
    return df


if __name__ == "__main__":
    import sys
    frame = load_raw(sys.argv[1])
    print(frame.shape)
    print(frame.groupby(["Query Name", "Month"]).size().unstack(fill_value=0))
