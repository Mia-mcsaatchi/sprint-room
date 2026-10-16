#!/usr/bin/env python3
"""
Load a Brandwatch raw (Bulk Mentions) export into a pandas DataFrame.

Brandwatch xlsx exports keep text in inline strings and put a metadata block
(Report/Brand/From/To/Label) above the real header, so neither openpyxl nor a
LibreOffice CSV conversion reads them reliably. This streams the sheet XML
directly (same approach as tools/bw_read.py) and uses the row whose column B is
"Query Name" as the header.

Row dates are UTC. Months are bucketed in UK time (Europe/London), which is how
the export window is defined (e.g. From "Jul 31 23:00 UTC" = 1 Aug 00:00 BST).
"""
import re, zipfile
from lxml import etree

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _cell_value(c, shared):
    t = c.get("t")
    if t == "inlineStr":
        t = c.find(f"{NS}is/{NS}t")
        return t.text if t is not None else ""
    v = c.find(f"{NS}v")
    if v is None:
        return None
    return shared[int(v.text)] if t == "s" else v.text


def _shared_strings(z):
    # Brandwatch exports use inline strings; files re-saved by Excel/openpyxl
    # (e.g. a cleaned export) use the shared-strings table instead.
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = etree.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(si.itertext()) for si in root.iter(f"{NS}si")]


def _iter_rows(path):
    with zipfile.ZipFile(path) as z:
        shared = _shared_strings(z)
        sheet = next((n for n in z.namelist() if n.endswith("worksheets/sheet1.xml")), None)
        sheet = sheet or sorted(n for n in z.namelist() if "worksheets/sheet" in n)[0]
        with z.open(sheet) as fh:
            for _, el in etree.iterparse(fh, events=("end",), tag=f"{NS}row"):
                row = {}
                for c in el.iter(f"{NS}c"):
                    ref = c.get("r")
                    if ref:
                        row[re.match(r"[A-Z]+", ref).group()] = _cell_value(c, shared)
                yield row
                el.clear()


def load_raw(path, tz="Europe/London"):
    import pandas as pd
    hdr, recs = None, []
    for row in _iter_rows(path):
        if hdr is None:
            if (row.get("B") or "").strip() == "Query Name":
                hdr = row
            continue
        recs.append({hdr.get(col, col): val for col, val in row.items()})
    if hdr is None:
        raise RuntimeError(f"{path}: no header row with column B == 'Query Name'")
    df = pd.DataFrame(recs, columns=list(dict.fromkeys(hdr.values())))
    df["Date"] = (pd.to_datetime(df["Date"], errors="coerce")
                    .dt.tz_localize("UTC").dt.tz_convert(tz))
    df["Month"] = df["Date"].dt.month
    df["Query Name"] = df["Query Name"].fillna("").astype(str).str.strip()
    for col in ("Title", "Full Text", "Sentiment", "Page Type", "Country", "Url"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
    return df
