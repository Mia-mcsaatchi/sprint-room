#!/usr/bin/env python3
"""
bw_read.py — Reliable reader for Brandwatch xlsx exports (raw + chart-data).

WHY THIS EXISTS
---------------
Brandwatch exports do NOT parse with openpyxl:
  * sharedStrings.xml is effectively empty; text lives in INLINE strings
    (<is><t> elements) so openpyxl returns blanks / "0 is out of range".
  * LibreOffice CSV conversion is unreliable on these files too.
The reliable approach is direct zipfile + lxml iterparse over
xl/worksheets/sheet1.xml, handling both inline strings and numeric <v> cells.
For 80MB+ raw exports use the streaming iterator and call row.clear().

HEADER DETECTION
----------------
Do NOT trust a fixed header row number and do NOT trust the filename.
Scan for the row whose column B == "Query Name". Everything above it is the
metadata block (Report/Brand/From/To/Label...). Data starts on the next row.

USAGE
-----
    python3 bw_read.py FILE.xlsx --head 20         # print first N rows
    python3 bw_read.py FILE.xlsx --querydist        # Query Name distribution
    from bw_read import iter_rows, header_index, query_name_distribution
"""
import argparse, re, sys, zipfile
from collections import Counter
from lxml import etree

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _sheet_name(z):
    for n in z.namelist():
        if re.search(r"worksheets/sheet1\.xml$", n):
            return n
    cands = [n for n in z.namelist() if "worksheets/sheet" in n]
    if not cands:
        raise RuntimeError("No worksheet XML found in archive")
    return sorted(cands)[0]


def _cell_value(c):
    """Return the value of a <c> cell, handling inline strings and numerics."""
    t = c.get("t")
    if t == "inlineStr":
        isel = c.find(f"{NS}is")
        if isel is None:
            return None
        tel = isel.find(f"{NS}t")
        return tel.text if tel is not None else ""
    v = c.find(f"{NS}v")
    return v.text if v is not None else None


def _col_letters(ref):
    return re.match(r"[A-Z]+", ref).group()


def iter_rows(path):
    """Yield each row as a dict {column_letter: value}. Streams; safe for big files."""
    with zipfile.ZipFile(path) as z:
        sheet = _sheet_name(z)
        with z.open(sheet) as fh:
            for _, el in etree.iterparse(fh, events=("end",), tag=f"{NS}row"):
                row = {}
                for c in el.iter(f"{NS}c"):
                    ref = c.get("r")
                    if ref:
                        row[_col_letters(ref)] = _cell_value(c)
                yield row
                el.clear()  # critical for memory on 80MB+ exports


def header_index(path, key="Query Name", key_col="B"):
    """Return (0-based row index, header dict) of the row where key_col == key."""
    for i, row in enumerate(iter_rows(path)):
        if (row.get(key_col) or "").strip() == key:
            return i, row
    raise RuntimeError(f"Header row (col {key_col} == {key!r}) not found")


def records(path):
    """Yield data rows below the header as dicts keyed by header label."""
    hdr_i, hdr = None, None
    for i, row in enumerate(iter_rows(path)):
        if hdr is None:
            if (row.get("B") or "").strip() == "Query Name":
                hdr_i, hdr = i, row
            continue
        yield {hdr.get(col, col): val for col, val in row.items()}


def query_name_distribution(path):
    """Counter of the 'Query Name' column — validate content, never the filename."""
    return Counter((r.get("Query Name") or "").strip() for r in records(path))


def load_dataframe(path, text_cols=("Title", "Full Text", "Url", "Domain", "Country",
                                     "Page Type", "Page Type Name", "Sentiment", "Language",
                                     "Author", "Query Name")):
    """Raw export -> pandas DataFrame keyed by header label, with Date/Month parsed.

    Use this instead of LibreOffice CSV conversion: on some exports (e.g. the Capri
    Aug-2026 file) soffice fails outright with "source file could not be loaded", and even
    when it works the 5-row metadata block lands above the real header. Columns are looked
    up by NAME — letters move between exports (177 vs 186 columns seen in the wild).

    Month bucketing: BW dates are UTC. A mention stamped 23:xx on the last day of the prior
    month is a local-time (BST/CEST) first-of-month mention, so it is bucketed forward.
    """
    import pandas as pd
    rows = [r for r in records(path) if any(v not in (None, "") for v in r.values())]
    df = pd.DataFrame(rows)
    if df.empty or "Date" not in df.columns:
        return pd.DataFrame(columns=["Query Name", "Date", "Month", "Year"])
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    local = df["Date"] + pd.Timedelta(hours=1)  # UTC -> UK/EU local (approx; only the 23:00 edge matters)
    df["Month"] = local.dt.month
    df["Year"] = local.dt.year
    return df


def _main():
    ap = argparse.ArgumentParser(description="Read a Brandwatch xlsx export reliably.")
    ap.add_argument("file")
    ap.add_argument("--head", type=int, metavar="N", help="print first N rows raw")
    ap.add_argument("--querydist", action="store_true",
                    help="print Query Name distribution (validates the export contents)")
    a = ap.parse_args()

    if a.querydist:
        for name, n in query_name_distribution(a.file).most_common():
            print(f"{n:>8}  {name}")
        return

    n = a.head or 12
    for i, row in enumerate(iter_rows(a.file)):
        compact = " | ".join(f"{c}={v}" for c, v in row.items() if v not in (None, ""))
        print(f"r{i+1}: {compact}")
        if i + 1 >= n:
            break


if __name__ == "__main__":
    sys.exit(_main())
