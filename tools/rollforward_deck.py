#!/usr/bin/env python3
"""
rollforward_deck.py - roll a 3-slide Sprint Room deck forward one month, in place.

For decks whose charts ALREADY embed their own backing xlsx (relationship type
.../package -> ../embeddings/Microsoft_Excel_Worksheet*.xlsx), e.g. the Puma deck
from mid-2026 onwards. Instead of duplicating charts into new part names (the
skill's build_skeleton.py route, which needs manual Content-Types/rels/notes fixes
every cycle), this ROTATES the existing parts:

    before:  title | latest month (N-1) | oldest month (N-2)
    after :  title | new month (N)      | latest month (N-1, byte-for-byte)

  1. The latest slide, its notes, its four charts and their embedded xlsx are copied
     over the oldest slide's parts (same part names, so [Content_Types].xml and
     presentation.xml need no edits and nothing is orphaned).
  2. The latest slide's parts are then refilled with the new month:
       - volume line: rolling window shifts one month left; the new column comes
         from the spec. Earlier months are carried from the published chart.
       - country / platform pies: values from the spec (category order is kept,
         so per-slice colours stay stable month to month).
       - sentiment combo: bars from the spec; diamonds = the latest slide's
         PUBLISHED bars (positive %, and 1 - negative %), series renamed to
         "Positive <prior month>" / "Negative <prior month>".
       - every value is also written into the chart's embedded xlsx at the exact
         cells its <c:f> formulas reference (cells holding formulas are kept).
       - subtitle, title-slide month/date, bullets (inline bold hyperlinks),
         speaker notes, docProps/app.xml slide titles.

Chart roles are detected from content (line = volume, bar+scatter = sentiment,
pie with "Online News"/"X" categories = platform, other pie = country), so the
per-carline chart order does not matter.

Usage:
    python3 rollforward_deck.py --last last_month.pptx --spec spec.json --out new.pptx

Spec (JSON):
{
  "month": "October", "year": 2026, "prior_month": "September",
  "report_date": "15.11.2026",
  "volume":   {"new": {"Puma": 1094, "Yaris Cross": 555, ...}},   # keys = series names in the chart
  "country":  {"UK": 250, "ES": 415, ...},                         # keys = pie categories
  "platform": {"Online News": 621, "X": 299, ...},
  "sentiment": {"Ford Puma": [pos, neu, neg], ...},                # shares (0-1), keys = bar categories
  "bullets": [ [["plain text", null], ["linked text", "https://..."], ...], ... ],
  "notes":   [ {"text": "Insight 1 - ...", "bold": true}, {"text": "..."}, ... ]
}
"""
import argparse, copy, io, json, re, shutil, sys, tempfile, zipfile
from pathlib import Path

from lxml import etree
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_HYPER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
A, C, P, R = (f"{{{NS[k]}}}" for k in ("a", "c", "p", "r"))


# ---------------------------------------------------------------- xml helpers
def load(path):
    return etree.parse(str(path), etree.XMLParser(remove_blank_text=False))


def save(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def rels_path(part):
    part = Path(part)
    return part.parent / "_rels" / (part.name + ".rels")


def read_rels(part):
    rp = rels_path(part)
    t = load(rp)
    return t, {e.get("Id"): e for e in t.getroot()}


def resolve(part, target):
    return (Path(part).parent / target).resolve()


# ---------------------------------------------------------------- deck layout
def slide_order(root):
    pres = root / "ppt" / "presentation.xml"
    t = load(pres)
    _, rels = read_rels(pres)
    ids = t.getroot().find("p:sldIdLst", NS)
    return [resolve(pres, rels[s.get(f"{R}id")].get("Target")) for s in ids]


def slide_charts(slide):
    """Chart parts of a slide, in on-slide (shape tree) order."""
    t = load(slide)
    _, rels = read_rels(slide)
    out = []
    for ch in t.getroot().iter(f"{C}chart"):
        out.append(resolve(slide, rels[ch.get(f"{R}id")].get("Target")))
    return out


def chart_package(chart):
    _, rels = read_rels(chart)
    for e in rels.values():
        if e.get("Type").endswith("/package"):
            return resolve(chart, e.get("Target"))
    sys.exit(f"{chart.name}: no embedded workbook (package relationship). "
             "This tool is for decks with embedded chart data.")


def notes_of(slide):
    _, rels = read_rels(slide)
    for e in rels.values():
        if e.get("Type").endswith("/notesSlide"):
            return resolve(slide, e.get("Target"))
    return None


def chart_role(chart):
    root = load(chart).getroot()
    pa = root.find(".//c:plotArea", NS)
    kinds = {etree.QName(x).localname for x in pa}
    if "lineChart" in kinds:
        return "volume"
    if "barChart" in kinds and "scatterChart" in kinds:
        return "sentiment"
    if "pieChart" in kinds or "doughnutChart" in kinds:
        cats = [v.text for v in root.findall(".//c:cat//c:v", NS)]
        return "platform" if {"Online News", "X", "Forums"} & set(cats) else "country"
    return "unknown"


# ---------------------------------------------------------------- chart caches
def cache_vals(ref):
    """Values of a c:strRef/c:numRef (or literal) element, by idx order."""
    pts = sorted(ref.iter(f"{C}pt"), key=lambda p: int(p.get("idx")))
    return [p.find("c:v", NS).text for p in pts]


def set_cache(parent, values, numeric):
    """Rewrite the strCache/numCache under a c:cat/c:val/c:xVal/c:yVal/c:tx element."""
    cache = parent.find(".//c:numCache", NS) if numeric else parent.find(".//c:strCache", NS)
    if cache is None:
        raise RuntimeError(f"no {'num' if numeric else 'str'}Cache under {parent.tag}")
    for pt in cache.findall("c:pt", NS):
        cache.remove(pt)
    cache.find("c:ptCount", NS).set("val", str(len(values)))
    for i, v in enumerate(values):
        pt = etree.SubElement(cache, f"{C}pt", idx=str(i))
        etree.SubElement(pt, f"{C}v").text = (repr(float(v)) if isinstance(v, float) else str(v))
    return cache


def formula(parent):
    f = parent.find(".//c:f", NS)
    return f.text if f is not None else None


class Workbook:
    """Embedded xlsx; writes values at the cells a chart's <c:f> refs point to."""

    def __init__(self, path):
        self.path = path
        self.wb = openpyxl.load_workbook(path)

    def write(self, ref, values):
        if not ref:
            return
        sheet, rng = ref.rsplit("!", 1)
        sheet = sheet.strip("'")
        ws = self.wb[sheet]
        c1, r1, c2, r2 = range_boundaries(rng.replace("$", ""))
        cells = [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        for (r, c), v in zip(cells, values):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # keep live formulas (e.g. the Average row)
            cell.value = v

    def replace_text(self, old, new):
        for ws in self.wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value == old:
                        cell.value = new

    def save(self):
        self.wb.save(self.path)


def num(v):
    f = float(v)
    return int(f) if f.is_integer() else f


# ---------------------------------------------------------------- chart updaters
def update_volume(chart, wb, spec):
    t = load(chart)
    root = t.getroot()
    sers = root.findall(".//c:lineChart/c:ser", NS)
    new_vals, months = {}, None
    names = [cache_vals(s.find("c:tx", NS))[0] for s in sers]
    for s, name in zip(sers, names):
        cat = s.find("c:cat", NS)
        old_m = cache_vals(cat)
        months = old_m[1:] + [spec["month"]]
        if name.lower() == "average":
            continue
        if name not in spec["volume"]["new"]:
            sys.exit(f"volume: spec has no value for series {name!r} (have {list(spec['volume']['new'])})")
        old_v = [num(v) for v in cache_vals(s.find("c:val", NS))]
        new_vals[name] = old_v[1:] + [spec["volume"]["new"][name]]
    if "Average" in names or "average" in [n.lower() for n in names]:
        comp = list(new_vals.values())
        new_vals[names[[n.lower() for n in names].index("average")]] = [
            round(sum(col) / len(col), 1) for col in zip(*comp)]
    for s, name in zip(sers, names):
        cat, val = s.find("c:cat", NS), s.find("c:val", NS)
        set_cache(cat, months, numeric=False)
        set_cache(val, new_vals[name], numeric=True)
        wb.write(formula(cat), months)
        wb.write(formula(val), new_vals[name])
    save(t, chart)
    return {"months": months, **new_vals}


def update_pie(chart, wb, values, label):
    t = load(chart)
    ser = t.getroot().find(".//c:pieChart/c:ser", NS)
    if ser is None:
        ser = t.getroot().find(".//c:doughnutChart/c:ser", NS)
    cat, val = ser.find("c:cat", NS), ser.find("c:val", NS)
    cats = cache_vals(cat)
    missing = [c for c in cats if c not in values]
    extra = [k for k in values if k not in cats]
    if missing:
        sys.exit(f"{label}: spec missing categories {missing}")
    if extra:
        print(f"  NOTE {label}: spec categories not in chart (not plotted): {extra}")
    vals = [values[c] for c in cats]
    set_cache(val, vals, numeric=True)
    wb.write(formula(val), vals)
    for lay in ser.iter(f"{C}manualLayout"):  # let labels re-place around new slices
        lay.getparent().remove(lay)
    save(t, chart)
    return dict(zip(cats, vals))


def update_sentiment(chart, wb, spec):
    t = load(chart)
    root = t.getroot()
    bars = root.findall(".//c:barChart/c:ser", NS)
    by_name = {cache_vals(s.find("c:tx", NS))[0].lower(): s for s in bars}
    cats = cache_vals(bars[0].find("c:cat", NS))
    # prior month = this chart's published bars (authority for the diamonds)
    prior_pos = [float(v) for v in cache_vals(by_name["positive"].find("c:val", NS))]
    prior_neg = [float(v) for v in cache_vals(by_name["negative"].find("c:val", NS))]
    missing = [c for c in cats if c not in spec["sentiment"]]
    if missing:
        sys.exit(f"sentiment: spec missing categories {missing}")
    for i, key in enumerate(["positive", "neutral", "negative"]):
        vals = [float(spec["sentiment"][c][i]) for c in cats]
        val = by_name[key].find("c:val", NS)
        set_cache(val, vals, numeric=True)
        wb.write(formula(val), vals)
    # strip frozen manual "NN%" label overrides on the bars
    for dl in root.iter(f"{C}dLbl"):
        tx = dl.find("c:tx", NS)
        if tx is not None and re.fullmatch(r"\s*\d+(\.\d+)?%\s*", "".join(tx.itertext())):
            dl.remove(tx)
    scat = root.findall(".//c:scatterChart/c:ser", NS)
    diamonds = {"positive": prior_pos, "negative": [1 - n for n in prior_neg]}
    for s in scat:
        tx = s.find("c:tx", NS)
        old_name = "".join(tx.itertext()).strip()
        kind = "positive" if old_name.lower().startswith("pos") else "negative"
        new_name = f"{kind.capitalize()} {spec['prior_month']}"
        if tx.find("c:strRef", NS) is not None:
            set_cache(tx, [new_name], numeric=False)
            wb.write(formula(tx), [new_name])
        else:
            tx.find("c:v", NS).text = new_name
        wb.replace_text(old_name, new_name)
        xv = s.find("c:xVal", NS)
        set_cache(xv, diamonds[kind], numeric=True)
        wb.write(formula(xv), diamonds[kind])
    save(t, chart)
    return {"cats": cats, "prior_pos": prior_pos, "prior_neg": prior_neg}


# ---------------------------------------------------------------- text helpers
def para_text(p):
    return "".join(x.text or "" for x in p.iter(f"{A}t"))


def set_runs(p, runs, rpr_tmpl, rels_tree, slide_rels, rid_seq):
    """Replace the runs of paragraph p with [(text, url|None), ...]."""
    for r in p.findall("a:r", NS) + p.findall("a:fld", NS) + p.findall("a:br", NS):
        p.remove(r)
    end = p.find("a:endParaRPr", NS)
    for text, url in runs:
        r = etree.Element(f"{A}r")
        rpr = copy.deepcopy(rpr_tmpl)
        for k in ("b", "err"):
            rpr.attrib.pop(k, None)
        for child in list(rpr):
            if etree.QName(child).localname == "hlinkClick":
                rpr.remove(child)
        if url:
            rid = next(rid_seq)
            rpr.set("b", "1")
            etree.SubElement(rpr, f"{A}hlinkClick").set(f"{R}id", rid)
            etree.SubElement(rels_tree.getroot(), f"{{{NS['rel']}}}Relationship",
                             Id=rid, Type=R_HYPER, Target=url, TargetMode="External")
        r.append(rpr)
        etree.SubElement(r, f"{A}t").text = text
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)


def rid_generator(existing):
    n = 100
    while True:
        n += 1
        if f"rId{n}" not in existing:
            yield f"rId{n}"


def update_slide_text(slide, spec):
    t = load(slide)
    root = t.getroot()
    rels_tree, rels = read_rels(slide)
    # drop the old month's hyperlink relationships (the carried-forward slide keeps its own copy)
    for rid, e in list(rels.items()):
        if e.get("Type") == R_HYPER:
            rels_tree.getroot().remove(e)
            del rels[rid]
    month_year = f"{spec['month']} {spec['year']}"
    sub_re = re.compile(r"^(.*– Social Media \(EU\d )([A-Za-z]+ \d{4})(\).*)$")
    hit = False
    for tnode in root.iter(f"{A}t"):
        m = sub_re.match(tnode.text or "")
        if m:
            tnode.text = m.group(1) + month_year + m.group(3)
            hit = True
    if not hit:
        sys.exit("subtitle '… – Social Media (EUx Month YYYY)' not found on the latest slide")
    # bullets: the text body holding >=3 bulleted paragraphs
    body = None
    for tb in root.iter(f"{P}txBody"):
        bps = [p for p in tb.findall("a:p", NS) if p.find("a:pPr/a:buChar", NS) is not None and para_text(p).strip()]
        if len(bps) >= 3:
            body = tb
            break
    if body is None:
        sys.exit("bullet text box not found on the latest slide")
    paras = [p for p in body.findall("a:p", NS)]
    tmpl_p = next(p for p in paras if p.find("a:pPr/a:buChar", NS) is not None)
    tmpl_rpr = next(r for r in tmpl_p.iter(f"{A}rPr"))
    for p in paras:
        body.remove(p)
    seq = rid_generator(rels)
    for runs in spec["bullets"]:
        p = copy.deepcopy(tmpl_p)
        set_runs(p, [tuple(x) for x in runs], tmpl_rpr, rels_tree, rels, seq)
        body.append(p)
    save(t, slide)
    save(rels_tree, rels_path(slide))


def update_notes(notes, spec):
    t = load(notes)
    root = t.getroot()
    body = None
    for tb in root.iter(f"{P}txBody"):
        if "Insight" in "".join(tb.itertext()) or len(tb.findall("a:p", NS)) > 3:
            body = tb
            break
    if body is None:
        sys.exit("notes body not found")
    paras = body.findall("a:p", NS)
    rpr = next(r for p in paras for r in p.iter(f"{A}rPr"))
    for p in paras:
        body.remove(p)
    for item in spec["notes"]:
        p = etree.SubElement(body, f"{A}p")
        r = etree.SubElement(p, f"{A}r")
        rp = copy.deepcopy(rpr)
        rp.attrib.pop("err", None)
        rp.attrib.pop("b", None)
        if item.get("bold"):
            rp.set("b", "1")
        r.append(rp)
        etree.SubElement(r, f"{A}t").text = item["text"]
    save(t, notes)


def update_title(slide, spec):
    t = load(slide)
    for tnode in t.getroot().iter(f"{A}t"):
        if re.match(r"^Social Media [A-Za-z]+ \d{4}$", tnode.text or ""):
            tnode.text = f"Social Media {spec['month']} {spec['year']}"
        elif re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", tnode.text or ""):
            tnode.text = spec["report_date"]
    save(t, slide)


def update_app(root, spec):
    app = root / "docProps" / "app.xml"
    if not app.exists():
        return
    x = app.read_text(encoding="utf-8")
    m = re.findall(r"– Social Media \((EU\d) ([A-Za-z]+ \d{4})\)", x)
    if len(m) >= 2:
        latest, oldest = m[0][1], m[1][1]
        new = f"{spec['month']} {spec['year']}"
        x = x.replace(f"({m[1][0]} {oldest})", "(@@OLD@@)", 1)
        x = x.replace(f"({m[0][0]} {latest})", f"({m[0][0]} {new})", 1)
        x = x.replace("(@@OLD@@)", f"({m[1][0]} {latest})", 1)
        app.write_text(x, encoding="utf-8")


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--last", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))

    tmp = Path(tempfile.mkdtemp(prefix="rollfwd_"))
    with zipfile.ZipFile(a.last) as z:
        names = z.namelist()
        z.extractall(tmp)
    root = tmp.resolve()

    order = slide_order(root)
    if len(order) != 3:
        sys.exit(f"expected 3 slides (title, latest, oldest); found {len(order)}")
    title, latest, oldest = order
    lc, oc = slide_charts(latest), slide_charts(oldest)
    if len(lc) != len(oc):
        sys.exit("latest/oldest slides have different chart counts")
    roles_l = [chart_role(c) for c in lc]
    roles_o = [chart_role(c) for c in oc]
    if roles_l != roles_o:
        sys.exit(f"chart roles differ between slides: {roles_l} vs {roles_o}")
    print("slides :", title.name, latest.name, oldest.name)
    print("charts :", ", ".join(f"{c.name}={r}" for c, r in zip(lc, roles_l)), "| carried to",
          ", ".join(c.name for c in oc))

    # 1. rotate latest -> oldest (charts, styles, colours, theme overrides, xlsx)
    for l, o in zip(lc, oc):
        _, lrels = read_rels(l)
        _, orels = read_rels(o)
        shutil.copyfile(l, o)
        ltargets = {e.get("Type"): resolve(l, e.get("Target")) for e in lrels.values()}
        for e in orels.values():
            src = ltargets.get(e.get("Type"))
            if src is not None and e.get("TargetMode") != "External":
                shutil.copyfile(src, resolve(o, e.get("Target")))
    lnotes, onotes = notes_of(latest), notes_of(oldest)
    shutil.copyfile(latest, oldest)
    if lnotes and onotes:
        shutil.copyfile(lnotes, onotes)
    # oldest slide rels := latest rels with chart + notes targets remapped
    lrel_tree, _ = read_rels(latest)
    new_rels = copy.deepcopy(lrel_tree)
    cmap = {l.name: o.name for l, o in zip(lc, oc)}
    for e in new_rels.getroot():
        tgt = e.get("Target")
        if e.get("Type").endswith("/chart"):
            e.set("Target", tgt.rsplit("/", 1)[0] + "/" + cmap[tgt.rsplit("/", 1)[1]])
        elif e.get("Type").endswith("/notesSlide") and onotes:
            e.set("Target", "../notesSlides/" + onotes.name)
    save(new_rels, rels_path(oldest))

    # 2. refill the latest slots with the new month
    summary = {}
    for c, role in zip(lc, roles_l):
        wb = Workbook(chart_package(c))
        if role == "volume":
            summary[role] = update_volume(c, wb, spec)
        elif role == "country":
            summary[role] = update_pie(c, wb, spec["country"], "country")
        elif role == "platform":
            summary[role] = update_pie(c, wb, spec["platform"], "platform")
        elif role == "sentiment":
            summary[role] = update_sentiment(c, wb, spec)
        else:
            sys.exit(f"{c.name}: unknown chart role")
        wb.save()
    update_slide_text(latest, spec)
    if lnotes:
        update_notes(lnotes, spec)
    update_title(title, spec)
    update_app(root, spec)

    # 3. repack ([Content_Types].xml first, original member order otherwise)
    out = Path(a.out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ["[Content_Types].xml"] + [n for n in names if n != "[Content_Types].xml"]:
            if not n.endswith("/"):
                z.write(root / n, n)
    shutil.rmtree(tmp)
    print(json.dumps(summary, indent=1, default=str))
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
