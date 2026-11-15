"""rollforward_from_spec.py — roll a delivered Sprint Room deck forward one month from a JSON spec.

Built for the Kuga cycle (Sep -> Oct 2026) when only the raw Bulk Mentions export was supplied
(no chart-data exports). Works on decks with the delivered Kuga layout:

  slide1 = title, slide2 = last month (charts 1-4), slide3 = month before (charts 5-8)
  chart order: volume (sheet 'Kuga - 2'), country pie ('Sheet1'), platform pie ('Sheet1'),
  sentiment combo ('Kuga': K11:P21, Positive Pre row 19, Negative Pre row 21)

Approach: overwrite the *oldest* slide's parts (slide3, charts 5-8, embeddings 4-7, notesSlide2)
with the new month and reorder sldIdLst to title -> new -> last. Reusing part names means no
orphaned rels / Content-Types overrides (the build_skeleton.py bugs) and the notesSlide link survives.
The last-month slide is left byte-identical. Sentiment diamonds are read from last month's
*published* bars (chart4), never from new raw data.

Usage:
    unzip last_month.pptx -d last/
    python3 rollforward_from_spec.py last/ new_unpacked/ spec.json
    # then: <pptx skill>/scripts/clean.py new_unpacked/ ; zip -X -r -D from inside new_unpacked/ ;
    #       <pptx skill>/scripts/office/validate.py out.pptx ; soffice --headless --convert-to pdf out.pptx

spec.json keys:
    volume     {"months": [4 labels], "series": {"Tiguan": [4 values], ...}}  (names = xlsx col A)
    country    {"UK": n, "ES": n, ...}         (keys = pie categories)
    platform   {"Forums": n, "Online News": n, ...}
    sentiment  {"Kuga": {"positive": f, "neutral": f, "negative": f}, ...}  (keys = bar categories)
    bullets    3 lists of segments; a segment is "text" or ["linked text", "url"] (bold hyperlink)
    notes      list of lines; prefix "## " for a bold heading line
    prev_subtitle_month / subtitle_month   e.g. "EU5 September 26" / "EU5 October 26"
    prev_month_long / month_long           e.g. "September 2026" / "October 2026"
    report_date                            e.g. "15.11.2026"
"""
import json, re, shutil, sys
from pathlib import Path
from lxml import etree
import openpyxl
from xml.sax.saxutils import quoteattr

SRC = Path(sys.argv[1])      # unpacked last-month deck
OUT = Path(sys.argv[2])      # output unpacked dir
D = json.load(open(sys.argv[3]))  # numbers + text

if OUT.exists():
    shutil.rmtree(OUT)
shutil.copytree(SRC, OUT)
P = OUT / "ppt"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
A = "{%s}" % NS["a"]; C = "{%s}" % NS["c"]; R = "{%s}" % NS["r"]


def save(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------- copy parts 1-4 -> 5-8
for i in range(1, 5):
    j = i + 4
    shutil.copy(P / f"charts/chart{i}.xml", P / f"charts/chart{j}.xml")
    shutil.copy(P / f"charts/colors{i}.xml", P / f"charts/colors{j}.xml")
    shutil.copy(P / f"charts/style{i}.xml", P / f"charts/style{j}.xml")
    shutil.copy(P / f"theme/themeOverride{i}.xml", P / f"theme/themeOverride{j}.xml")
    src_x = "Microsoft_Excel_Worksheet.xlsx" if i == 1 else f"Microsoft_Excel_Worksheet{i-1}.xlsx"
    shutil.copy(P / f"embeddings/{src_x}", P / f"embeddings/Microsoft_Excel_Worksheet{j-1}.xlsx")
    rels = (P / f"charts/_rels/chart{j}.xml.rels").read_text()
    assert f"Microsoft_Excel_Worksheet{j-1}.xlsx" in rels and f"style{j}.xml" in rels, rels


# ---------------------------------------------------------------- chart helpers
def set_cache(parent, tag, values, is_num):
    cache = parent.find(f".//{C}{'numCache' if is_num else 'strCache'}")
    for pt in cache.findall(f"{C}pt"):
        cache.remove(pt)
    cache.find(f"{C}ptCount").set("val", str(len(values)))
    for idx, v in enumerate(values):
        pt = etree.SubElement(cache, f"{C}pt", idx=str(idx))
        etree.SubElement(pt, f"{C}v").text = repr(float(v)) if is_num else str(v)


def series(tree):
    return list(tree.iter(f"{C}ser"))


def ser_name(ser):
    tx = ser.find(f"{C}tx")
    return "".join(tx.find(f".//{C}v").itertext()) if tx is not None and tx.find(f".//{C}v") is not None else ""


# Volume (chart5)
vol = D["volume"]
t = etree.parse(str(P / "charts/chart5.xml"))
for ser in series(t):
    name = ser_name(ser)
    set_cache(ser.find(f"{C}cat"), "cat", vol["months"], False)
    vals = vol["series"][name] if name != "Average" else [
        sum(vol["series"][k][m] for k in vol["series"]) / len(vol["series"]) for m in range(4)]
    set_cache(ser.find(f"{C}val"), "val", vals, True)
save(t, P / "charts/chart5.xml")
wb = openpyxl.load_workbook(P / "embeddings/Microsoft_Excel_Worksheet4.xlsx")
ws = wb["Kuga - 2"]
for m, col in enumerate("BCDE"):
    ws[f"{col}11"] = vol["months"][m]
    for r in range(12, 18):
        ws[f"{col}{r}"] = vol["series"][ws[f"A{r}"].value][m]
wb.save(P / "embeddings/Microsoft_Excel_Worksheet4.xlsx")


# Pies (chart6 country, chart7 platform)
def pie(chart, xlsx, data):
    t = etree.parse(str(P / f"charts/{chart}"))
    ser = series(t)[0]
    cats = [p.findtext(f"{C}v") for p in ser.find(f"{C}cat").iter(f"{C}pt")]
    set_cache(ser.find(f"{C}val"), "val", [data[c] for c in cats], True)
    for ml in t.iter(f"{C}manualLayout"):  # let labels re-place around new slices
        if ml.getparent().tag == f"{C}layout" and ml.getparent().getparent().tag == f"{C}dLbl":
            ml.getparent().remove(ml)
    save(t, P / f"charts/{chart}")
    wb = openpyxl.load_workbook(P / f"embeddings/{xlsx}")
    ws = wb["Sheet1"]
    for r in range(12, 12 + len(cats)):
        ws[f"B{r}"] = data[ws[f"A{r}"].value]
    wb.save(P / f"embeddings/{xlsx}")
    return cats


pie("chart6.xml", "Microsoft_Excel_Worksheet5.xlsx", D["country"])
pie("chart7.xml", "Microsoft_Excel_Worksheet6.xlsx", D["platform"])

# Sentiment (chart8): bars = new month; diamonds = last month's *published* bars (from chart4)
t4 = etree.parse(str(P / "charts/chart4.xml"))
prior = {}
for ser in series(t4):
    nm = ser_name(ser)
    if nm in ("positive", "negative"):
        cats = [p.findtext(f"{C}v") for p in ser.find(f"{C}cat").iter(f"{C}pt")]
        prior[nm] = [float(p.findtext(f"{C}v")) for p in ser.find(f"{C}val").iter(f"{C}pt")]
order = cats
t = etree.parse(str(P / "charts/chart8.xml"))
sent = D["sentiment"]
for ser in series(t):
    nm = ser_name(ser)
    if nm in ("positive", "neutral", "negative"):
        assert [p.findtext(f"{C}v") for p in ser.find(f"{C}cat").iter(f"{C}pt")] == order
        set_cache(ser.find(f"{C}val"), "val", [sent[m][nm] for m in order], True)
    elif nm == "Positive Pre":
        set_cache(ser.find(f"{C}xVal"), "x", prior["positive"], True)
    elif nm == "Negative Pre":
        set_cache(ser.find(f"{C}xVal"), "x", [1 - v for v in prior["negative"]], True)
    else:
        raise SystemExit(f"unexpected series {nm!r}")
# strip any frozen manual data-label text
for rich in list(t.iter(f"{C}rich")):
    if rich.getparent().tag == f"{C}tx" and rich.getparent().getparent().tag == f"{C}dLbl":
        dl = rich.getparent().getparent(); dl.remove(rich.getparent())
save(t, P / "charts/chart8.xml")
wb = openpyxl.load_workbook(P / "embeddings/Microsoft_Excel_Worksheet7.xlsx")
ws = wb["Kuga"]
for k, col in enumerate("KLMNOP"):
    m = ws[f"{col}11"].value
    assert m == order[k]
    ws[f"{col}12"] = sent[m]["positive"]; ws[f"{col}13"] = sent[m]["neutral"]; ws[f"{col}14"] = sent[m]["negative"]
    ws[f"{col}19"] = prior["positive"][k]; ws[f"{col}21"] = 1 - prior["negative"][k]
wb.save(P / "embeddings/Microsoft_Excel_Worksheet7.xlsx")


# ---------------------------------------------------------------- slide text helpers
def runs_for(segments, size):
    """segments: list of str | [text, url]. Returns list of (run element, url or None)."""
    out = []
    for seg in segments:
        r = etree.Element(f"{A}r")
        rpr = etree.SubElement(r, f"{A}rPr", lang="en-GB", sz=str(size), dirty="0")
        if isinstance(seg, list):
            rpr.set("b", "1")
            out.append((r, seg[1]))
            text = seg[0]
        else:
            out.append((r, None)); text = seg
        etree.SubElement(r, f"{A}t").text = text
    return out


# New slide3 from slide2
s = etree.parse(str(P / "slides/slide2.xml"))
for tnode in s.iter(f"{A}t"):
    if tnode.text and D["prev_subtitle_month"] in tnode.text:
        tnode.text = tnode.text.replace(D["prev_subtitle_month"], D["subtitle_month"])
body = None
for sp in s.iter("{%s}sp" % NS["p"]):
    if "Mentions volume" in "".join(sp.itertext()):
        body = sp.find(".//p:txBody", NS)
paras = body.findall(f"{A}p")
assert len(paras) == 3
rels_links = []
next_rid = 101
for p, segs in zip(paras, D["bullets"]):
    for r in p.findall(f"{A}r"):
        p.remove(r)
    for r, url in runs_for(segs, 1200):
        if url:
            rid = f"rId{next_rid}"; next_rid += 1
            hl = etree.SubElement(r.find(f"{A}rPr"), f"{A}hlinkClick"); hl.set(f"{R}id", rid)
            rels_links.append((rid, url))
        p.append(r)
save(s, P / "slides/slide3.xml")

rel = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout10.xml"/>',
       '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/>']
for rid, url in rels_links:
    rel.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target={quoteattr(url)} TargetMode="External"/>')
for n, ch in zip(range(8, 12), range(5, 9)):
    rel.append(f'<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{ch}.xml"/>')
rel.append("</Relationships>")
(P / "slides/_rels/slide3.xml.rels").write_text("".join(rel))
# sanity: slide2 graphicFrames use rId8-11 for chart1-4 in the same order
s2rels = (P / "slides/_rels/slide2.xml.rels").read_text()
for n, ch in zip(range(8, 12), range(1, 5)):
    assert f'Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{ch}.xml"' in s2rels

# Notes: notesSlide2 from notesSlide1 template
n = etree.parse(str(P / "notesSlides/notesSlide1.xml"))
nb = None
for sp in n.iter("{%s}sp" % NS["p"]):
    ph = sp.find(".//p:ph", NS)
    if ph is not None and ph.get("type") == "body":
        nb = sp.find(".//p:txBody", NS)
for p in nb.findall(f"{A}p"):
    nb.remove(p)
for line in D["notes"]:
    p = etree.SubElement(nb, f"{A}p")
    r = etree.SubElement(p, f"{A}r")
    rpr = etree.SubElement(r, f"{A}rPr", lang="en-GB", sz="1100", dirty="0")
    if line.startswith("## "):
        rpr.set("b", "1"); line = line[3:]
    etree.SubElement(r, f"{A}t").text = line
save(n, P / "notesSlides/notesSlide2.xml")
nrels = (P / "notesSlides/_rels/notesSlide2.xml.rels").read_text()
assert "../slides/slide3.xml" in nrels, nrels

# Title slide
s1 = etree.parse(str(P / "slides/slide1.xml"))
for tnode in s1.iter(f"{A}t"):
    if tnode.text and D["prev_month_long"] in tnode.text:
        tnode.text = tnode.text.replace(D["prev_month_long"], D["month_long"])
    if tnode.text and re.fullmatch(r"\d\d\.\d\d\.\d{4}", tnode.text.strip()):
        tnode.text = D["report_date"]
save(s1, P / "slides/slide1.xml")

# Slide order: title -> new month (slide3, rId7) -> last month (slide2, rId6)
pres = (P / "presentation.xml").read_text()
pres, k = re.subn(r'(<p:sldId id="\d+" r:id="rId6"/>)(<p:sldId id="\d+" r:id="rId7"/>)', r"\2\1", pres)
assert k == 1
(P / "presentation.xml").write_text(pres)
print("built", OUT)
