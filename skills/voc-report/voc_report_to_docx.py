#!/usr/bin/env python3
"""
Convert a plain-text weekly EU Voice-of-Customer report into a formatted .docx
that imports cleanly into Google Docs (upload -> Open with Google Docs).

Usage:
    python3 voc_report_to_docx.py input.txt output.docx

Formatting spec (reverse-engineered from the user's reference Google Doc export
"Transit Custom & Competitors - Customer Highlights - w_c 15th June 2026.docx"):
    HIGHLIGHTS ................. 22pt bold
    Month/period line .......... 18pt bold   e.g. "This Month So Far (...)"
    Vehicle name ............... 16pt bold
    "During Last Week ..." ..... 14pt bold
    "... PHASE ... TOP DISCUSSED TOPICS" .. 12pt bold
    "... comments collected" ... 12pt BLUE (6D9EEB)
    body text .................. 12pt
    category bullet ............ 12pt  (level-0 bullet ●, indent 720/hanging 360)
    issue sub-bullet ........... 12pt  (level-1 bullet ○, indent 1440/hanging 360)
Body font is Arial throughout. Headings are plain bold Normal text, NOT Word heading
styles. Vertical spacing is created by empty paragraphs between blocks (see assemble)
-- there is no per-paragraph spacing, and NO blank line between consecutive bullets.
"""
import sys, re, zipfile
from xml.sax.saxutils import escape

# ---- line classification ----------------------------------------------
CAT_RE = re.compile(r'^(.*?)\(\d+\s+mentions?(?:\s+by\s+\d+\s+owners?)?\)\s*$')

def is_category(line):
    """A category label looks like 'Topic Name (N mentions [by M owners])' with
    nothing after the closing paren, a short head, and no sentence punctuation."""
    m = CAT_RE.match(line)
    if not m:
        return False
    head = m.group(1).strip()
    if not head:
        return False
    if ',' in line or ':' in line or ';' in line:
        return False
    words = [w for w in re.split(r'\s+', head) if w not in ('/', '&', '-')]
    return len(words) <= 6 and len(head) <= 55

def classify(lines):
    n = len(lines)
    out = []
    in_bullets = False
    for i, line in enumerate(lines):
        nxt = lines[i + 1] if i + 1 < n else ""
        prev = lines[i - 1] if i > 0 else ""
        if i == 0 or line.upper() == "HIGHLIGHTS":
            out.append(("title", line)); in_bullets = False
        elif prev.upper() == "HIGHLIGHTS":
            out.append(("month", line)); in_bullets = False
        elif nxt.startswith("During Last Week"):
            out.append(("h1", line)); in_bullets = False          # vehicle header
        elif line.startswith("During Last Week"):
            out.append(("week", line)); in_bullets = False
        elif "TOP DISCUSSED TOPICS" in line.upper():
            out.append(("ph", line)); in_bullets = False
        elif prev.startswith("During Last Week"):
            out.append(("collected", line)); in_bullets = False   # comments-collected line
        elif is_category(line):
            out.append(("cat", line)); in_bullets = True
        else:
            out.append(("issue" if in_bullets else "body", line))
    return out

# ---- OOXML builders ----------------------------------------------------
def rpr(bold=False, size=None, color=None, extra=""):
    props = ""
    if bold:  props += "<w:b/><w:bCs/>"
    if color: props += f'<w:color w:val="{color}"/>'
    if size:  props += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    props += extra
    return f"<w:rPr>{props}</w:rPr>" if props else ""

def run(text, bold=False, size=None, color=None):
    return f'<w:r>{rpr(bold, size, color)}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def block(text, bold=False, size=24, color=None):
    """A plain paragraph. The paragraph-mark rPr mirrors the run size (as the
    reference does) so the line height matches."""
    return f'<w:p><w:pPr>{rpr(bold, size)}</w:pPr>{run(text, bold, size, color)}</w:p>'

def bullet(text, level):
    ind = '<w:ind w:left="720" w:hanging="360"/>' if level == 0 else '<w:ind w:left="1440" w:hanging="360"/>'
    ppr = f'<w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr>{ind}{rpr(size=24)}</w:pPr>'
    return f'<w:p>{ppr}{run(text, size=24)}</w:p>'

# Sizes are half-points (pt*2). Headings are plain bold Normal text (no heading
# styles), matching the user's reference doc. Vertical gaps come from a single
# empty 12pt paragraph between blocks (see assemble); global line spacing is
# pinned to 1.15 (line=276) in docDefaults so gaps stay tight.
def para(kind, text):
    if kind == "title":     return block(text, bold=True, size=44)   # HIGHLIGHTS 22pt
    if kind == "month":     return block(text, bold=True, size=36)   # period line 18pt
    if kind == "h1":        return block(text, bold=True, size=32)   # vehicle name 16pt
    if kind == "week":      return block(text, bold=True, size=28)   # During Last Week 14pt
    if kind == "ph":        return block(text, bold=True, size=24)   # phase header 12pt
    if kind == "collected": return block(text, size=24, color="6D9EEB")  # 12pt blue
    if kind == "body":      return block(text, size=24)              # 12pt
    if kind == "cat":       return bullet(text, 0)                   # level-0 bullet 12pt
    if kind == "issue":     return bullet(text, 1)                   # level-1 bullet 12pt
    raise ValueError(kind)

EMPTY_P = f'<w:p><w:pPr>{rpr(size=24)}</w:pPr></w:p>'  # 12pt blank-line separator

def assemble(items):
    """Join paragraphs, inserting a blank line between adjacent blocks except
    between two consecutive bullets (categories/issues form one tight list)."""
    parts = []
    prev = None
    for k, t in items:
        if prev is not None and not (prev in ("cat", "issue") and k in ("cat", "issue")):
            parts.append(EMPTY_P)
        parts.append(para(k, t))
        prev = k
    return "".join(parts)

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en_GB"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>'''

NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#9679;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#9675;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

def build(text, out_path):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    items = classify(lines)
    body_xml = assemble(items)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        + body_xml +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
    return items

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: python3 voc_report_to_docx.py input.txt output.docx", file=sys.stderr)
        sys.exit(2)
    with open(sys.argv[1], encoding="utf-8") as f:
        text = f.read()
    items = build(text, sys.argv[2])
    from collections import Counter
    counts = Counter(k for k, _ in items)
    print("wrote", sys.argv[2])
    for k in ("title", "month", "h1", "week", "ph", "collected", "body", "cat", "issue"):
        if counts.get(k):
            print(f"  {k:10} {counts[k]}")
