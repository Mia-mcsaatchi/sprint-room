#!/usr/bin/env python3
"""
write_notes.py — create/attach a speaker-notes slide for the new-month slide.

build_skeleton_raw drops the cloned slide's notes ref (so last month's notes aren't shared).
This clones an existing notesSlide as a template, replaces its body paragraphs with the new
month's overflow (country breakdowns, platform-by-country, sentiment-by-platform, competitor
caveats — per editorial-standards.md), wires the rels/Content-Types, and links it to the slide.

Notes markdown: one paragraph per non-empty line; a leading "- " makes it a bullet; inline
[LINK:phrase](url) becomes a hyperlink run.

Usage:
    python3 write_notes.py unpacked/ --target-slide slide6.xml --notes notes.md \
        [--template-notes notesSlide2.xml]
"""
import argparse, os, re, glob

def xesc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def build_paras(lines, rid_start, rels):
    out, rid = [], rid_start
    for raw in lines:
        line = raw.rstrip('\n')
        if not line.strip():
            continue
        bullet = line.lstrip().startswith('- ')
        text = line.lstrip()[2:] if bullet else line
        runs, pos = [], 0
        for m in re.finditer(r'\[LINK:([^\]]+)\]\(([^)]+)\)', text):
            if m.start() > pos:
                runs.append(f'<a:r><a:rPr lang="en-GB" sz="1100" dirty="0"/><a:t xml:space="preserve">{xesc(text[pos:m.start()])}</a:t></a:r>')
            this = f'rId{rid}'; rels.append((this, m.group(2))); rid += 1
            runs.append(f'<a:r><a:rPr lang="en-GB" sz="1100" b="1" dirty="0"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{this}"/></a:rPr><a:t>{xesc(m.group(1))}</a:t></a:r>')
            pos = m.end()
        if pos < len(text):
            runs.append(f'<a:r><a:rPr lang="en-GB" sz="1100" dirty="0"/><a:t xml:space="preserve">{xesc(text[pos:])}</a:t></a:r>')
        pPr = '<a:pPr marL="171450" indent="-171450"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>' if bullet else '<a:pPr/>'
        out.append(f'<a:p>{pPr}{"".join(runs)}</a:p>')
    return '\n'.join(out), rid

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('unpacked')
    ap.add_argument('--target-slide', required=True)
    ap.add_argument('--notes', required=True)
    ap.add_argument('--template-notes', default=None)
    a = ap.parse_args()
    up = a.unpacked
    ndir = f'{up}/ppt/notesSlides'
    existing = glob.glob(f'{ndir}/notesSlide*.xml')
    template = a.template_notes and f'{ndir}/{a.template_notes}' or (existing[0] if existing else None)
    if not template or not os.path.exists(template):
        raise SystemExit('No template notesSlide found to clone')
    nums = [int(re.search(r'notesSlide(\d+)\.xml', p).group(1)) for p in existing]
    new_n = max(nums) + 1
    xml = open(template, encoding='utf-8').read()

    # replace the body placeholder's paragraphs
    lines = open(a.notes, encoding='utf-8').read().splitlines()
    rels = []
    paras, _ = build_paras(lines, 100, rels)  # rid space 100+ won't clash with rId1/2
    def repl_body(m):
        sp = m.group(0)
        if not re.search(r'<p:ph[^>]*type="body"', sp):
            return sp
        return re.sub(r'(<a:p>.*</a:p>)', paras, sp, count=1, flags=re.S)
    xml = re.sub(r'<p:sp>.*?</p:sp>', repl_body, xml, flags=re.S)
    open(f'{ndir}/notesSlide{new_n}.xml', 'w', encoding='utf-8').write(xml)

    # rels: notesMaster + target slide + hyperlinks
    rel_lines = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>',
                 f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/{a.target_slide}"/>']
    for rid, url in rels:
        rel_lines.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/>')
    open(f'{ndir}/_rels/notesSlide{new_n}.xml.rels', 'w', encoding='utf-8').write(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rel_lines) + '</Relationships>')

    # Content-Types override
    ct_path = f'{up}/[Content_Types].xml'
    ct = open(ct_path, encoding='utf-8').read()
    ov = f'<Override PartName="/ppt/notesSlides/notesSlide{new_n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
    if ov not in ct:
        ct = ct.replace('</Types>', ov + '</Types>'); open(ct_path, 'w', encoding='utf-8').write(ct)

    # link into the target slide's rels with a fresh rId
    srels_path = f'{up}/ppt/slides/_rels/{a.target_slide}.rels'
    srels = open(srels_path, encoding='utf-8').read()
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', srels)]
    nrid = f'rId{max(used)+1 if used else 1}'
    add = f'<Relationship Id="{nrid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{new_n}.xml"/>'
    srels = srels.replace('</Relationships>', add + '</Relationships>')
    open(srels_path, 'w', encoding='utf-8').write(srels)
    print(f"Wrote notesSlide{new_n}.xml and linked it to {a.target_slide} ({nrid}); {len(rels)} note hyperlink(s)")

if __name__ == '__main__':
    main()
