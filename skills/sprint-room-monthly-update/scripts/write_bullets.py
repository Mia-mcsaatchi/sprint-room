#!/usr/bin/env python3
"""
Write the three bullets into slide4 from an analyst-confirmed menu.

The menu is a markdown file with this exact structure:

    ## bullet1
    > Mentions volume decreased ...
    
    ## bullet2
    > {Carline}'s decline was driven by news (...). Top driver: [LINK:Stoll & Sauer recall PR](url) (35 mentions, German regional news), {explanation}. Secondary driver: [LINK:Explorer Collection](url) (10 mentions, tailing off ...
    
    ## bullet3
    > {Carline} sentiment ...
    
    ## hyperlinks
    | label | url |
    |---|---|
    | Stoll & Sauer recall PR | https://... |
    | Explorer Collection | https://... |

Linked phrases in bullets use the syntax `[LINK:phrase](url)` or `[LINK:phrase]` if the URL
is in the hyperlinks table (matched by label).

The script parses each bullet, creates relationship entries in slide4.xml.rels, and writes
the bullet paragraphs with inline <a:hlinkClick> on the linked runs.

Usage:
    python3 write_bullets.py path/to/slide4.xml --menu menu.md
"""
import argparse, re, os

def parse_menu(path):
    with open(path) as f:
        content = f.read()
    sections = re.split(r'^## ', content, flags=re.MULTILINE)
    parsed = {}
    for s in sections[1:]:
        head, body = s.split('\n', 1) if '\n' in s else (s, '')
        head = head.strip()
        if head.startswith('bullet'):
            # Extract the > quote line(s)
            quotes = re.findall(r'^>\s*(.*)$', body, re.MULTILINE)
            parsed[head] = ' '.join(quotes).strip()
        elif head == 'hyperlinks':
            # Parse markdown table
            links = {}
            for line in body.splitlines():
                line = line.strip()
                if line.startswith('|') and not line.startswith('|---'):
                    parts = [p.strip() for p in line.strip('|').split('|')]
                    if len(parts) == 2 and parts[0] not in ('label','linked phrase') and parts[1].startswith('http'):
                        links[parts[0]] = parts[1]
            parsed['hyperlinks'] = links
    return parsed

def parse_bullet_with_links(text, links_table):
    """Returns list of (text_segment, hyperlink_url_or_None).
    Recognises [LINK:phrase](url) and [LINK:phrase] (resolves via table)."""
    segments = []
    pos = 0
    pattern = re.compile(r'\[LINK:([^\]]+)\](?:\(([^)]+)\))?')
    for m in pattern.finditer(text):
        if m.start() > pos:
            segments.append((text[pos:m.start()], None))
        phrase = m.group(1)
        url = m.group(2) or links_table.get(phrase)
        if not url:
            print(f"  WARN: no URL found for [LINK:{phrase}]; treating as plain text")
            segments.append((phrase, None))
        else:
            segments.append((phrase, url))
        pos = m.end()
    if pos < len(text):
        segments.append((text[pos:], None))
    return segments

def xml_escape(s):
    return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def build_bullet_xml(segments, rid_offset_start, rels_to_add):
    """Build the <a:p> for one bullet. Returns (xml, next_rid_offset)."""
    runs = []
    rid = rid_offset_start
    for text, url in segments:
        if url:
            this_rid = f'rId{rid}'
            rels_to_add.append((this_rid, url))
            rid += 1
            r = (f'<a:r><a:rPr lang="en-GB" sz="1200" b="1" dirty="0">'
                 f'<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{this_rid}"/>'
                 f'</a:rPr><a:t>{xml_escape(text)}</a:t></a:r>')
        else:
            r = (f'<a:r><a:rPr lang="en-GB" sz="1200" dirty="0"/>'
                 f'<a:t xml:space="preserve">{xml_escape(text)}</a:t></a:r>')
        runs.append(r)
    bullet = ('<a:p><a:pPr marL="171450" indent="-171450">'
              '<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
              '<a:buChar char="•"/></a:pPr>' + ''.join(runs) + '</a:p>')
    return bullet, rid

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('slide_xml')
    ap.add_argument('--menu', required=True)
    args = ap.parse_args()

    menu = parse_menu(args.menu)
    links_table = menu.get('hyperlinks', {})

    with open(args.slide_xml) as f:
        slide_xml = f.read()

    # Find the highest rId currently in use in slide rels
    rels_path = args.slide_xml.replace('/slides/','/slides/_rels/') + '.rels'
    with open(rels_path) as f:
        rels_xml = f.read()
    existing_rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
    next_rid = max(existing_rids) + 1 if existing_rids else 1

    rels_to_add = []
    bullets_xml = []
    for key in ['bullet1','bullet2','bullet3']:
        text = menu.get(key, '')
        if not text: continue
        segs = parse_bullet_with_links(text, links_table)
        b, next_rid = build_bullet_xml(segs, next_rid, rels_to_add)
        bullets_xml.append(b)

    new_bullets_block = '\n          '.join(bullets_xml)

    # Replace the three bullets inside the summary text frame
    # Find the txBody that contains the existing bullets — the one with "Mentions volume" text
    # We'll find the first <a:p> after a "Mentions volume" indicator and replace through the
    # third <a:p>'s closing tag.
    # Simpler: find the existing three bullet <a:p>...</a:p> sequence and replace it.
    # The summary text-frame is anchored by id="53" or similar; we look for the txBody
    # containing the volume bullet.
    txbody_m = re.search(
        r'(<p:txBody>(?:(?!</p:txBody>).)*?Mentions volume(?:(?!</p:txBody>).)*?</p:txBody>)',
        slide_xml, re.DOTALL)
    if not txbody_m:
        print("ERROR: could not locate summary txBody"); return
    old_txbody = txbody_m.group(1)
    # Inside this txBody, replace from first <a:p>...</a:p> through last </a:p>
    # Keep the bodyPr / lstStyle. Replace all <a:p>...</a:p> entries.
    new_txbody = re.sub(r'(<a:lstStyle[^/]*/>\s*)(<a:p>.*</a:p>)\s*(</p:txBody>)',
                        lambda m: m.group(1) + new_bullets_block + '\n        ' + m.group(3),
                        old_txbody, count=1, flags=re.DOTALL)
    if new_txbody == old_txbody:
        # Fallback: just replace all <a:p>...</a:p>
        # Strip all paragraphs and add ours
        new_txbody = re.sub(r'(<a:p>.*?</a:p>\s*)+', new_bullets_block, old_txbody, count=1, flags=re.DOTALL)
    slide_xml = slide_xml.replace(old_txbody, new_txbody, 1)

    with open(args.slide_xml, 'w') as f:
        f.write(slide_xml)

    # Add hyperlink rels to slide rels
    new_rels = ''
    for rid, url in rels_to_add:
        new_rels += f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/>\n'
    rels_xml = rels_xml.replace('</Relationships>', new_rels + '</Relationships>')
    with open(rels_path,'w') as f:
        f.write(rels_xml)
    print(f"Added {len(rels_to_add)} hyperlink rels to {rels_path}")

if __name__ == '__main__':
    main()
