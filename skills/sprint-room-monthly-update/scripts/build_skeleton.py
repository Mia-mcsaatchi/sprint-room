#!/usr/bin/env python3
"""
Build the new-month deck skeleton from the last-month deck.

Steps performed:
1. Unpack last_month.pptx
2. Duplicate slide2 (last month) → slide4 (new month)
3. Reorder presentation.xml sldIdLst: title → new → last (drop older slides)
4. Duplicate chart1–4 → chart9–12 (and colors/styles/themeOverrides)
5. Repoint slide4 chart rels at chart9–12
6. Update title slide month + date
7. Update slide4 subtitle to new month + carline
8. Register chart9–12 + xml/colors/styles parts in Content_Types

Usage:
    python3 build_skeleton.py --last-month last.pptx --carline Capri \\
        --new-month "April 2026" --report-date "10.05.2026" --out skeleton/
"""
import argparse, subprocess, os, shutil, re, sys, zipfile

PPTX_SCRIPTS = os.environ.get('PPTX_SCRIPTS', '/mnt/skills/public/pptx/scripts')

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--last-month', required=True)
    ap.add_argument('--carline', required=True)
    ap.add_argument('--new-month', required=True, help='e.g. "April 2026"')
    ap.add_argument('--report-date', required=True, help='e.g. "10.05.2026"')
    ap.add_argument('--out', required=True)
    args = ap.parse_args()

    out_dir = args.out
    os.makedirs(out_dir, exist_ok=True)
    unpacked = os.path.join(out_dir, 'unpacked')
    if os.path.exists(unpacked):
        shutil.rmtree(unpacked)

    # 1. Unpack
    print(f"[1/8] Unpacking {args.last_month}...")
    zipfile.ZipFile(args.last_month).extractall(unpacked)

    # 2. Duplicate slide2 → new slide (appended at the end of sldIdLst)
    print("[2/8] Duplicating slide2...")
    r = subprocess.run(['python3', os.path.join(PPTX_SCRIPTS, 'add_slide.py'),
                        unpacked, 'slide2.xml'], capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stderr); sys.exit(1)
    m = re.search(r'Created ppt/slides/(slide\d+)\.xml', r.stdout)
    if not m:
        print("Could not parse new slide from add_slide output:\n" + r.stdout); sys.exit(1)
    new_slide = m.group(1)
    with open(os.path.join(unpacked, 'ppt/presentation.xml')) as f:
        new_slide_id, new_slide_rid = re.findall(
            r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', f.read())[-1]
    print(f"  → {new_slide}.xml id={new_slide_id}, rid={new_slide_rid}")

    # 3. Reorder presentation.xml sldIdLst
    print("[3/8] Reordering sldIdLst...")
    pres_path = os.path.join(unpacked, 'ppt/presentation.xml')
    with open(pres_path) as f:
        pres = f.read()
    # Get current sldIdLst
    cur_m = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', pres, re.DOTALL)
    cur = cur_m.group(1)
    slides = re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', cur)
    # Expected: [title, last, prev_last, ..., new_just_added]
    # We want: [title, new, last] — drop prev_last and earlier
    title = slides[0]
    last = slides[1]
    new = (new_slide_id, new_slide_rid)
    new_order = f'\n    <p:sldId id="{title[0]}" r:id="{title[1]}"/>\n    <p:sldId id="{new[0]}" r:id="{new[1]}"/>\n    <p:sldId id="{last[0]}" r:id="{last[1]}"/>\n  '
    pres = re.sub(r'(<p:sldIdLst>).*?(</p:sldIdLst>)', f'\\1{new_order}\\2', pres, flags=re.DOTALL)
    # Mirror in p14 sldIdLst
    p14_m = re.search(r'<p14:sldIdLst>(.*?)</p14:sldIdLst>', pres, re.DOTALL)
    if p14_m:
        new_p14 = f'\n            <p14:sldId id="{title[0]}"/>\n            <p14:sldId id="{new[0]}"/>\n            <p14:sldId id="{last[0]}"/>\n          '
        pres = re.sub(r'(<p14:sldIdLst>).*?(</p14:sldIdLst>)', f'\\1{new_p14}\\2', pres, flags=re.DOTALL)
    with open(pres_path, 'w') as f:
        f.write(pres)

    # 4. Duplicate chart1–4 → chart9–12 (plus colors, styles, theme override and
    #    embedded workbook - each chart must own its parts, never share them)
    print("[4/8] Duplicating charts 1–4 → 9–12...")
    charts_dir = os.path.join(unpacked, 'ppt/charts')
    copied_parts = []
    for n in [1,2,3,4]:
        new_n = n + 8
        shutil.copy(os.path.join(charts_dir, f'chart{n}.xml'),
                    os.path.join(charts_dir, f'chart{new_n}.xml'))
        for prefix in ['colors','style']:
            src = os.path.join(charts_dir, f'{prefix}{n}.xml')
            if os.path.exists(src):
                shutil.copy(src, os.path.join(charts_dir, f'{prefix}{new_n}.xml'))
        # Copy rels and fix internal refs
        rels_src = os.path.join(charts_dir, '_rels', f'chart{n}.xml.rels')
        rels_dst = os.path.join(charts_dir, '_rels', f'chart{new_n}.xml.rels')
        if os.path.exists(rels_src):
            shutil.copy(rels_src, rels_dst)
            with open(rels_dst) as f:
                rels = f.read()
            rels = rels.replace(f'colors{n}.xml', f'colors{new_n}.xml')
            rels = rels.replace(f'style{n}.xml', f'style{new_n}.xml')
            for kind in ('themeOverride', 'package'):
                tm = re.search(rf'Type="[^"]*/{kind}" Target="\.\./([^"]+)"', rels)
                if not tm:
                    continue
                old = tm.group(1)
                stem, ext = os.path.splitext(old)
                new = f'{re.sub(r"\d+$", "", stem)}{new_n}{ext}'
                shutil.copy(os.path.join(unpacked, 'ppt', old), os.path.join(unpacked, 'ppt', new))
                rels = rels.replace(f'Target="../{old}"', f'Target="../{new}"')
                if kind == 'themeOverride':
                    copied_parts.append(f'/ppt/{new}')
            with open(rels_dst, 'w') as f:
                f.write(rels)

    # 5. Repoint slide4 chart rels to chart9–12
    print(f"[5/8] Repointing {new_slide} rels at chart9–12...")
    slide4_rels = os.path.join(unpacked, f'ppt/slides/_rels/{new_slide}.xml.rels')
    with open(slide4_rels) as f:
        r4 = f.read()
    r4 = r4.replace('charts/chart1.xml','charts/chart9.xml')
    r4 = r4.replace('charts/chart2.xml','charts/chart10.xml')
    r4 = r4.replace('charts/chart3.xml','charts/chart11.xml')
    r4 = r4.replace('charts/chart4.xml','charts/chart12.xml')
    with open(slide4_rels,'w') as f:
        f.write(r4)

    # 6. Update title slide (slide1)
    print("[6/8] Updating title slide...")
    slide1 = os.path.join(unpacked, 'ppt/slides/slide1.xml')
    with open(slide1) as f:
        s1 = f.read()
    # Replace "Social Media <month> <year>" pattern
    s1 = re.sub(r'(<a:t>Social Media )[^<]+(</a:t>)', f'\\g<1>{args.new_month}\\g<2>', s1)
    # Replace any DD.MM.YYYY date pattern
    s1 = re.sub(r'<a:t>\d{2}\.\d{2}\.\d{4}</a:t>', f'<a:t>{args.report_date}</a:t>', s1, count=1)
    with open(slide1,'w') as f:
        f.write(s1)

    # 7. Update slide4 subtitle (e.g. "Capri – Social Media (EU7 April 2026)")
    print(f"[7/8] Updating {new_slide} subtitle...")
    slide4 = os.path.join(unpacked, f'ppt/slides/{new_slide}.xml')
    with open(slide4) as f:
        s4 = f.read()
    s4 = re.sub(r'(<a:t>[^<]+– Social Media \(EU7 )[^<)]+(\)</a:t>)',
                f'\\g<1>{args.new_month}\\g<2>', s4)
    # Carline prefix: replace whatever's before " – Social Media" with carline
    s4 = re.sub(r'<a:t>[^<–]+(\s+–\s+Social Media \(EU7 [^<)]+\))</a:t>',
                f'<a:t>{args.carline}\\g<1></a:t>', s4)
    with open(slide4,'w') as f:
        f.write(s4)

    # 8. Register chart9–12 in Content_Types
    print("[8/8] Registering new chart parts in Content_Types...")
    ct_path = os.path.join(unpacked, '[Content_Types].xml')
    with open(ct_path) as f:
        ct = f.read()
    new_entries = ''
    for n in [9,10,11,12]:
        new_entries += f'  <Override PartName="/ppt/charts/chart{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>\n'
        new_entries += f'  <Override PartName="/ppt/charts/style{n}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>\n'
        new_entries += f'  <Override PartName="/ppt/charts/colors{n}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>\n'
    for part in copied_parts:
        new_entries += f'  <Override PartName="{part}" ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/>\n'
    # Insert before </Types>
    ct = ct.replace('</Types>', new_entries + '</Types>')
    with open(ct_path,'w') as f:
        f.write(ct)

    print(f"\nSkeleton built at {unpacked}")
    print("Next: Stage 4 — replace chart data on chart9–12.")

if __name__ == '__main__':
    main()
