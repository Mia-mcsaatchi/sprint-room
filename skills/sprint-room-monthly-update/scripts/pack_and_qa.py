#!/usr/bin/env python3
"""
Pack the unpacked deck with validation ON, render to PDF for QA, run integrity checks.

Usage:
    python3 pack_and_qa.py unpacked/ output.pptx --original last_month.pptx
"""
import argparse, subprocess, os, re, sys, zipfile

PPTX_SCRIPTS = os.environ.get('PPTX_SCRIPTS', '/mnt/skills/public/pptx/scripts')


def pack(unpacked, output):
    """Zip the unpacked tree into output ([Content_Types].xml first)."""
    if os.path.exists(output):
        os.remove(output)
    files = []
    for root, _, names in os.walk(unpacked):
        for n in names:
            full = os.path.join(root, n)
            files.append((os.path.relpath(full, unpacked).replace(os.sep, '/'), full))
    files.sort(key=lambda f: (f[0] != '[Content_Types].xml', f[0]))
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for arc, full in files:
            z.write(full, arc)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('unpacked')
    ap.add_argument('output')
    ap.add_argument('--original', help='source deck, to baseline schema validation against')
    args = ap.parse_args()

    # Clean
    print("[1/4] Cleaning orphaned files...")
    r = subprocess.run(['python3', os.path.join(PPTX_SCRIPTS, 'clean.py'), args.unpacked],
                       capture_output=True, text=True)
    print(r.stdout)

    # Pack with validation ON
    print("[2/4] Packing with strict validation...")
    pack(args.unpacked, args.output)
    cmd = ['python3', os.path.join(PPTX_SCRIPTS, 'office', 'validate.py'), args.output]
    if args.original:
        cmd += ['--original', args.original]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print("VALIDATION FAILED:")
        print(r.stdout)
        print(r.stderr)
        print("\nDo NOT skip validation. Fix the underlying issue first. See "
              "references/embed_xlsx_pitfalls.md and references/pptx_structure.md")
        sys.exit(1)
    print(r.stdout)

    # Integrity checks
    print("[3/4] Running integrity checks...")
    with zipfile.ZipFile(args.output) as z:
        # Check no duplicate Override entries
        ct = z.read('[Content_Types].xml').decode()
        overrides = re.findall(r'<Override PartName="[^"]+"', ct)
        dupes = [o for o in set(overrides) if overrides.count(o)>1]
        if dupes:
            print(f"  ✗ Duplicate Content_Types overrides: {dupes}")
            sys.exit(1)
        # Check each chart embedding exists and is unique
        targets = []
        for n in [9,10,11,12]:
            rels = z.read(f'ppt/charts/_rels/chart{n}.xml.rels').decode()
            m = re.search(r'<Relationship Id="rId4"[^>]*Target="([^"]+)"', rels)
            if m:
                target = m.group(1).replace('../','ppt/')
                if target not in z.namelist():
                    print(f"  ✗ chart{n} rId4 → {target} (does not exist in package)")
                    sys.exit(1)
                targets.append(m.group(1))
        if len(targets) != len(set(targets)):
            print(f"  ✗ Multiple charts share one embedding: {targets}")
            sys.exit(1)
        print(f"  ✓ Content_Types unique; chart embeddings exist and are distinct")

    # Render
    print("[4/4] Rendering to PDF for visual QA...")
    out_dir = os.path.dirname(os.path.abspath(args.output))
    r = subprocess.run(['python3', os.path.join(PPTX_SCRIPTS, 'office', 'soffice.py'),
                        '--headless', '--convert-to', 'pdf', '--outdir', out_dir, args.output],
                       capture_output=True, text=True)
    print(r.stdout)
    pdf_path = os.path.splitext(os.path.abspath(args.output))[0] + '.pdf'
    if os.path.exists(pdf_path):
        prefix = os.path.join(out_dir, 'slide_qa')
        subprocess.run(['pdftoppm', '-jpeg', '-r', '150', pdf_path, prefix],
                       capture_output=True, text=True)
        print(f"\nPDF: {pdf_path}\nQA images: {prefix}-N.jpg (slide 2 = the new month)")
    else:
        print("PDF render failed:", r.stderr)

    print(f"\nFinal deck: {args.output}")
    print("Reminder: LibreOffice render shows known quirks (only 1-2 diamonds, bullet 3 clip). "
          "These are LO bugs; PowerPoint renders correctly.")

if __name__ == '__main__':
    main()
