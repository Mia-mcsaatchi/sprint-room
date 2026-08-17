#!/usr/bin/env python3
"""
Pack the unpacked deck (zip), validate it, run integrity checks, and try to render a QA PDF.

The current pptx skill has no pack.py — a .pptx is a zip of the unpacked tree — and validation
moved to scripts/office/validate.py. Integrity checks are chart-number-agnostic (they scan every
chart's rId4, not a hardcoded 9–12) so this works for any carline/layout. PDF render is best-effort
(LibreOffice import filters are broken in some sandboxes); a failure there does not fail the build.

Usage:
    python3 pack_and_qa.py unpacked/ output.pptx [--original last_month.pptx]
"""
import argparse, subprocess, os, re, sys, zipfile

PPTX = '/mnt/skills/public/pptx/scripts'

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('unpacked')
    ap.add_argument('output')
    ap.add_argument('--original', default=None, help='template deck to baseline schema checks against')
    args = ap.parse_args()

    # 1. Clean orphaned parts
    print("[1/4] Cleaning orphaned files...")
    subprocess.run(['python3', f'{PPTX}/clean.py', args.unpacked], capture_output=True, text=True)

    # 2. Pack (zip) then validate with the skill's validator
    print("[2/4] Packing (zip) + validating...")
    if os.path.exists(args.output):
        os.remove(args.output)
    with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _d, files in os.walk(args.unpacked):
            for f in files:
                p = os.path.join(root, f)
                z.write(p, os.path.relpath(p, args.unpacked))
    vcmd = ['python3', f'{PPTX}/office/validate.py', args.output]
    if args.original:
        vcmd += ['--original', args.original]
    r = subprocess.run(vcmd, capture_output=True, text=True)
    print(r.stdout)
    if r.returncode != 0:
        print("VALIDATION FAILED:\n" + r.stderr)
        print("Do NOT skip validation — fix the underlying issue. See references/embed_xlsx_pitfalls.md")
        sys.exit(1)

    # 3. Integrity checks (chart-number agnostic)
    print("[3/4] Integrity checks...")
    with zipfile.ZipFile(args.output) as z:
        ct = z.read('[Content_Types].xml').decode()
        overrides = re.findall(r'<Override PartName="[^"]+"', ct)
        dupes = [o for o in set(overrides) if overrides.count(o) > 1]
        if dupes:
            print(f"  ✗ Duplicate Content_Types overrides: {dupes}"); sys.exit(1)
        # every chart's embedded package (rId4 → embeddings/*.xlsx) must exist and be unique
        embeds = []
        for name in z.namelist():
            m = re.match(r'ppt/charts/_rels/(chart\d+)\.xml\.rels$', name)
            if not m:
                continue
            rels = z.read(name).decode()
            em = re.search(r'Target="(\.\./embeddings/[^"]+)"', rels)
            if em:
                tgt = em.group(1).replace('../', 'ppt/')
                if tgt not in z.namelist():
                    print(f"  ✗ {m.group(1)} embedding {tgt} missing from package"); sys.exit(1)
                embeds.append(tgt)
        if len(embeds) != len(set(embeds)):
            print(f"  ✗ Charts share an embedding: {embeds}"); sys.exit(1)
        print(f"  ✓ Content_Types unique; {len(embeds)} chart embedding(s) exist and are distinct")

    # 4. Best-effort PDF/PNG render for visual QA
    print("[4/4] Rendering QA (best-effort)...")
    outdir = os.path.dirname(os.path.abspath(args.output)) or '.'
    subprocess.run(['python3', f'{PPTX}/office/soffice.py', '--headless', '--convert-to', 'pdf',
                    '--outdir', outdir, args.output], capture_output=True, text=True)
    pdf = os.path.join(outdir, os.path.basename(args.output).replace('.pptx', '.pdf'))
    if os.path.exists(pdf):
        try:
            import pymupdf
            doc = pymupdf.open(pdf)
            for i in range(doc.page_count):
                doc[i].get_pixmap(dpi=120).save(os.path.join(outdir, f'qa_slide{i+1}.png'))
            print(f"  ✓ QA images: {outdir}/qa_slide*.png")
        except Exception as e:
            print(f"  (PDF rendered but PNG step skipped: {e})")
    else:
        print("  (LibreOffice could not render a PDF here — QA in PowerPoint. Deck is valid.)")

    print(f"\nFinal deck: {args.output}")
    print("LibreOffice-only quirks that are NOT deck bugs: 1–2 diamonds shown, bullet-3 clip, "
          "stacked right-edge labels. Chase real data anomalies only.")

if __name__ == '__main__':
    main()
