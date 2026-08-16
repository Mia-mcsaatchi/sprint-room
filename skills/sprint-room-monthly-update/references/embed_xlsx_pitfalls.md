# Embed XLSX Pitfalls

These two bugs caused "PowerPoint found a problem with content" errors during the original Explorer build. Avoid them.

## Pitfall 1: sharing one xlsx between multiple charts

Symptom: PowerPoint refuses to open the file with the "found a problem" dialog. python-pptx and LibreOffice both open the file fine (they're more lenient).

Cause: I wired three charts (volume, country, sentiment) at one embedded xlsx because they all reference the same sheet name `'Explorer July'`. OOXML's chart-data-embedding convention is **one embedded file per chart**, not one file per sheet. PowerPoint treats shared embeddings as malformed.

**Fix:** create one xlsx per chart. The sheet name can still match what the chart's `<c:f>` formulas expect — it just needs to be in a dedicated workbook.

```
ppt/embeddings/
├── Microsoft_Excel_Worksheet.xlsx     ← chart 9 (volume) backing data
├── Microsoft_Excel_Worksheet1.xlsx    ← chart 10 (country) backing data
├── Microsoft_Excel_Worksheet2.xlsx    ← chart 11 (sentiment) backing data
└── Microsoft_Excel_Worksheet3.xlsx    ← chart 12 (platform) backing data
```

Even if multiple charts share the same logical "Explorer July" sheet, each chart gets a physical workbook that contains its own copy.

## Pitfall 2: duplicate `<Override>` entries in `[Content_Types].xml`

Symptom: same "PowerPoint found a problem" dialog.

Cause: when fixing pitfall 1, I added 4 new Override entries but the regex meant to remove the old 2 entries used `\s*<Override...>` — but the leading whitespace had already been consumed by the previous append operation, so the old entries didn't match the pattern. Result: 6 entries total, 2 of which were duplicates.

**Fix:** use a regex that matches with-or-without leading whitespace:

```python
ct = re.sub(r'\s*<Override PartName="/ppt/embeddings/Microsoft_Excel_Worksheet[0-9]*\.xlsx"[^>]*/>', '', ct)
```

Then add the four fresh entries. Verify uniqueness:

```python
import re
matches = re.findall(r'<Override PartName="/ppt/embeddings/[^"]+"[^>]*/>', ct)
assert len(matches) == len(set(matches)), "Duplicate entries in Content_Types"
```

## Pitfall 3: relationship Type for internal embedded xlsx

The original chart rels (linking to external Egnyte URLs) used:
```
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
TargetMode="External"
```

For internal embedded xlsx, the correct Type is:
```
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
```
(no TargetMode — internal targets don't get one)

## Pitfall 4: `<c:externalData>` inside the chart XML

The chart's `<c:externalData r:id="rId4">` element points to whatever the rel rId4 says. **Do not change this element** — it stays the same regardless of whether rId4 points to an external URL or an internal xlsx.

The `r:id` is the chart's pointer; the relationship is what gets rewired.

## Pitfall 5: sheet names + cell ranges must match chart formulas

Each chart has `<c:f>` formulas like `'Explorer July'!$A$12` or `Explorer!$B$12:$B$17`. These are interpreted against the backing xlsx, so:
- The xlsx must contain a sheet with that name (case-sensitive, including spaces)
- Cell values must be at the referenced positions

If you put the right values at the wrong cells, PowerPoint may open the file but the chart will display the wrong data when the user clicks "Edit Data".

## Pitfall 6: pack with validation ON

`pack.py --validate false` was a habit during iterative debugging. **Always pack with validation on** for the final output. Validation catches:
- Empty `<c:ext .../>` elements (need a child or removal)
- Unresolved relationships (rels pointing to deleted parts)
- Missing Override entries in Content_Types
- Duplicate or conflicting entries

If validation fails, fix the underlying issue rather than skipping validation. The validation tool exists exactly to catch what PowerPoint's stricter parser will reject.

## Sanity check after packing

```python
import zipfile, re
with zipfile.ZipFile('output.pptx') as z:
    # 1. All chart rels point to existing parts
    for n in [9,10,11,12]:
        rels = z.read(f'ppt/charts/_rels/chart{n}.xml.rels').decode()
        m = re.search(r'<Relationship Id="rId4"[^>]*Target="([^"]+)"', rels)
        target = m.group(1).replace('../', 'ppt/')
        assert target in z.namelist(), f"chart{n} rId4 → missing {target}"
    
    # 2. No duplicate Override entries
    ct = z.read('[Content_Types].xml').decode()
    overrides = re.findall(r'<Override PartName="/ppt/embeddings/[^"]+"', ct)
    assert len(overrides) == len(set(overrides)), "Duplicate Override in Content_Types"
    
    # 3. Each chart has a distinct embedded file
    targets = []
    for n in [9,10,11,12]:
        rels = z.read(f'ppt/charts/_rels/chart{n}.xml.rels').decode()
        m = re.search(r'<Relationship Id="rId4"[^>]*Target="([^"]+)"', rels)
        targets.append(m.group(1))
    assert len(targets) == len(set(targets)), "Multiple charts share one embedding"
```
