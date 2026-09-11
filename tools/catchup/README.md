# Catch-up build — when a month was never charted

Used for **Capri, August 2026**: the July slide was never produced, so the August deck needed
**two** new slides (July and August), built from the **last delivered deck (June)** plus the
**noise-cleaned raw export** — no Brandwatch chart exports at all. Result: title → August → July.

The three scripts here are the reference implementation. They are wired to the Capri template
(see "Template facts" below); read those constants from the last deck's chart XML before reusing
them for another carline.

## Sequence

```bash
# 0. cleaned workbook from run.py flags + the cleaning pass (Clean sheet = kept rows, original columns)
# 1. chart inputs for both months, sentiment under three retag tiers -> pick a tier with the analyst
python3 tools/catchup/chart_inputs_from_clean.py work/Capri/2026-08/Capri_Aug2026_raw_CLEAN.xlsx work/Capri/2026-08/CHART_INPUTS.md
# 2. content pack (Tier 1 retags baked in): volumes, pies, sentiment per model, country splits,
#    clusters with URLs, X/forum themes, biggest competitor mover
python3 tools/catchup/content_pack_from_clean.py work/Capri/2026-08/Capri_Aug2026_raw_CLEAN.xlsx work/Capri/2026-08/content_pack.json work/Capri/2026-08/CONTENT_PACK.md
# 3. write work/Capri/2026-08/slide_text.py  (JULY / AUGUST dicts: subtitle, bullets as (text,url) runs, notes) - client content, not committed
# 4. unzip the last delivered deck to work/Capri/2026-08/june_unpacked/ then build
PPTX_SCRIPTS=/path/to/pptx-skill/scripts python3 tools/catchup/build_deck_from_clean.py \
    --work monthly-run/work/Capri/2026-08 --last-pptx monthly-run/inputs/Capri/2026-08/june_deck.pptx \
    --out monthly-run/work/Capri/2026-08/Capri_Social_Media_August_2026_Sprint_Room.pptx --report-date 11.09.2026
```

`build_deck_from_clean.py` needs the pptx skill's `add_slide.py`, `clean.py` and `office/validate.py`.
It: duplicates the last slide twice (`add_slide.py`), clones chart1–4 → chart9–12 and chart13–16
(XML, colours, styles, rels, one backing xlsx **per chart** with sheet names matching the `<c:f>`
refs), creates a notesSlide per new slide, reorders `<p:sldIdLst>` (and the `p14` mirror), patches
every chart cache with lxml, rewrites bullets with inline `<a:hlinkClick>` rels, repositions the
manual line-end labels, edits the title slide, runs `clean.py`, zips from inside the dir and
validates against the original.

## Rules this path follows

- **Diamonds.** July slide: June diamonds = the *published* June chart (carry-forward). August slide:
  July diamonds = the cleaned July data — there is no published July chart, so data is the only source.
- **Same retag tier on both months**, so bars and diamonds are like-for-like.
- **Volume line.** Earlier months come from the last deck's chart cache; the two new months from
  cleaned data (unique URLs, local-time month boundary). State the methodology step in the notes:
  the earlier months were Brandwatch chart exports before the boolean was cleaned.
- Passing mentions stay in; only rows not about the model are excluded (`Excluded` sheet with reasons).

## Template facts — Capri deck (differ from the documented Explorer template)

| Part | Capri June-2026 deck |
|---|---|
| chart1 | volume line, 8 series in this order: Polestar 2, E-3008, EV6, ID.5, Q4 Sportback e-tron, Capri, **Average**, Enyaq Coupe; Average = mean of all **7** models; `<c:f>` = `Capri!$A$1:$E$9` |
| chart2 | **platform** pie (Online News, X, Forums, Blogs, Review, Tumblr), sheet `Capri-sources` |
| chart3 | **country** pie with 2-letter codes UK, IT, FR, DE, ES, NL, NO |
| chart4 | sentiment combo, **7** categories: Audi Q4 Sportback e-tron, Kia EV6, Peugeot E 3008, Polestar 2, Skoda Enyaq Coupe, VW ID.5, Ford Capri; diamond Y = 2.1, 6.45, 10.8, 15.05, 19.3, 23.53, 27.8 |
| line labels | Capri, Average and Enyaq Coupe have **no data labels**; three manual text boxes (`TextBox 25/28/14`) sit at the line ends. Value → y is linear (fit: 31 → 3,309,154 EMU, 130 → 3,042,321 EMU) **only if the value axis stays 0–600** — the builder pins `<c:max val="600"/>` and moves the boxes |
| pies | `<c:manualLayout>` on data labels is tuned to the old slices — strip it (labels auto-place) |
| bullets | `TextBox 16`, three `<a:p>`, 12pt, linked runs bold with `<a:hlinkClick>`; total ~850–1,000 characters fits; longer overflows the slide bottom |
| notes | one `notesSlide` per slide, bold heading paragraphs + body paragraphs |

The June deck's backing workbooks used sheet names from the Explorer template (`Explorer July`),
which did not match the chart formulas — the builder regenerates them to match (`Capri`,
`Capri-sources`).

## Sandbox caveat

LibreOffice in the remote sandbox could not load any file ("source file could not be loaded"), so
no PDF render was possible. QA was `validate.py --original`, `markitdown` text dump and an XML
re-read of every chart cache, relationship and label position. Do the visual pass in PowerPoint.
