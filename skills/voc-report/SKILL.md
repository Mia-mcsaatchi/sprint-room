---
name: voc-report
description: Convert the plain-text weekly EU Voice-of-Customer (VoC) automotive report into a formatted .docx that imports cleanly into Google Docs. Use whenever the user pastes a VoC / "HIGHLIGHTS" weekly report (EU5/EU7 vehicle mentions, ownership/pre-ownership phases, top discussed topics) and wants a Google Doc or formatted document made from it.
---

# Weekly VoC report → formatted Google Doc

The user runs a weekly EU Voice-of-Customer report. They paste raw plain text and
want it back as a formatted document they can open in Google Docs. This skill parses
the text and generates a styled `.docx`.

## How to run it

1. Save the pasted report text to a file (e.g. in the scratchpad):
   write it to `input.txt`.
2. Run:
   ```
   python3 ~/.claude/skills/voc-report/voc_report_to_docx.py input.txt output.docx
   ```
   The script prints a count of each element type it detected — sanity-check these
   against the report (e.g. one `title`, one `month`, N vehicles as `h1`, N `week`,
   the phase headers as `ph`, the "comments collected" lines as `collected`).
3. Copy the result to `~/Downloads/` with a descriptive name
   (e.g. `EU_VoC_Report_<period>.docx`) so the user can find it.
4. Tell the user to upload it to Google Drive and choose **Open with → Google Docs**;
   Google preserves the heading styles, bold, colours and nested bullets.

Requires only `python3` (stdlib) — no Node, pandoc or LibreOffice. There is no way
to render the .docx to an image on this machine, so verify structurally (the printed
element counts + XML validity) rather than visually.

## Formatting spec (matches the user's manual Google Doc)

| Element | Size | Style |
|---|---|---|
| `HIGHLIGHTS` (title) | 20pt | bold black, Heading 1 |
| Month/period line (line after HIGHLIGHTS, e.g. "During June") | 18pt | bold black, Heading 2 |
| Vehicle name (e.g. "Ford Explorer EV") | 16pt | bold black, Heading 1 |
| "During Last Week (w/c …)" | 14pt | bold black, Heading 2 |
| "… PHASE - TOP DISCUSSED TOPICS" | 12pt | bold black caps, Heading 3 |
| "… comments … were collected …" line | 12pt | **blue #3D85C6** |
| body text | 12pt | black |
| category (e.g. "Climate (3 mentions)") | 13pt | level-0 bullet (●), space before |
| issue description | 13pt | level-1 nested bullet (○), space after |

Body font is Arial throughout. Real Word heading styles are used so Google's document
outline picks them up.

## How the parser classifies each line

Blank lines are ignored; every non-empty line becomes one paragraph.

- **title** — first line, or any line equal to `HIGHLIGHTS`.
- **month** — the line immediately after `HIGHLIGHTS`.
- **vehicle (h1)** — any line whose *next* line starts with `During Last Week`.
- **week** — a line starting with `During Last Week`.
- **ph** — a line containing `TOP DISCUSSED TOPICS`.
- **collected** — the line immediately after a `During Last Week` line (rendered blue).
- **category (cat)** — matches `Topic Name (N mentions [by M owners])` with nothing
  after the closing paren, ≤6 words / ≤55 chars in the head, and no comma/colon/semicolon.
- **issue** — any other line while inside a bullet region (i.e. after the first category
  of the current phase).
- **body** — any other line before the first category of a phase (the
  "N owners generated…", "Among the positive mentions…", "Queries/issues noted…" lines,
  and the HIGHLIGHTS paragraphs).

The bullet region resets at every title / month / vehicle / week / phase header.

## Adjusting the styling

All sizes/colours live in `para()` and `STYLES` inside
`voc_report_to_docx.py` (sizes are half-points: pt × 2). If the user asks to tweak a
size, colour, bullet glyph or spacing, edit there and rerun. If the classifier
mislabels a line, refine `is_category` / `classify`.
