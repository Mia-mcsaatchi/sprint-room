# Skills

Verbatim, version-controlled copies of the automation skills. This is the source of truth —
when a skill changes in the working Claude environment, commit the change back here so they
don't drift.

## Included

### `sprint-room-monthly-update/` (core)
Rolls a carline deck forward one month: rebuilds the four charts, embeds editable backing xlsx,
writes evidenced bullets with inline hyperlinks, and drops the oldest comparison slide.
- `SKILL.md` — the 7-stage workflow (Stage 0 scope → Stage 7 pack/QA).
- `scripts/` — `validate_inputs.py`, `extract_drivers.py`, `build_skeleton.py`,
  `update_{volume,country,sentiment,platform}_chart.py`, `write_bullets.py`,
  `embed_backing_xlsx.py`, `pack_and_qa.py`.
- `references/` — `pptx_structure.md`, `bullet_writing.md`, `embed_xlsx_pitfalls.md`.

### `voc-report/` (related)
Converts the weekly EU Voice-of-Customer text report into a clean `.docx` that imports well into
Google Docs.
- `SKILL.md`, `voc_report_to_docx.py`.

## Not included (separate project)

### `amazon-influencer:influencer-slicer-qa`
QA / validate / repair the monthly Amazon Influencer Slicer (`Influencer_Slicer_YYYY_MM.xlsm`).
This is a **different client/workstream** from Ford Sprint Room, so it's deliberately left out of
this repo to keep the handover clean. Options:
- give it its own repo (recommended), or
- add it here under a top-level `amazon-influencer/` tree if you'd rather keep one repo.

Its files (for reference): `SKILL.md`, `scripts/` (`qa_battery.py`, `delivery_gate.py`,
`qa_publish.py`, `rebuild_xlsm.py`, `splice_publish.py`), `references/` (`qa-checks.md`,
`imputation.md`, `known-issues.md`, `system-map.md`, `tracker-investigation.md`,
`delivery-qa.md`, `rebuild-xlsm.md`).
