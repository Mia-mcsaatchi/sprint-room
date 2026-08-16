# Workflow — The Six Phases

Condensed reference. The authoritative, screenshot-annotated version is
[`../process/Sprint_Room_Process_Reformatted.docx`](../process/Sprint_Room_Process_Reformatted.docx).
The end-to-end automation design is
[`../process/Sprint_Room_Automation_Architecture_v2.pdf`](../process/Sprint_Room_Automation_Architecture_v2.pdf).

Phases 1 and 2 can run in parallel; Phases 3–6 are sequential.

| Phase | What you do | Where | Output |
|-------|-------------|-------|--------|
| 1 | Clean Brandwatch queries — check peaks & exclude noise | Brandwatch (all 4 carlines) | Cleaner queries before reporting starts |
| 2 | Set up the new month's deck | PowerPoint (from Trello card) | Title slide updated, slide 2 ready |
| 3 | Clean dashboard sentiment & update charts | Brandwatch → PPT | Volume/source/sentiment charts refreshed |
| 4 | Export raw data & update the semiprocess sheet | Brandwatch → Google Sheets | Pivots ready for both months |
| 5 | Write insights from charts & pivots | PPT slide + notes | Bulleted insight on the new slide |
| 6 | Final QA & deliver | PPT, Sheets, Trello, email | Sent to Val (CC Mike, Seb, Mia) |

## Phase 1 — Clean the queries (analyst judgement)

For every carline: set the date range to the reporting month, click into **each peak** in the
Mention Volume chart (ideally every date), note key phrases from irrelevant items in their
original language, then add them to the boolean exclusions. **Batch the exclusions** — don't
update per peak (each change forces a Brandwatch refresh). Re-run to catch anything missed.

## Phase 2 — Set up the deck

Download last month's deck from the Trello card. Update the title slide (new month + today's
date). Duplicate the first non-title slide → new slide 2; update its subtitle, delete its notes,
and move the old insights down + highlight yellow as reference. Mark the former slide 2 as the
previous month's slide (kept for client reference). Delete the old slide 4.

## Phase 3 — Clean sentiment & update charts

In the Mention List widget, filter to the carline + Online News + positive/negative. Re-tag any
non-carline-focused item to **neutral** (passing mentions, multi-carline articles, other-model
articles). Fix clearly wrong pos/neg tags. Delete blatantly irrelevant items. **Every retag must
also be made in Brandwatch, not just the chart.**

Chart date scopes: volume = reporting month + previous 3; source & sentiment = reporting month
only. **Save the dashboard** before exporting. Then update each chart:
- Volume: shift the newer three months left, leave room for the new month, **don't overwrite the
  Average row** (formulae).
- Source pies: can be read straight from Brandwatch (hover segments / "Breakdown by").
- Sentiment: fiddliest — left tables editable, right tables formulated (don't touch). Move each
  block down one; paste-as-values red→bottom, green→red; paste fresh export into the top block;
  update the month label. **Verify model order** — BW exports don't always match the table order.

## Phase 4 — Export raw data & update the semiprocess

Export each competitor's data (including Ford) via Data Download → **Excel** (not CSV). Paste
into the *Raw Data – Updated* tab, column D, bottom-most free space (hold Ctrl+V ~10s). Extend
helper formulas in columns A (`=MONTH(F)`) and B (`=GOOGLETRANSLATE(G)`) with Ctrl+D. Archive the
oldest month to *Raw Data – Archive* and **delete its rows** (don't clear — blanks break pivots).
Update the 5 pivot tables per tab to cover both months.

## Phase 5 — Write the insights

Follow [`editorial-standards.md`](editorial-standards.md). Bullets 1–3 on the slide; overflow
(extra drivers, country breakdowns) in the notes.

## Phase 6 — QA & deliver

QA data consistency across charts / Brandwatch / semiprocess. Check subtitle month + title-slide
date. Tidy the semiprocess (both months shown, filters cleared, analysis columns removed, col C
re-hidden, share = anyone-with-link-can-edit). Save & upload the PPT to the new Trello card, put
the semiprocess URL in the card, email **Val** CC **Mike, Seb, Mia**.

## What the skill automates vs what stays manual

The `sprint-room-monthly-update` skill covers roughly Phases 2–5 mechanics (skeleton, charts,
backing xlsx, bullet scaffolding). **Manual / analyst-only:** Phase 1 boolean cleaning, Phase 3
sentiment retagging, Phase 4 Google-Sheets pivots, and Phase 6 Trello upload + email.
