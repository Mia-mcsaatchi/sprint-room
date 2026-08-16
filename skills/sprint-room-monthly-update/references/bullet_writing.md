# Bullet Writing Reference

Sprint Room slides have three main bullets plus speaker notes. The style is consistent across cycles. Get it wrong and Val will push back.

## Bullet 1 — Volume (always the same shape)

Lead with the carline first, then list the competitors in volume-Δ order (largest movers first, both directions).

**Template:**
> Mentions volume {increased/decreased} for {all/most} competitors from {prior month}, with Ford {Carline} {Δ+/Δ-, %Δ}, alongside {Competitor 1} ({Δ}), {Competitor 2} ({Δ}), and {Competitor 3} ({Δ}). {Competitor 4} ({Δ}) and {Competitor 5} ({Δ}) saw {small increases/small decreases / etc.}.

**First-mention naming convention:**
- Carline: full "Ford <Carline>" (e.g. "Ford Explorer", "Ford Capri")
- Competitors: brand + model on first mention then just model thereafter ("Renault Megane E-Tech" → "Megane E-Tech")
- After first mention in the deck, drop the brand prefix

## Bullet 2 — Drivers

Quantify every named driver. Drop drivers with <5 mentions to slide notes.

**Template:**
> {Carline}'s {increase/decline} was driven by {primary platform} ({Δ}). Top driver: {named story} ({mention_count} mentions, {region/source}), {brief explanation}. Secondary driver: {named story} ({count}). {Other competitor}'s {Δ news count} mirrors/contrasts the pattern: {what's behind it}.

**Hyperlink rule:** every named driver phrase should be a hyperlink to a clean representative URL from that cluster:
- Prefer the cleanest domain (e.g. `finanznachrichten.de` over `24hamburg.de` for syndicated content)
- Pick news/blog over Twitter when both are available — articles are more authoritative reference points
- Hyperlinked text is bold (`b="1"` in the run); PowerPoint applies the theme color and underline

**Don't dismiss large clusters as "boolean noise" without reading the Full Text column.** A cluster of 35+ syndicated articles is almost never pure noise — it's usually a syndicated PR or wire story that genuinely involves the carline, even if the headline mentions a different model. Read the full text. The Stoll & Sauer "Kuga" PR in April 2026 was about Ford recall 25S79 which affects Explorer + Capri — pure Explorer content, just with a Kuga-led headline.

## Bullet 3 — Sentiment

Always cover the carline. Cover other competitors only if a meaningful shift (>2pp) occurred.

**Template (carline first):**
> {Carline} {sentiment direction summary, both positive and negative with pp deltas and absolute %s}. {No clusters of issues were identified as drivers / OR the issue cluster narrative}. {Other competitor name} {their notable shift if any, with caveat on base size if small}.

**Val's standard phrase**: when negative-driver volume <30 for the carline, use the exact phrase **"no clusters of issues were identified as drivers"**. Don't paraphrase. Val expects it word-for-word.

**Base-size caveat**: if a competitor has a big % shift on small absolute numbers, flag it: "+2.2pp negative, on a small base of 8 mentions". Otherwise the reader thinks the world is ending when really it's just noise.

## Speaker notes

Notes carry the country breakdown, ID.4 (or biggest competitor) driver detail, and closing data-quality / context commentary.

**Country breakdown**: provide if any country has >30 mentions in the carline data. Show as "Country: {count} ({Δ})". Two blocks: total country breakdown, and news-only country breakdown.

**Competitor driver detail**: pick the biggest competitor mover. 3 bullets each with quantification and a representative URL.

**Closing commentary**: cohort-wide sentiment context (if relevant), data-quality flags, recommendations for next cycle's boolean cleanup.

## Tone

- Past tense for what happened (decreased, rose, was driven by)
- Present tense for ongoing characteristics (X is the largest competitor, Y dominates news)
- Numbers in plain text, not formatted with commas for small numbers
- pp = percentage points; %Δ = percentage change. Don't conflate them.
- "X pp" not "X%" when reporting a sentiment shift (because the % is part of what shifted)

## Examples — good vs bad

**Bad:** "ID.4 Twitter mentions grew +118 in its place."  
*Why bad: no context for why, meaningless juxtaposition.*

**Good:** "ID.4's -349 news decline mirrors the pattern as March's VW/Cupra battery recall faded; April's largest story was VW halting ID.4 US production (39 mentions, German news)."

**Bad:** "Peugeot E-3008 saw a notable +8.2pp positive swing (3.9% → 12.1%) driven by syndicated Spanish coverage of the 3008 Plug-in Hybrid's 100km EV range (18 positive mentions)."  
*Why bad: was based on stale sentiment tags; "what's positive about this?" — the headline is informational, not laudatory. The "swing" disappeared on retag.*

**Good (the fix):** Don't write the swing bullet if the data doesn't support it. Move to notes as informational: "Peugeot E-3008 sentiment was largely flat MoM."

## When data has been retagged

If the analyst has retagged sentiment between Brandwatch and the chart export, the chart export's numbers and the raw-data-derived numbers may disagree. Convention:
- **Chart bars and diamonds** use the chart export's percentages (this is what the client sees on screen)
- **Bullet narratives** use the raw data's percentages (this is what's actually true after retagging)
- **Prior-month diamonds**: always come from last month's *published* chart, never from the new raw data — last month's slide is "frozen" as what the client received

If chart export and raw data disagree by >0.5pp on any sentiment value, flag it to the analyst and ask which to use.
