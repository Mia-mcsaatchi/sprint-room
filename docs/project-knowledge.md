# Project Knowledge

## What the programme is

A monthly social-media competitive report ("Sprint Room") for four Ford carlines, built from
Brandwatch data. Each carline gets a `.pptx` with three slides: **title → new month → previous
month** (the previous-month slide is carried forward untouched for client comparison).

## People

| Who | Role |
|-----|------|
| **Mia** | Social media analyst (McCann). Builds the decks; sign-off on analytical decisions within a build. |
| **Val** | Senior reviewer. Owner of the house-style rules. Hard-gate reviewer before delivery. |
| **Mike, Seb** | Downstream recipients (CC on delivery). |

Delivery goes **to Val, CC Mike / Seb / Mia**.

## Carlines & competitor sets

| Carline | Segment | Competitor set | # competitors |
|---------|---------|----------------|---------------|
| Kuga | EU5 | Dacia Bigster, Jaecoo 7, Hyundai Tucson, Peugeot 3008, VW Tiguan | 5 |
| Puma | EU5 | Peugeot 2008, Toyota Yaris Cross, Hyundai Kona, VW T-Cross | 4 |
| Explorer | EU7 | Skoda Enyaq, VW ID.4, Renault Megane E-Tech, Peugeot E-3008, Volvo XC40 Recharge/EX40 | 5 |
| Capri | EU7 | Peugeot E-3008, VW ID.5, Skoda Enyaq Coupe, Polestar 2, Audi Q4 Sportback e-tron, Kia EV6 | 6 |

Out of scope entirely: **Focus, Mustang Mach-E, Puma Gen-E**.

> The competitor list is a **per-carline variable**, not a constant. Read the real names from
> the deck's `chart3.xml` `<c:cat>` strings each cycle rather than assuming.

## Reporting run order

Brandwatch dashboard work is done in two passes:
1. **Explorer + Capri**
2. **Puma + Kuga**

## Known state / conventions confirmed in past builds

- **Brandwatch retains Phase 3 sentiment retags across export cycles** — confirmed on the
  Explorer June build; re-pulling shows negligible drift. The BW-cleaned export is the
  authoritative sentiment source when available.
- **Carry-forward rule:** in a new month's build, the prior month's data is left entirely
  untouched and carried forward from the *published* prior-month deck — never re-derived.
- **Open watch-item:** the Capri/Explorer towing PR query asymmetry (Capri's `NEAR/5` operator
  catching fewer articles) may need revisiting in a future cycle.

## Working style (for whoever inherits reviewing with Mia)

- Directive and concise; approves with shorthand.
- Wants judgement calls **flagged explicitly**, not silently auto-corrected.
- After bullet-menu approval, a single "yes" carries through the remaining stages — no further
  checkpoints expected.
- Exercises close editorial judgement on analytical claims and will push back / request
  re-analysis when framing doesn't hold up.
