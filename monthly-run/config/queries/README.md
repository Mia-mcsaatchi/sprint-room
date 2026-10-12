# Brandwatch queries

Versioned copies of the Brandwatch boolean queries, so changes can be reviewed and rolled back.
Brandwatch itself has no history: **back up the live query here before you edit it.**

| File | Status | What it is |
|---|---|---|
| `Explorer_v2_2026-09.txt` | **Live** since September 2026 | Refined Explorer query (paste the whole file into the query editor). |
| `Explorer_v2_addon_2026-10.txt` | Proposed | Extra exclusions for noise that v2 let back in (yachts, mining "explorers", Windows Explorer, Cagiva/Eco Rider products, NASA). Paste inside v2's final `NOT ( … )` block, before its closing bracket. Tested against the Aug–Sep 2026 export: removes 23 of 40 remaining noise mentions, 0 genuine. |

## What changed in Explorer v2 (vs the pre-September query)

- **Inclusion:** added owner/spec terms to the electric-word lists — `kWh`, `Akku`, `RWD`, `AWD`,
  `MEB`, `batterie`/`batteria`/`batería`, Dutch `bereik`, Norwegian `elbil`/`elbilen`/`rekkevidde`;
  "2026" counts as "new".
- **Exclusions removed (they deleted genuine posts):** single everyday words such as German `war`,
  `Tod`, `Gericht`, `Angriff`; French `occasion`, `vole`, `saisi`, `tribunal`, `avocat`; Spanish
  `a la venta`, `renting`, `ocasión`, `agentes`, `coma`; Italian `incidente`, `agenti`, `caccia`,
  `Voler`; English `war`, `dead`, `died`, `killed`, `stocks`, `injury`, `attorney`, `stromer`;
  `site:whatcar.com`, `site:parkers.co.uk`; the "Ford recurre al gigante chino BYD…" and Goodwood
  titles. Crime/accident/used-car **phrases** were kept.
- **Narrowed:** `Explorer NEAR/10 (years, ICE, hybrid…)` → `NEAR/3` without 2021/2022/ST/Platinum;
  `bike`, `tiger`, `Microsoft`, `yacht` only count near "explorer".
- **Added noise exclusions** (guarded with `AND NOT ford`): Xiaomi SkyNomad "Explorer Edition",
  Fermi Explorer, Anthropic "Econ Scenario Explorer", mining/stock "explorers", ships/aircraft,
  scooters and gadgets; French non-Ford mentions (`language:fr AND NOT (ford OR …)`).

Effect on the Aug–Sep 2026 Explorer pull: September noise fell from 86 mentions (41%) to 18 (9%),
and about 50 genuine mentions a month were recovered (mostly German owner/forum posts). Brandwatch
back-filled earlier months on the new basis.

## Competitor queries

The competitor queries share the same generic exclusion block and were **not** changed in
September 2026. Apply the same exclusion fixes to all six queries in one month so share of voice
stays comparable, and back the old versions up here first.
