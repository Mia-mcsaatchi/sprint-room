# Glossary

| Term | Meaning |
|------|---------|
| **Sprint Room** | The Brandwatch dashboard set used for Ford monthly reporting. Search "Ford Sprint Room" to find the dashboards. |
| **Carline** | A specific Ford model line. In scope: Capri, Explorer, Kuga, Puma. Out of scope: Focus, Mustang Mach-E, Puma Gen-E. |
| **Boolean exclusions** | Phrase-level exclusions added to a Brandwatch query to filter irrelevant content found during peak-checking (Phase 1). |
| **Mention List widget** | The widget at the bottom of each Brandwatch dashboard listing individual mentions; used for sentiment cleanup (Phase 3). |
| **Semiprocess** | The per-carline Google Sheets workbook holding raw data, helper formulas, and pivot tables. Key tabs: *Raw Data – Updated*, *Raw Data – Archive*. |
| **Process File tab** | A tab inside the semiprocess containing the 5 pivot tables summarising mentions drivers by platform and carline. |
| **Cluster** | A mentions driver with volume below ~30. When none qualify, use "no clusters of issues were identified as drivers". |
| **Passing mention** | A mention where the carline is referenced but isn't the article's focus (e.g. listed alongside several models). Always tagged neutral; notes-only, never a driver. |
| **Driver** | A genuinely carline-focused content cluster large enough to attribute (≥5 mentions for the slide; clusters are <30). |
| **Diamonds** | The prior-month sentiment markers overlaid on the current-month stacked bars (XY scatter). Sourced from the *published* prior deck. |
| **Backing xlsx** | The per-chart embedded workbook that makes a chart editable via "Edit Data in Excel". One per chart, never shared. |
| **£100 rule** | Don't hard-exclude a row unless you'd bet £100 it's noise. |
| **Carry-forward** | Prior-month data is left untouched and taken from the published prior-month deck, not re-derived. |
| **LENS** | The Ford social-media programme these reports feed into. |
| **EU5 / EU7** | Segment groupings. Kuga & Puma are EU5; Explorer & Capri are EU7. (The skeleton builder hardcodes EU7 — fix manually for EU5.) |
