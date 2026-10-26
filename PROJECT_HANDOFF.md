# AVESTA Excel Model — Project Handoff

Last reviewed: 26 October 2026 (Africa/Cairo).

## Experimental solar-production copy - 26 October 2026

The user subsequently authorized a separate experimental workbook, while retaining the original authoritative workbook and deferring adoption. See `SOLAR_PRODUCTION_SPEC.md` for the trial's equations, sources and decisions.

- Trial: `outputs/solar-trial/AVESTA_Cairo_Seasonal_Solar_Trial.xlsx`.
- Adds a sixth worksheet, `Solar Production`, with fixed Cairo coordinates and editable tilt and compass azimuth. This extra worksheet is explicitly authorized for the trial.
- Uses an embedded PVGIS 5.3 / SARAH3 monthly-yield grid (475 orientations, 2005-2023 climate, retrieved 26 October 2026), with native Excel bilinear interpolation. Default: 30-degree tilt, south, crystalline silicon, free-standing, 14% PVGIS system-loss input.
- Links seasonal production into the existing solar-only and combined monthly models; displays monthly and annual production and bill savings, with annual net savings separately. A study-year selector permits reviewing later years.
- The original workbook, its original five-sheet structure, and its fixed-yield methodology remain unchanged. The descriptions below continue to describe that authoritative original unless explicitly identified as trial behavior.
- The user subsequently authorized committing and pushing this trial and its documentation to `excel-sheet-branch`. Adoption remains deferred; the original workbook is retained. Saved scenario values in the trial remain historical captures of the older method.
- **Next requested step:** remind the user that battery benefit should be evaluated with hourly PV/load/state-of-charge calculations, not monthly energy budgets. This hourly dispatch is not implemented by the solar trial.

Local development and verification evidence: `build/solar-trial/`. These supporting files are not workbook runtime dependencies.
Eleven recalculated trial cases and independent energy/financial checks passed. Three direct PVGIS spot checks found annual interpolation differences below 0.18% and monthly differences up to 1.28% at those tested angles; these are not universal error bounds. The XLSX passed Open XML SDK validation with zero errors. Microsoft Excel UI behavior remains untested.

## Start here

This document is the continuation reference for the **Excel solar and battery scenario model**, not the separate web application in this repository. The repository README primarily describes that web application.

The current workbook has five worksheets, a live model for three system types, eight saved manager-design scenarios, and live diagnostic checks. It uses native Excel formulas and requires no Python, VBA, macros, or external scripts for ordinary use. The user manually copies scenario settings and results as values to retain comparisons.

**Communicate with the user in Egyptian Arabic. Keep workbook content, technical identifiers, and this document in English.**

Before continuing, inspect the actual workbook and Git state. This document is a dated reference, not a substitute for checking files. The user also works on another computer; edits visible in screenshots are not necessarily present in the repository workbook. Preserve any newer user-edited workbook rather than replacing it with an older repository copy.

## 1. Authoritative files and repository

- Repository: <https://github.com/boodyelkady1998-sudo/avesta-bess-sizing-tool>
- Excel branch: `excel-sheet-branch`.
- Current workbook: [`outputs/scenario-comparison/AVESTA_El_Mo3alem_Eight_Scenarios_Excel_Compatible.xlsx`](outputs/scenario-comparison/AVESTA_El_Mo3alem_Eight_Scenarios_Excel_Compatible.xlsx).
- GitHub workbook: <https://github.com/boodyelkady1998-sudo/avesta-bess-sizing-tool/blob/excel-sheet-branch/outputs/scenario-comparison/AVESTA_El_Mo3alem_Eight_Scenarios_Excel_Compatible.xlsx>
- Latest reviewed workbook commit: `ae7fa4b06c13fec080fe88b93de09a3177582959` — `Add blank future changes and saved scenarios sheets`.
- That commit was created on 25 October at 18:06:31 Cairo time and pushed on 26 October at approximately 09:13:59 Cairo time, according to the local remote-tracking reflog. Commit time and push time are different.
- Local workspace at handoff: `C:/Users/moham/OneDrive/Desktop/Projects/Sizing_tool`.

There is one current workbook in `outputs/scenario-comparison`. Its historical filename still mentions eight scenarios; this does not limit the number of scenarios the user may save manually. Older files remain in `outputs/excel-v1`; do not mistake them for the authoritative version or delete them without checking scope.

Update the existing authoritative workbook rather than publishing successive named copies. Do not upload or push future edits without a user request. This handoff file was created locally; its creation alone does not imply it was committed or uploaded.

## 2. Product decisions

Implemented or explicitly agreed:

- Use an ordinary `.xlsx` workbook for now. No runtime dependency on Python and no save-scenario macro.
- Location is fixed to Cairo. Use one average daily PV yield, not a monthly Cairo production reference table.
- Editable assumptions and equipment catalog are separate from case inputs.
- Three live configurations: `Solar only`, `Battery only`, and `Solar and battery`.
- System role is an output inferred from savings and backup capability, not an input choosing a financial objective.
- A new use case can be evaluated by changing inputs and equipment quantities. Scenario capture is manual and values-only.
- Unknown roof area means no roof constraint is imposed; roof feasibility remains unverified. It does not mean unlimited space has been confirmed.
- Unknown load distribution uses constant load across hours. With 12 daytime hours, this gives a 50% daytime energy share.
- Unknown critical load does not prevent energy calculations. Backup energy/power can be reported, but backup duration for the unknown load cannot.
- Outage-loss valuation, automatic sizing optimization, and new dispatch algorithms remain future work.

The original three-worksheet constraint was explicitly superseded when the user requested two additional blank worksheets. The current required structure is **five worksheets**.

## 3. Worksheet and cell map

### Inputs

| Cell | Purpose |
|---|---|
| C5 | Case name |
| C6 | `Bill only`, `Consumption only`, or `Bill and consumption` |
| C7 | Monthly consumption, kWh; not a cumulative meter reading |
| C8 | Monthly bill, EGP |
| C9 | `Residential` or `Commercial` tariff category |
| C10 | Billing period days |
| C11 | `Official` or `Actual flat rate` pricing |
| C14:C15 | `Specified` / `Unspecified` roof area, then available m2 |
| C17:C18 | `Known` / `Constant load`, then daytime share if known |
| C20:C21 | `Known` / `Unknown` critical load, then kW if known |
| L3:L14 | Derived input status, consumption, pricing, and bill comparison |

For bill-only input, the model inverts the selected tariff. A bill that does not identify a unique consumption is flagged. Actual flat-rate pricing requires both bill and positive consumption; otherwise the model uses the official basis and the checks warn about the fallback.

### Assumptions and Catalog

| Cells | Purpose |
|---|---|
| C6:C26 | Financial, technical, calendar, and replacement assumptions |
| C31 | Average daily Cairo PV yield, kWh/kWp/day |
| Rows 50:57 | Residential tariff calculation data |
| Rows 60:65 | Commercial tariff calculation data |
| Rows 73:74 | Panel catalog: ID, description, Wp, area, price |
| Rows 79:81 | Battery catalog: nominal kWh, usable fraction, charge/discharge kW, price, compatibility group |
| Rows 86:89 | Inverter catalog: AC kW, PV limit kWp, charge kW, backup flag, price, compatibility group, phases |
| Lower existing documentation section, through row 168 | Formula explanations and model documentation |

Catalog lookup ranges are currently fixed. Merely adding a catalog row does **not** guarantee that dropdowns, formulas, and checks will include it. Extend all affected references and validate them when expanding the catalog.

### Scenarios

| Area | Purpose |
|---|---|
| C:E | Live results: solar, battery, combined respectively |
| Rows 6:26 | Configuration status and principal results |
| H7:L18 | Equipment controls; J = solar, K = battery, L = combined |
| J9:L9 | Panel model |
| J10:L10 | Panel quantity |
| J11:L11 | Battery model |
| J12:L12 | Battery quantity |
| J13:L13 | Inverter model |
| J14:L14 | Reserved share of usable battery capacity |
| K15:K16 | Battery-only mode (`Backup` / `Grid cycle`) and daily grid-charged delivery limit |
| H20:L23 | Live error/warning counts and `Open checks` links |
| Rows 50:70 | Sizing, costs, compatibility, and original validity logic |
| Rows 78:98 | Annual cash-flow calculations |
| Rows 161:400, 411:650, 661:900 | Monthly energy and bill calculations for the three types |
| W:AF, through row 121 | Eight original saved scenario columns; W labels, X units, Y:AF cases 1–8 |
| AH:AL | Shared assumptions, sources, and limitations for those captures |
| A920:N980 | Live checks; C:F solar, G:J battery, K:N combined |
| Rows 926:980 | 55 checks for each system type |
| A988:N1032 | Formula-error diagnostics, first failing cell per scanned column segment |

The formula scan covers defined existing model ranges, not every possible future cell. Extend it if calculation areas are expanded. The checks are diagnostic outputs; they do not replace the original `C65:E65` calculation gates. A new check can flag an issue while older result formulas still display numbers. Do not use numeric appearance as evidence of validity.

Error counts count failed checks, not necessarily distinct root causes: one overload can fail both the specific limit check and original-model-validity check. Formula-error checks currently report workbook-wide failures and can warn all three types when a failure occurs in one type. Some inactive equipment fields still need valid catalog IDs because legacy helper formulas reference them.

### Future Changes

Completely blank and unprotected at handoff. Intended for the user's future change requests. No template, headings, or backlog was inserted into it.

### Saved Scenarios

Completely blank and unprotected at handoff. Intended for manually copied scenario settings/results. No capture button, automatic append behavior, or fixed scenario limit was implemented. The eight original captures remain in `Scenarios`; they were not moved here.

## 4. Current defaults and model boundaries

The repository workbook currently contains the original example inputs: 2,421 kWh, EGP 8,279, 30 billing days, commercial tariff, **Official** pricing, unspecified roof area, constant load, and unknown critical load. A 30-day representative billing period is a planning assumption, not proof of the original bill's exact duration.

| Assumption | Cell | Current value |
|---|---|---:|
| Study period | C6 | 10 years |
| Nominal discount rate | C7 | 10% |
| Tariff / maintenance escalation | C8 / C9 | 0% / 0% |
| Annual PV degradation | C10 | 0.5% |
| Charge / discharge efficiency | C11 / C12 | 97.5% / 97.5% |
| Daytime duration | C13 | 12 hours |
| PV / grid charging window | C14 / C15 | 6 / 8 hours |
| Panel-area installation multiplier | C16 | 1.5 |
| Installation and BOS | C17 | 10% of equipment cost |
| Annual maintenance | C18 | 1% of total CAPEX |
| Battery replacement year / cost fraction | C19 / C20 | 11 / 100% |
| Inverter replacement year / cost fraction | C21 / C22 | 12 / 100% |
| Annual battery degradation | C23 | 2% |
| First forecast year | C24 | 2027 |
| Inverse-bill tolerance | C25 | EGP 0.02 |
| Battery auxiliary consumption | C26 | 0 kWh/day |
| Average daily PV yield | C31 | 4.98 kWh/kWp/day |

All cells in this table are on `Assumptions and Catalog`. Replacement events are outside the default ten-year study. These are model assumptions, not supplier warranties, verified current quotations, or universal engineering values.

Core logic:

- Installed PV kWp = panel count × panel Wp / 1,000. There is no direct kWp override input.
- Monthly PV generation = kWp × daily yield × calendar days × annual degradation factor. The yield already incorporates the previously assumed PV losses; do not apply a second generic performance ratio without revising the model basis.
- Direct solar use is limited by available PV and estimated daytime energy demand. Excess PV has no export revenue in the model.
- Solar-only cannot shift surplus daytime generation to nighttime consumption. A large annual generation total does not imply full bill elimination.
- Battery dispatch uses simplified energy budgets, efficiency, reserve, charge/discharge limits, and charging-window assumptions. It is not an hourly time-series dispatch model.
- Grid cycling at a single flat electricity rate does not create tariff-arbitrage savings; losses can increase consumption.
- Bill savings = baseline bill minus post-installation bill. Net savings subtract maintenance and applicable replacement costs.
- NPV = initial negative CAPEX plus discounted annual net cash flows. Negative NPV is a financial result, not a formula error.
- Simple and discounted payback are separate. A positive simple payback within the study does not guarantee positive NPV.
- Full-charge backup energy/duration is conditional on that charge actually being available. Critical load can exceed the inverter/battery power limit even when nominal kWh seems adequate.

Not modeled or not established: outage financial benefit, export payments, verified time-of-use/peak-demand charges, financing, taxes, motor-starting transients, detailed phase balancing, hourly load/PV coincidence, and monthly solar seasonality. Current tariffs and equipment data require verification before commercial decisions. Historical source claims about Egyptian regulations are not automatically verified facts.

## 5. Manager scenarios and El Mo3alem source work

Primary shop document: `AVESTA Egypt El Mo3alem Shop BESS Report.docx` (tracked in the repository). Confirmed reference quantities are 2,421 kWh and EGP 8,279.

Manager document: `PV + Battery Storage System for a Juice Shop in Egypt V1 -Mohamed Ismail.docx`, obtained from the `equations-presentation` branch. The annotated source used for analysis is available locally as `build/scenario-audit/manager-source.docx`; extracted text is in `manager-source.txt`. The similarly named root document may be a different revision. Use the annotated source and verify branch/path if recovering files on another computer.

The manager's estimated 109.53 kWh/day / 3,286 kWh/month is a separate appliance-based estimate; do not substitute it for confirmed billed consumption without explaining the change. The 4 kW critical-load figure is an estimate, not a measured simultaneous peak.

Explicit user decisions for the saved eight-case exercise:

1. Use the manager's total capital costs, disclosing component-total inconsistencies. Use the workbook's remaining assumptions.
2. For scenarios 6–8, retain the supported solar/backup results but mark full operating-strategy evaluation **incomplete** until tariff and operational data exist. Changing equipment size is not implementing arbitrage or peak shaving.
3. Save each case sequentially as static values, preserve all earlier captures, independently check calculations, then restore the original live inputs/catalog.

| # | Type | Panels | PV kWp | Battery kWh | Inverter kW used in test | Manager CAPEX, EGP |
|---:|---|---|---:|---:|---:|---:|
| 1 | Solar only | 46 × 550 W | 25.30 | 0 | 25, assumed | 413,500 |
| 2 | Battery only | None | 0 | 10 | 3 | 172,500 |
| 3 | Solar and battery | 3 × 670 W | 2.01 | 10 | 3 | 197,300 |
| 4 | Solar and battery | 7 × 670 W | 4.69 | 10 | 5 | 247,000 |
| 5 | Solar and battery | 15 × 670 W | 10.05 | 20 | 10 | 467,500 |
| 6 | Solar and battery; full strategy incomplete | 3 × 670 W | 2.01 | 30 | 5 | 454,800 |
| 7 | Solar and battery; full strategy incomplete | 15 × 670 W | 10.05 | 50 | 10 | 812,850 |
| 8 | Solar and battery; full strategy incomplete | 3 × 670 W | 2.01 | 50 | 5 | 662,500 |

Important capture qualifications:

- Scenario 1 is 25.3 kWp, despite the source's rounded “25 kW” label.
- Scenarios 3–8 exceed their stated roof allowances under the unchanged 1.5 area multiplier: 12.15 vs 10 m2, 28.35 vs 25 m2, or 60.75 vs 50 m2. Their conditional financial diagnostics are not approved feasible designs.
- A 3 kW inverter in scenarios 2–3 cannot supply the entire estimated 4 kW critical load.
- Scenario 4 component midpoints total EGP 247,830 versus declared EGP 247,000; scenario 5 components total EGP 467,850 versus declared EGP 467,500. Declared totals were retained and differences disclosed.
- Temporary test catalog overrides supplied missing inverter ratings and manager costs. They were restored afterward. The current catalog does not reproduce all eight designs without deliberate catalog/cost changes.
- In particular, current `GRID10` has a 13 kWp PV limit. Entering 46 `PV550` panels with that inverter correctly raises an overload check.
- Saved values do not update when inputs or assumptions change. Rerun a case to refresh it; do not silently overwrite its historical result.

## 6. Later questions: NPV and maintenance

The user's later screenshot represented a separate manual test, not the original manager-cost capture. It showed solar CAPEX of EGP 354,640, 25.3 kWp PV with a 20 kW inverter, and **Actual flat rate** pricing. The battery test showed EGP 183,040 CAPEX, 10 kWh, and a 5 kW inverter. Those differ from the original source-case costs and inverter assumptions.

The solar screenshot's NPV of **−EGP 66,721.50** was independently reproduced using the displayed costs, 2,421 kWh / 30 days, a 50% daytime energy share, EGP 8,279 / 2,421 kWh flat pricing, 1% maintenance, calendar years 2027–2036, and 10% discounting. Year-one net savings were EGP 46,817.52. Of 45,987.81 kWh generated, the simplified model directly used 14,727.75 kWh and left 31,260.06 kWh without financial credit.

The manager's scenario 1 instead assumes EGP 132,000 first-year savings, near-total load coverage, and increasing savings in later years. Treating production as full financial savings conflicts with nighttime load and the source's own no-export premise. Its printed discounted-flow column also contains a summation discrepancy: EGP 1,005,132 less EGP 413,500 equals **+EGP 591,632**, not the printed +EGP 691,632. Recomputing discounted flows from the rounded cash-flow column gives approximately +EGP 591,655; the small difference is rounding, unlike the EGP 100,000 discrepancy.

Battery-only backup can have zero bill savings and negative NPV because initial cost and maintenance still exist. In that screenshot, annual maintenance was 1% × EGP 183,040 = EGP 1,830.40. Parentheses denote a negative value; a dash in the existing number format can mean zero.

**Maintenance remains a model limitation:** the same 1% of total system CAPEX is applied to all three types. It is not a verified annual service cost for a lithium battery. Separating PV cleaning/service from battery-system/inverter inspection costs was discussed but **not implemented**. Replacement cost remains a distinct assumption. Relevant manufacturer references discussed were BYD battery cleaning guidance and periodic connection/cabinet inspections; they do not justify a universal 1% fee:

- <https://site.bydbatterybox.com/uploads/downloads/250219_BYD%20Battery-Box%20Premium%20HVM-US_Operating%20Manual_V1.0_EN-67b59e88b96db.pdf>
- <https://www.bydbatterybox.com/uploads/downloads/230810_BYD_Battery-Box_Premium_HVS%26HVM_Combiner_Box_Installation_Manual_V1.3-64d5a84ca9198.pdf>

Renaming `Year 1 savings after maintenance` to `Year 1 net savings / (cost)` was suggested but **not implemented**. Do not describe it as already changed.

## 7. Verification and Excel compatibility history

Completed checks:

- Eight manager-design cases were calculated sequentially, checked against independent energy/bill/cash-flow calculations, and saved as values.
- Six earlier boundary probes covered zero PV, zero battery, night-only load, grid cycling at a flat rate, insufficient roof area, and excessive critical load.
- Live diagnostics were tested in 18 recalculated cases: baseline, PV overload, insufficient roof, missing bill, fractional panel quantity, negative battery quantity, unknown model, incompatibility, critical-load overload, invalid efficiency, invalid reserve, invalid battery operating mode, invalid daytime share, simultaneous faults, an injected formula error, unknown critical load, blank inactive optional fields, and unknown tariff category.
- Default diagnostics: zero errors in all three types; warnings of 1 / 1 / 2 for solar / battery / combined, reflecting unknown roof area and/or critical load.
- Existing original cells, formulas, inputs, and eight captures were checked for preservation. Adding the two blank tabs left the original worksheet XML unchanged.

Compatibility incidents and repairs:

1. ONLYOFFICE opened a previous export, but Microsoft Excel requested recovery and then showed empty sheets.
2. Microsoft's Open XML SDK validator identified 30 schema errors: three misordered `sheetProtection` elements and 27 undeclared revision UID attributes. Protection nodes were reordered and irrelevant revision UID metadata removed. Cell data/formulas were preserved.
3. Later, Excel refused adding a worksheet with “This action won't work on multiple selections.” Investigation found a missing workbook-level `bookViews` element. It was added with a valid active workbook view when the two requested blank tabs were added. This is a plausible compatibility repair, **not a native-Excel-confirmed diagnosis**.
4. The final five-tab workbook passed Open XML SDK validation with zero errors. Windows .NET package parsing also required standard default namespaces for content types and package relationships; preserve those when editing ZIP/XML parts.

Native Microsoft Excel was not available to the agent for direct UI testing. LibreOffice recalculation, artifact-tool inspection/rendering, read-only openpyxl inspection, independent calculations, and Open XML SDK validation were used. A successful schema check does not prove every native Excel interaction works. Do not claim otherwise.

The user was advised to save a new `.xlsx` copy from Excel, close, and reopen as a workaround for selection-state issues. The newly added tabs have not been confirmed by the user on their other computer at this handoff.

## 8. Local implementation aids and cautions

These are development aids, **not workbook runtime dependencies**. Many are untracked local files and will not exist in a fresh clone unless separately supplied:

| Path | Purpose |
|---|---|
| `build/excel-v1/scenario-audit.mjs` | Artifact-tool scenario execution and comparison-region authoring |
| `build/scenario-audit/cases.json` | Source-case settings and temporary catalog overrides |
| `build/scenario-audit/verify-case.py` | Independent scenario checks with LibreOffice recalculation |
| `build/scenario-audit/result-*.json` | Recorded per-case evidence |
| `build/scenario-audit/verify-edges.py` | Earlier boundary checks |
| `build/excel-v1/add-checks.mjs` | Live diagnostic formulas and authoring |
| `build/checks/assemble.py` | Selective native merge preserving original workbook features |
| `build/checks/test.py`, `test-results.json` | 18-case diagnostic verification and results |
| `build/checks/finalize.py`, `format-summary.py` | Additional conditional formatting and summary formatting |
| `build/checks/schema-five-tabs.json` | Final validation output; zero-error runs may leave an empty report |
| `build/compatibility/validate.ps1` | Open XML SDK validation |
| `build/compatibility/sdk/` | Downloaded Microsoft Open XML SDK 2.20.0 libraries |
| `build/excel-v1/blank-tabs.mjs`, `build/checks/add-blank-tabs.py` | Creation and grafting of the two blank tabs |

**Do not blindly rerun these scripts.** They were developed incrementally: some refer to obsolete source/output filenames, some expect exactly three source tabs, and some depend on a later XML normalization or style-order fix. The committed workbook is authoritative. The old comparison assembler also predates the compatibility repairs. Reconstruct the intended minimal change against the current file, and validate the final exported package.

Use the currently available spreadsheet skill and bundled runtime paths, discovered through workspace dependencies. During prior work, authoring used `@oai/artifact-tool`, native ZIP/XML handling preserved features the export did not safely retain, and openpyxl was used read-only. The bundled runtimes were under `C:/Users/moham/.cache/codex-runtimes/codex-primary-runtime/dependencies/`. Do not assume that absolute path on another machine.

The workspace contains unrelated untracked material such as `.claude/`, attachments, documents, research, and build output. Stage only files relevant to the requested change. Do not commit private attachments or unrelated work by using an indiscriminate `git add .`.

## 9. Future work — not implemented or automatically authorized

1. Separate maintenance assumptions for PV and battery/inverter systems, using explicit supplier information or clearly labeled estimates.
2. Improve net-savings/cost wording and distinguish negative financial values from configuration errors more clearly.
3. Expand the catalog to support the manager designs and new use cases without overwriting unrelated models; update fixed references and validations together.
4. Make check counts less repetitive and consider more precise per-scenario formula-error reporting. If changing result gating, design it in the calculation logic and avoid circular dependencies through audit outputs.
5. Establish a user-preferred manual capture layout on `Saved Scenarios` only if requested. It is intentionally blank now.
6. Add actual outage valuation, TOU/peak-demand tariffs, or hourly dispatch only with required data and separate verification.
7. Automatic sizing/search was discussed conceptually, but no optimization algorithm was implemented. The user deferred it and chose manual scenario exploration.
8. Verify Microsoft Excel operations directly when access becomes available, especially worksheet creation and manual copying/pasting.

## 10. Continuation procedure

1. Read this document and inspect `git status`, the active branch, and the authoritative workbook. Check for a newer user-supplied file and content in `Future Changes` / `Saved Scenarios`.
2. Restate the requested next change in Egyptian Arabic. Do not treat the backlog as permission to implement everything.
3. Preserve existing inputs, saved scenarios, formulas, native features, and sheet structure unless the current request authorizes changing them.
4. Keep assumptions explicit. Do not change them simply to obtain positive NPV or reproduce a desired answer.
5. For calculation changes, independently verify affected energy balances, equipment limits, costs, bills, NPV, and payback. For formatting or blank-tab changes, use proportionate structural/preservation checks.
6. Validate the final XLSX package and report the scope of native-engine testing honestly. Check new views visually when they change.
7. Deliver the updated authoritative workbook. Push to `excel-sheet-branch` only when asked, replacing the existing workbook rather than proliferating copies.
8. Update this handoff with implemented changes, remaining limitations, verified commit/file identity, and any user decisions. Never mark a proposal as completed without evidence.

Suggested opening prompt for a future chat:

> Read PROJECT_HANDOFF.md, then inspect the referenced workbook and current Git state. Explain any difference between this handoff and the actual files before making changes. Communicate with me in Egyptian Arabic and keep file content in English. Preserve my current inputs and saved scenarios. My next request is: [describe the change].
