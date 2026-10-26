# Solar Production Specification - Experimental Workbook

Date: 26 October 2026 (Africa/Cairo).
Status: implemented for trial; not approved to replace the original workbook.

## Scope and ownership

- Original: `outputs/scenario-comparison/AVESTA_El_Mo3alem_Eight_Scenarios_Excel_Compatible.xlsx`.
- Experimental copy: `outputs/solar-trial/AVESTA_Cairo_Seasonal_Solar_Trial.xlsx`.
- New worksheet: `Solar Production`. Existing five worksheets and historical saved scenarios are retained.
- The user subsequently authorized publishing the trial workbook and its documentation to GitHub on `excel-sheet-branch`. This does not authorize replacing the original workbook or adopting the trial as authoritative.
- User decision: Cairo is fixed. Panel tilt and orientation must be editable and affect production. Display monthly production and bill savings, followed by annual totals.

## Data source and fixed assumptions

Source: European Commission Joint Research Centre, PVGIS 5.3 `PVcalc` API.

- Latitude: 30.0444 degrees north; longitude: 31.2357 degrees east (representative Cairo city location, not a surveyed roof).
- Radiation database: PVGIS-SARAH3, climate period 2005-2023 as returned by the API.
- Meteorological database: ERA5.
- Reference capacity: 1 kWp.
- Module technology: crystalline silicon (`crystSi`).
- Mounting: free-standing (`free`), fixed plane.
- Terrain horizon: enabled. Nearby buildings, trees and row shading are not surveyed or modeled.
- PVGIS system-loss input: 14%. This is a modeling assumption, not a measured site loss. PVGIS also models its other PV performance effects; 14% is not the total combined loss.
- Retrieved: 26 October 2026.

Default source request:
<https://re.jrc.ec.europa.eu/api/v5_3/PVcalc?lat=30.0444&lon=31.2357&peakpower=1&loss=14&angle=30&aspect=0&outputformat=json&raddatabase=PVGIS-SARAH3&pvtechchoice=crystSi&mountingplace=free&usehorizon=1>

Documentation:
<https://joint-research-centre.ec.europa.eu/photovoltaic-geographical-information-system-pvgis/using-pvgis-5/pvgis-5-user-manual_en>

The original model's 4.98 kWh/kWp/day is consistent, at its rounded precision, with this default query. The new workbook uses monthly values instead of forcing the previous annual rounded value.

## User controls and offline behavior

- `Solar Production!C5`: tilt from horizontal, 0-90 degrees. Default 30.
- `C6`: compass azimuth clockwise from north, 0-360 degrees. Default 180 (south); 90 east, 270 west, 0/360 north.
- Conversion to PVGIS aspect: `compass azimuth - 180`. PVGIS uses 0 south, -90 east, +90 west.
- `C7`: displayed study year, integer 1 through the study horizon, maximum 20 consistent with the existing model.
- `C8`: displayed calendar year, linked to the existing first forecast year.
- `C12`: input validity. Invalid solar inputs produce unavailable numeric results rather than plausible zeros.
- The same orientation applies to both live PV configurations. Multiple roof planes are outside this trial.

The workbook works without Python, macros, an add-in, or an internet connection. It includes 475 PVGIS configurations: tilt every 5 degrees from 0 through 90, and compass azimuth every 15 degrees from 0 through 360. The 5,700 monthly source rows are visible in `Solar Production!A70:F5769`, with individual source URLs. Original API JSON responses are retained locally in `build/solar-trial/data` for reproducibility.

At grid points, values equal the corresponding published monthly API values. Between grid points, ordinary Excel formulas perform bilinear interpolation. This is an explicit approximation, not a fresh PVGIS computation. Opening the dynamic source hyperlink requests the precise current angles from PVGIS but does not refresh the embedded grid automatically. Changing climate database, fixed loss assumption, technology or location requires regenerating and validating the source grid.

For lower and upper tilt `t0,t1` and azimuth `a0,a1`:

```
u = (tilt - t0) / (t1 - t0)
v = (azimuth - a0) / (a1 - a0)
Y_m = (1-u)*(1-v)*Y_m(t0,a0)
    + u*(1-v)*Y_m(t1,a0)
    + (1-u)*v*Y_m(t0,a1)
    + u*v*Y_m(t1,a1)
```

Boundary intervals use 85-90 degrees tilt and 345-360 degrees azimuth where needed, avoiding division by zero. The 0 and 360 degree north endpoints are both included.

## Energy calculations and units

`Y_m` is the interpolated PVGIS monthly electrical energy for the 1 kWp reference, expressed as kWh/kWp/month. It is not solar irradiance, panel efficiency, or daylight duration. Do not apply another generic performance ratio or PV conversion efficiency.

Use PVGIS monthly `E_m` rather than its more coarsely rounded daily `E_d`. Define reference month days as a non-leap calendar: 31,28,31,30,31,30,31,31,30,31,30,31.

```
daily_yield_m = Y_m / reference_days_m
installed_kWp = panel_quantity * panel_Wp / 1000
generation_m_y = installed_kWp * daily_yield_m * calendar_days_m_y
                 * (1 - annual_PV_degradation)^(study_year - 1)
annual_generation_y = SUM(generation_m_y for the 12 months)
```

The first forecast year and annual degradation remain linked to existing assumptions (defaults 2027 and 0.5%). Leap years add the applicable February day. This calendar normalization is a planning convention, not a reconstruction of PVGIS's underlying historical calendar.

Default reference monthly values (kWh/kWp):

| Month | Yield |
|---|---:|
| January | 131.30 |
| February | 129.45 |
| March | 161.06 |
| April | 160.86 |
| May | 162.72 |
| June | 158.62 |
| July | 164.66 |
| August | 166.80 |
| September | 161.00 |
| October | 154.53 |
| November | 136.97 |
| December | 130.40 |
| Sum | 1818.37 |

PVGIS reports annual `E_y = 1818.36`; the 0.01 difference is rounding of published monthly values. The workbook consistently sums the monthly series.

## Integration and result meanings

- Replace monthly generation formulas in `Scenarios!F161:F400` and `F661:F900` with the monthly yield lookup. Downstream consumption, bill, cash-flow, NPV and payback formulas remain linked through the existing model.
- `Assumptions and Catalog!C31` becomes a derived annual-average yield, retained for legacy diagnostics and reference. It no longer drives monthly generation.
- `Solar Production!A19:J32` shows monthly specific yield, solar-only generation and bill savings, combined-system generation and bill savings, battery-only bill savings, and annual totals for the selected study year.
- Annual net savings are linked separately in row 36, including the existing annual maintenance and replacement treatment. Monthly bill savings do not subtract these annual costs.
- `Scenarios!A35` links to the monthly results. Existing top-level results continue to show year one, as labeled.
- Bill savings = existing baseline bill minus existing post-installation bill. They are not total generation multiplied by an assumed tariff.
- Unused PV has no export revenue. Tariff and equipment assumptions are unchanged and are not newly verified market inputs.
- Historical saved scenarios are static records of the previous method and do not update.

## Limits and next agreed step

This trial improves seasonality and orientation dependence. It does not implement hourly PV/load coincidence or hourly battery state of charge. Existing daytime share, 12-hour daytime duration and six-hour PV charging window remain simplified assumptions. Monthly aggregation can overstate direct use and battery benefit.

**Next step requested by the user: evaluate and display battery benefit using hourly calculations, not monthly budgets.** Implement only after reviewing this solar-production trial, with aligned PV and load timestamps, charging/discharging limits, reserve, efficiency and state-of-charge carryover.

Inverter-specific clipping, DC-coupled battery routing, nearby shading, orientation-dependent roof spacing, weather forecasts and probability-of-exceedance estimates are outside this trial. PVGIS climate averages are not a prediction of the weather in a specific future calendar year.

## Verification

The original workbook is hash-checked for preservation. Existing cells outside the targeted formula/documentation changes are compared before final cache transfer; original native workbook parts are retained instead of replacing the package with a general import/export.

Verification uses LibreOffice recalculation, independent Python energy/bill/cash-flow calculations, direct PVGIS spot checks of interpolation, Open XML SDK schema validation, and visual inspection. Microsoft Excel desktop interaction is not claimed unless separately tested.

Completed: 11 recalculated cases passed (baseline, east, west, horizontal north/south, vertical north at both compass endpoints, intermediate angles, leap year, invalid tilt and zero PV). Independent checks covered monthly production/direct use/battery flows/losses/grid imports/bills/energy balance, annual cash flows, NPV, simple and discounted payback, and displayed annual totals. Final workbook passed Open XML SDK validation with zero errors. The new worksheet was visually inspected; the renderer's unsupported HYPERLINK function was replaced with its LibreOffice-calculated display label in previews only. Actual workbook links remain native Excel formulas.

Direct API validation of the 5-degree / 15-degree grid:

| Tilt | Compass azimuth | Annual interpolation error | Largest absolute monthly error |
|---:|---:|---:|---:|
| 23 | 113 | -0.1501% | 0.2489% |
| 37 | 202 | -0.1750% | 0.4571% |
| 67 | 303 | -0.0163% | 1.2758% |

These are three spot checks, not a guaranteed bound over all angles. Solar-resource and equipment-model uncertainty are separate from interpolation error. At horizontal tilt, separately requested compass endpoints differ by 0.04 kWh/kWp/year in the rounded PVGIS source; the source symmetry check permits 0.10 kWh/kWp/year while spreadsheet equation checks use 0.000001 absolute tolerance.
