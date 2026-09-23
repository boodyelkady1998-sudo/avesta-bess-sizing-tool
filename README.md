# AVESTA C&I BESS Sizing Tool

A responsive, browser-based sizing tool for Commercial & Industrial battery energy storage systems, including an Egypt-focused Retail & Kiosk Load Survey.

## Use the app

Open `index.html` in a browser, or visit the GitHub Pages link after Pages is enabled.

## Notes

- Works on desktop and mobile browsers.
- Runs entirely in the browser.
- Core sizing calculations are contained in `index.html`.
- Includes Simple and Engineer/Advanced survey modes.
- Includes editable cooler, freezer, display, lighting, CCTV/POS, air-conditioning, and custom-load presets.
- Calculates connected load, compressor startup peak, daily energy, required battery capacity, and recommended inverter rating.
- Uses Egypt defaults of 230 V / 50 Hz with Cairo and Giza ambient presets.
- Stores survey edits and optional compressed nameplate photos locally in the browser.

## Home Tier Sizing (`tiers.html`)

A standalone residential page, linked from the main navigation as **Home Tier Sizing**.

- Input: the customer's last bill (EGP) or last month's consumption (kWh).
- Converts a bill to kWh using the Egyptian residential tariff tiers (fees and future increases ignored).
- Compares Grid only, Solar only and Solar + AVESTA LV battery: new bill, tier, system price, savings and simple payback.
- Supports self-consumption and net-metering scenarios.
- Tariff rates, customer prices and sizing assumptions are editable on the page and stored locally in the browser.
