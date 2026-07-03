# ENTSO-E Pumped Storage & Nuclear Power Analysis

Python scripts to query the [ENTSO-E Transparency Platform](https://transparency.entsoe.eu) for pumped storage and nuclear power generation data across Switzerland, Austria and Germany.

## Background

This project was motivated by a data transparency finding: Switzerland, as a non-EU member, does not report pumped storage **consumption** (pumping direction) to ENTSO-E, while Austria and Germany (EU members, bound by Transparency Regulation 543/2013) report both directions fully and at 15-minute resolution.

A formal inquiry was sent to **Swissgrid** and **ElCom** (Swiss electricity regulator) in June 2026. A related GitHub issue was opened in [electricitymaps-contrib](https://github.com/electricitymaps/electricitymaps-contrib) regarding incorrect Swiss nuclear data on electricitymaps.com.

## Scripts

| File | Description |
|---|---|
| `akw_leistung_ch_v5.py` | Swiss nuclear plants per block (Beznau 1/2, Gösgen, Leibstadt) + pumped storage CH. Includes derived Gösgen estimate via Total_Nuclear_A75 − other blocks. |
| `ch_pumpspeicher_v4.py` | Pumped storage Switzerland: Turbinated_MW + Pumped_MW (pumped column always empty – CH doesn't report) |
| `at_pumpspeicher_v4.py` | Pumped storage Austria: Turbinated_MW + Pumped_MW (both fully reported, 15 min resolution) |
| `de_pumpspeicher_v4.py` | Pumped storage Germany: Turbinated_MW + Pumped_MW (both reported, multiple TSO zones aggregated) |
| `boxwhisker_injector.py` | Helper module: injects a Box-Whisker chart (chartEx format) into an xlsx file post-generation, since openpyxl doesn't support this chart type natively |

## Installation

```bash
python3 -m venv venv
source venv/bin/activate
pip install entsoe-py pandas openpyxl
```

## Setup: ENTSO-E API Token

1. Register a free account at https://transparency.entsoe.eu
2. Request API access by email to transparency@entsoe.eu (subject: "Restful API access") — usually activated within 1–3 working days
3. Generate your token under "My Account Settings" → "Web Api Security Token"
4. Add your token to the scripts: `ENTSOE_TOKEN_HARDCODED = "your-token-here"`

> ⚠️ **Never commit your token to a public repository.** The `.gitignore` file excludes the scripts by default — if you want to track them, remove the token first.

## Usage

```bash
# Live snapshot (current values)
python3 at_pumpspeicher_v4.py

# Weekly export as Excel with Box-Whisker chart
python3 at_pumpspeicher_v4.py --week
python3 ch_pumpspeicher_v4.py --week
python3 de_pumpspeicher_v4.py --week

# Custom time range (e.g. 14 days)
python3 at_pumpspeicher_v4.py --week --days 14

# Custom output filename
python3 at_pumpspeicher_v4.py --week --out my_output.xlsx

# Debug: show raw XML response
python3 at_pumpspeicher_v4.py --raw
```

## Key Findings

### Pumped Storage Comparison

| Country | Turbinated_MW | Pumped_MW | Resolution | Completeness |
|---|---|---|---|---|
| Switzerland (CH) | ✅ available | ❌ not reported | 1 hour | Incomplete |
| Austria (AT) | ✅ available | ✅ available | 15 minutes | Nearly complete |
| Germany (DE) | ✅ available | ✅ available | 15 minutes | Nearly complete |

### Swiss Nuclear Plants (June 2026)

| Plant | Net Capacity | ENTSO-E Reporting |
|---|---|---|
| Beznau 1 (Axpo) | 365 MW | A73 data available (small gaps) |
| Beznau 2 (Axpo) | 365 MW | A73 data available (small gaps) |
| Kernkraftwerk Gösgen | ~1,010 MW | A73 data largely absent → derived from total |
| Leibstadt | ~1,233 MW | A73 data complete and reliable |

**Gösgen derivation:** `Gösgen_derived = Total_Nuclear_A75 − Beznau1 − Beznau2 − Leibstadt`
Results in NaN (rather than a potentially incorrect value) when any other block is missing for a given timestamp.

### Notable Events (June 2026)

- **Beznau heat shutdown:** Beznau 1 & 2 were fully shut down on 26 June 2026 because the Aare river temperature exceeded 25°C (cooling water limit) during a heat wave — visible in the ENTSO-E data immediately, but not reflected in electricitymaps.com for several days.
- **Gösgen restart:** After a 10-month outage (safety upgrades + annual revision), Gösgen received restart clearance from ENSI on 26 June 2026.

## Technical Notes

### boxwhisker_injector.py

Excel's Box-Whisker chart type (`chartEx`) is not supported by openpyxl. The injector module works around this by:
1. Writing the data normally with pandas/openpyxl
2. Re-opening the `.xlsx` file as a ZIP archive
3. Injecting the `chartEx` XML parts (chart definition, style, colors, relationships) directly

This approach is based on a reference file created manually in Excel. Known limitations:
- Axis titles cannot be set programmatically (undocumented schema restriction)
- Compatibility with older Excel versions, Google Sheets or LibreOffice is not guaranteed
- The module must be in the same directory as the calling scripts

### Rate Limiting

The scripts use 3-second pauses between daily requests and 5-second pauses between query blocks to avoid hitting ENTSO-E rate limits. A 401 HTTP error (rather than the expected 429) was observed when making ~21 rapid consecutive requests — hence the conservative pacing.

## License

MIT License — see [LICENSE](LICENSE)

## Author

Peter Berger — Bergi IT Consulting, Zürich, Switzerland
