# ENTSO-E Pumped Storage & Nuclear Power Analysis

Python scripts to query the [ENTSO-E Transparency Platform](https://transparency.entsoe.eu) for pumped storage and nuclear power generation data across Switzerland, Austria and Germany.

## Background

This project was motivated by a data transparency finding: Switzerland, as a non-EU member, does not report pumped storage **consumption** (pumping direction) to ENTSO-E, while Austria and Germany (EU members, bound by Transparency Regulation 543/2013) report both directions fully and at 15-minute resolution.

A formal inquiry was sent to **Swissgrid** and **ElCom** (Swiss electricity regulator) in June 2026. A related GitHub issue was opened in [electricitymaps-contrib](https://github.com/electricitymaps/electricitymaps-contrib) regarding incorrect Swiss nuclear data on electricitymaps.com.

## Scripts

| File | Description |
|---|---|
| `akw_leistung_ch_v5.py` | Swiss nuclear plants per block (Beznau 1/2, Gösgen, Leibstadt) + pumped storage CH. Includes derived Gösgen estimate via Total_Nuclear_A75 − other blocks. Exports 3 Excel sheets: AKW (raw data), AKW_Grafik (Box-Whisker chart: Total_Nuklear vs. Leibstadt), Pumpspeicher (Box-Whisker chart: Turbiniert vs. Hochgepumpt). |
| `ch_pumpspeicher_v4.py` | Pumped storage Switzerland: Turbiniert_MW + Hochgepumpt_MW (pumped column always empty – CH doesn't report) |
| `at_pumpspeicher_v4.py` | Pumped storage Austria: Turbiniert_MW + Hochgepumpt_MW (both fully reported, 15 min resolution) |
| `de_pumpspeicher_v4.py` | Pumped storage Germany: Turbiniert_MW + Hochgepumpt_MW (both reported, multiple TSO zones aggregated) |
| `boxwhisker_injector.py` | Helper module: injects a Box-Whisker chart (chartEx format) into an xlsx file post-generation, since openpyxl doesn't support this chart type natively |

## Shell Wrappers

To avoid having to manually activate the virtual environment before each run, use the shell wrappers:

```bash
# Make executable once after first download
chmod +x run_akw.sh run_at_pumpspeicher.sh run_ch_pumpspeicher.sh run_de_pumpspeicher.sh

# Then simply run (no source venv/bin/activate needed)
./run_akw.sh --week
./run_at_pumpspeicher.sh --week
./run_ch_pumpspeicher.sh --week
./run_de_pumpspeicher.sh --week
```

| Wrapper | Calls |
|---|---|
| `run_akw.sh` | `akw_leistung_ch_v5.py` |
| `run_ch_pumpspeicher.sh` | `ch_pumpspeicher_v4.py` |
| `run_at_pumpspeicher.sh` | `at_pumpspeicher_v4.py` |
| `run_de_pumpspeicher.sh` | `de_pumpspeicher_v4.py` |

## Installation

```bash
python3 -m venv venv
source venv/bin/activate
pip install entsoe-py pandas openpyxl python-dotenv
```

## Setup: ENTSO-E API Token

### 1. Get a token
1. Register a free account at https://transparency.entsoe.eu
2. Request API access by email to transparency@entsoe.eu (subject: "Restful API access") — usually activated within 1–3 working days
3. Generate your token under "My Account Settings" → "Web Api Security Token"

### 2. Store the token in a `.env` file

Create a file named `.env` in the project directory (same folder as the scripts):

```
ENTSOE_TOKEN=your-token-here
```

> ⚠️ **Never commit the `.env` file to a public repository.** It is already listed in `.gitignore` and will be automatically excluded from all commits.

The scripts read the token automatically via `python-dotenv` — no manual export or environment variable setup needed.

## Usage

```bash
# Live snapshot (current values)
./run_akw.sh

# Weekly export as Excel with Box-Whisker charts
./run_akw.sh --week
./run_at_pumpspeicher.sh --week

# Custom time range (e.g. 14 days)
./run_at_pumpspeicher.sh --week --days 14

# Debug: show raw XML response
./run_at_pumpspeicher.sh --raw
```

## Key Findings

### Pumped Storage Comparison

| Country | Turbiniert_MW | Hochgepumpt_MW | Resolution | Completeness |
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

### Notable Events (June 2026)

- **Beznau heat shutdown:** Beznau 1 & 2 were fully shut down on 26 June 2026 because the Aare river temperature exceeded 25°C — visible in ENTSO-E data immediately, but not in electricitymaps.com for several days.
- **Gösgen restart:** After a 10-month outage, Gösgen received restart clearance from ENSI on 26 June 2026.

## Technical Notes

### boxwhisker_injector.py

Excel's Box-Whisker chart type (`chartEx`) is not supported by openpyxl. The injector module works around this by re-opening the `.xlsx` file as a ZIP archive and injecting the `chartEx` XML parts directly. Supports multiple charts per workbook via `sheet_index` parameter.

### Rate Limiting

The scripts use 3-second pauses between daily requests and 5-second pauses between query blocks to avoid ENTSO-E rate limits.

## License

MIT License — see [LICENSE](LICENSE)

## Author

Peter Berger — Bergi IT Consulting, Zürich, Switzerland
