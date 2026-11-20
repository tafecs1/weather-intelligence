# India Weather Intelligence (GitHub build)

Runs `weather.py` (the Weather.ipynb logic, unchanged) every day on GitHub Actions and writes the
7 tabs to the fixed Weather Google Sheet. Stand-alone: it does not read anything from the Agri,
Crop or Groundwater repos.

## What changed vs the notebook (all marked `[GH]` in weather.py)
1. **Name standard**: State / District / Tehsil printed with the official LGD spelling; original
   TEHSIL_MASTER spelling kept in `Source state/district/tehsil name` columns. `STANDARDIZE_NAMES=0`
   reproduces the notebook output exactly.
2. **Tehsil code clean-up**: zero-padded TEHSIL_MASTER codes ("05796") aligned to official LGD codes; duplicate rows
   left under an old (pre-split) parent district are removed, or moved to the official district if no correct copy exists.
   Every action is listed in LGD NAME MATCHING AUDIT. `FIX_TEHSIL_CODES=0` turns it off.
3. **IMERG diagnostics**: the satellite-rainfall check looks at the last 10 days only (same result, no full-archive scan)
   and logs the reason for every dataset tried.
4. **RUN_STATES env**: lets you test one state from the Actions "Run workflow" box.
Nothing else: same sources, thresholds, scores, tabs, sheet IDs.

## Outputs
| Where | What |
|---|---|
| Google Sheet (`WEATHER_SHEET_ID`) | the 7 tabs: WEATHER MANAGEMENT SUMMARY, INDIA DISTRICT INTELLIGENCE, INDIA TEHSIL INTELLIGENCE, APPROACHING WEATHER ALERTS, HISTORY MOM DISTRICT / TEHSIL / STATE |
| Actions artifact (90 days) | full Excel (all tabs), forecast vintage, audit and logs |
| Committed to repo `data/` | master caches, alert history, IMD bulletin records (kept between runs) |

## Secrets (Settings → Secrets and variables → Actions)
| Secret | Required | Notes |
|---|---|---|
| `GEE_SERVICE_ACCOUNT_JSON` | yes | whole key JSON |
| `GEE_PROJECT` | yes | Earth Engine Cloud project ID |
| `WEATHER_SHEET_ID` | yes | sheet shared with the service account as **Editor** |
| `TEHSIL_MASTER_SHEET_ID` | only if no CSV | or commit `data/masters/TEHSIL_MASTER.csv` (preferred: fully independent) |
| `IMD_API_KEY`, `IMD_AUTH_SCHEME` | no | IMD stays off until provided |

## Cross-sheet name rule (use the same in all 4 repos)
Join on **LGD codes**, never on names: `LGD state code`, `LGD district code`, `LGD sub-district code`.
Names printed = official LGD names (India Data Portal CKAN resource `42cd589e-2dc4-4f9c-a245-f201cee357e7`).
