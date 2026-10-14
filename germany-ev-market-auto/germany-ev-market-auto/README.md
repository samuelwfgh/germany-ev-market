# Germany EV / car-market dashboard

A static GitHub Pages site inspired by Robbie Andrew's Norway EV page, adapted for the German passenger-car market.

The site is deliberately serverless. `index.html` renders the charts, while a scheduled GitHub Action refreshes the repository's JSON datasets. If an upstream source changes or cannot be parsed, the updater **keeps the last known-good data** instead of overwriting the site with zeros or guessed values.

## Repository layout

```text
germany-ev-market/
├── index.html
├── requirements.txt
├── README.md
├── DATA_SOURCES.md
├── .nojekyll
├── data/
│   ├── new_registrations.json
│   ├── used_transfers.json
│   ├── road_fuels.json
│   ├── charging_infrastructure.json
│   ├── hydrogen_stations.json
│   └── metadata.json
├── scripts/
│   ├── update_data.py
│   └── validate_data.py
└── .github/
    └── workflows/
        └── update-data.yml
```

## GitHub setup

1. Upload **the complete folder contents** to the root of your existing `germany-ev-market` repository. Preserve the `.github/workflows/` and `data/` folders.
2. Keep GitHub Pages on **Settings → Pages → Deploy from a branch → `main` → `/ (root)`**.
3. Open **Settings → Actions → General → Workflow permissions**. If your repository/account requires it, choose **Read and write permissions** so the updater can commit refreshed JSON files.
4. Open **Actions → Update car-market data → Run workflow** once. This initial run fills/refreshes the datasets immediately instead of waiting for the daily schedule.
5. After the workflow commits a data change, GitHub Pages will redeploy from `main` automatically.

The workflow runs every day at **06:17 UTC**. It does not create a commit if nothing changed.

## How the updater works

`scripts/update_data.py` checks the following sources:

- New registrations: long-run German rows from the consolidated monthly car-sales dataset, with a current KBA-derived monthly overlay when it can be parsed safely.
- Used transfers: KBA FZ 9 workbooks.
- Road fuels: BAFA *Amtliche Mineralöldaten*.
- Charging infrastructure: Bundesnetzagentur.
- Hydrogen stations: documented IPHE Germany snapshots.

The script validates ranges and totals before writing data. A failed parser/source request becomes a warning in `data/metadata.json`; the previous data remain untouched.

Hydrogen is intentionally **not interpolated monthly** because a consistent official monthly operational-series does not exist. Only documented snapshots are appended.

Road-fuel data can have a substantial publication lag and revisions. Missing petrol/diesel months remain `null`; they are never filled by estimation.

## Run locally

Because the page loads JSON files with `fetch()`, serve the folder instead of opening `index.html` directly from `file://`.

```bash
python -m http.server 8000
```

Then open `http://localhost:8000/`.

To update manually:

```bash
python -m pip install -r requirements.txt
python scripts/update_data.py --data-dir data --verbose
python scripts/validate_data.py data
```

To validate the currently committed files without internet access:

```bash
python scripts/update_data.py --data-dir data --no-network
```

## Updating logic if a source changes

The parsers are intentionally conservative. If KBA/BAFA changes a workbook layout, the action should show a warning and preserve the existing JSON. Inspect the failed Action log, update the corresponding parser in `scripts/update_data.py`, run the validator, and commit the parser fix.

## Data philosophy

The site should prefer **missing data over invented data**. Derived values are only used when the arithmetic is explicit and the underlying figures are published (for example, HEV = total hybrids minus PHEV when both source values are present). No historical gaps are silently interpolated.
