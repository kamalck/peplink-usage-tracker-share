# Peplink Usage Tracker

Small Flask app for Peplink captive-portal usage reporting. It joins a **Peplink event log** (.xlsx, "which user logged in from which MAC")
with a **per-device usage report** (.csv, "usage per MAC") and produces one list of
**User ID → MAC → usage**, viewable in the browser or downloaded as Excel.

## Features
- **Ongoing tracking (Import / Report pages)**: upload event-log exports and daily usage CSVs whenever you have them — stored in SQLite (`data/peplink.db`). Overlapping event-log exports are deduplicated (one row per login); re-importing a day's usage replaces it. Multiple sites/vessels supported (site read from file name, e.g. `SITE1`). Report page: pick a date range → per-user-per-day pivot, device-day detail, Excel download. Import page shows days loaded, gaps, and event-log coverage.
- **Per-day mode**: upload several daily *Top client usage* exports (date taken from the file name) → usage per user **per day**, pivot sheet + device-day detail. Each device-day is attributed to the account that logged in on that device that day (most logins wins); if nobody logged in that day, the last earlier login is carried over and marked as such; if the event log only starts after that day (the portal export is capped at 10,000 rows), the next known login is used and marked *Inferred*
- Upload both files in the browser → on-screen preview (sortable, filterable) or Excel download
- Excel has 3 sheets: *Device usage by user*, *Usage per user* (rolled up per account), *Unmatched devices*
- Only successful logins (`guest account ID "x" [login]`) are counted; failed passwords / quota events are ignored
- A MAC used by several accounts shows the most-frequent one as primary, plus all accounts seen with login counts and the last login
- Duplicate MAC rows in the usage CSV are summed; MACs are normalised (case, `-` vs `:`)
- Works with any site export (SITE1, other sites) — column names are matched case-insensitively
- CLI mode for scripting/batch use
- Also accepts the portal's *Top client usage* export (`% Usage` column carried through); rows keyed by IP instead of MAC are listed under Unmatched with a note
- No data is stored; files are processed in memory

## Data
`data/peplink.db` holds everything imported. The compose file mounts `./data` so it survives rebuilds; back it up like any other SQLite database.

## Run with Docker
```bash
cd peplink-usage-tracker          # the folder containing docker-compose.yml
cp .env.example .env              # then set ADMIN_PASSWORD (see Configuration)
docker compose up -d --build
```
Open `http://<your-host>:5020`. To update after pulling new code, run `docker compose up -d --build` again; `data/` and `.env` are untouched.

## Run locally without Docker
```bash
pip install -r app/requirements.txt
python app/app.py                 # http://localhost:5020
```

## CLI
```bash
python app/app.py merge "SITE1 - Starlink event log.xlsx" "SITE1 usage.csv" -o SITE1_usage.xlsx   # one-off totals
python app/app.py merge "SITE1 - Starlink event log.xlsx" SITE1_top_client_usage_2026-10-0*.csv            # one-off per-day
python app/app.py import *_SITE1_-_Starlink_*.csv "SITE1 - Starlink event log.xlsx"                         # into the store
python app/app.py report --site SITE1 --start 2026-10-03 --end 2026-10-09 -o SITE1_week.xlsx
```

## Changelog
- **v1.2.1** (2026-10-10) — Renamed to *Peplink Usage Tracker* (DB file `data/peplink.db`, container `peplink-usage-tracker`); moved into its own git repo.
- **v1.2.0** (2026-10-09) — Ongoing tracking: SQLite store, Import page (dedup overlapping event logs, per-day usage upsert, multi-site), Report page with date range + Excel, CLI `import`/`report`; layout with nav; one-off merge moved to /quick.
- **v1.1.1** (2026-10-09) — Per-day mode: *Inferred from next login* fallback for days before the event log starts; count shown in summary.
- **v1.1.0** (2026-10-09) — Per-day mode: multiple daily usage CSVs → per-user-per-day pivot and device-day detail (web + CLI). Attribution by same-day login, with carry-over fallback.
- **v1.0.1** (2026-10-09) — Support *Top client usage* CSV export: `% Usage` column added to device sheet/preview; IP-only rows flagged in Unmatched; IP-only count on summary.
- **v1.0.0** (2026-10-09) — Initial release: web upload + preview + Excel download, CLI mode, Dockerfile/compose.
