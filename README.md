# Peplink Usage Tracker

Small Flask app for Peplink captive-portal usage reporting. It joins a **Peplink event log** (.xlsx, "which user logged in from which MAC")
with a **per-device usage report** (.csv, "usage per MAC") and produces one list of
**User ID → MAC → usage**, viewable in the browser or downloaded as Excel.

## Features
- **Ongoing tracking (Import / Report pages)**: upload event-log exports and daily usage CSVs whenever you have them — stored in SQLite (`data/peplink.db`). Overlapping event-log exports are deduplicated (one row per login); re-importing a day's usage replaces it. Multiple sites/vessels supported (site read from file name, e.g. `SITE1`). Report page: pick a date range → per-user-per-day pivot, device-day detail, Excel download. Import page shows days loaded, gaps, and event-log coverage.
- **Total period / day wise**: the Report page (and Quick merge with several daily files) opens on **Total period** — one row per user and per device for the whole range — and switches to **Day wise** (one column per day, with a Total / Upload / Download / Upload-Download selector). Upload and download are shown next to the total everywhere, with GB uploaded / downloaded tiles. The Excel buttons follow the view (total-period workbook, or a day-wise workbook with Total, Upload and Download sheets); the view is kept in the address (`?view=day`).
- **Per device per day**: the report's *Per device per day* tab and Excel sheet list every device (MAC, hostname, attributed user(s)) with its usage for each day.
- **Per-day mode**: upload several daily *Top client usage* exports (date taken from the file name) → usage per user **per day**, pivot sheet + device-day detail. Each device-day is attributed to the account that logged in on that device that day (most logins wins); if nobody logged in that day, the last earlier login is carried over and marked as such; if the event log only starts after that day (the portal export is capped at 10,000 rows), the next known login is used and marked *Inferred*
- Upload both files in the browser → on-screen preview (sortable, filterable) or Excel download
- Excel has 3 sheets: *Device usage by user*, *Usage per user* (rolled up per account), *Unmatched devices*
- Only successful logins (`guest account ID "x" [login]`) are counted; failed passwords / quota events are ignored
- A MAC used by several accounts shows the most-frequent one as primary, plus all accounts seen with login counts and the last login
- Duplicate MAC rows in the usage CSV are summed; MACs are normalised (case, `-` vs `:`)
- Works with any site export (SITE1, other sites) — column names are matched case-insensitively
- CLI mode for scripting/batch use
- Also accepts the portal's *Top client usage* export (`% Usage` column carried through); rows keyed by IP instead of MAC are listed under Unmatched with a note
- **7-day and daily uploads**: the Import page takes daily usage CSVs (one day each) and 7-day/period CSVs (one export for a whole period; dates from the file name or a start-date field). Period files are stored separately and shown as per-device and per-user period totals (with Excel download).
- **Sites overview**: the Import page lists every site with its last entry added (when, kind, file, period covered) plus latest usage day and latest login; click a site (or pick it in the Site box) and its import status loads below without a page reload.
- **Delete data**: Import page → *Delete data* (bottom of the page) removes daily usage, period usage, logins, or everything for a site, optionally limited to a date range (everything needs the site name typed).
- **Login required** (one or more accounts): user name and password come from the host's `.env` (never from the repo). With no `ADMIN_PASSWORD` set the app stays locked (503) instead of open. CSRF-protected forms, 5 failed sign-ins lock that address out for 5 minutes, sessions last 12 hours. `/health` stays open for the container healthcheck.
- The *Quick merge* page stores nothing (files are processed in memory); the *Import* / *Report* pages keep their data in SQLite (see Data)

## Configuration (`.env`)
Secrets and site-specific values live in an untracked `.env` next to `docker-compose.yml`; `.env.example` (tracked) documents them. `docker-compose.yml` reads them as `${VAR:-default}`, so nothing sensitive is in the repo and sharing the code shares no credentials.

| Variable | Purpose |
|---|---|
| `ADMIN_USERNAME` | sign-in user name (default `admin`) |
| `ADMIN_PASSWORD` | sign-in password — **required**; the app is locked while empty |
| `ADMIN_USERS` | optional extra accounts, `user:password;user2:password2` |
| `SECRET_KEY` | signs the login cookie; if empty, one is generated into `data/.secret_key` |

Create it once on the host:
```bash
cd peplink-usage-tracker
printf 'ADMIN_USERNAME=admin\nADMIN_PASSWORD=%s\nSECRET_KEY=%s\n' "$(openssl rand -base64 18 | tr -d '/+=')" "$(openssl rand -hex 32)" > .env
chmod 600 .env && cat .env        # note the generated password, then: docker compose up -d
```

## Data
`data/peplink.db` holds everything imported. The compose file mounts `./data` so it survives rebuilds; back it up with your other SQLite DBs.

## Download and install

**Requirements:** Docker with Docker Compose (recommended), or Python 3.10+ to run without Docker.

1. **Download** the code - either clone it (`git clone <repository-url>`) or use the repository's *Code → Download ZIP* button and unzip it.
2. **Configure** (once). In the project folder, copy the example settings and choose a password:
   ```bash
   cd peplink-usage-tracker           # the folder that contains docker-compose.yml
   cp .env.example .env
   ```
   Edit `.env` and set `ADMIN_PASSWORD` (required - the app stays locked while it is empty). Optionally set `ADMIN_USERNAME`, extra accounts in `ADMIN_USERS`, and `SECRET_KEY` (see *Configuration* above).
3. **Start**:
   ```bash
   docker compose up -d --build
   ```
4. **Open** `http://localhost:5020` (or `http://<your-host>:5020` from another machine) and sign in with the user name and password from `.env`.
5. **Use it**: *Import* your event-log `.xlsx` and daily usage `.csv` exports, then open *Report*. Imported data is kept in `./data` and survives restarts and rebuilds.

**Update:** download or pull the newer version, then run `docker compose up -d --build` again; `data/` and `.env` are untouched.
**Stop:** `docker compose down`. **Logs:** `docker compose logs -f`. **Health check:** `http://localhost:5020/health`.

No Docker? See *Run locally without Docker* below.

## Run locally without Docker
```bash
pip install -r app/requirements.txt
ADMIN_PASSWORD=test python app/app.py     # http://localhost:5020  (Windows PowerShell: $env:ADMIN_PASSWORD='test'; python app/app.py)
```

## CLI
```bash
python app/app.py merge "SITE1 - Starlink event log.xlsx" "SITE1 usage.csv" -o SITE1_usage.xlsx   # one-off totals
python app/app.py merge "SITE1 - Starlink event log.xlsx" SITE1_top_client_usage_2026-10-0*.csv            # one-off per-day
python app/app.py import *_SITE1_-_Starlink_*.csv "SITE1 - Starlink event log.xlsx"                         # into the store
python app/app.py report --view day --site SITE1 --start 2026-10-03 --end 2026-10-09 -o SITE1_week.xlsx
```

## Changelog
- **v1.7.1** (2026-10-14) — 7-day/period page (and single-file Quick merge) now match the report: Per user / Per device / Unmatched tabs with upload and download, bars, TOTAL rows, GB uploaded / downloaded / unmatched tiles; the period Excel's *Usage per user* and *Unmatched devices* sheets gain Upload and Download columns. **Fix:** the period page's *Download Excel* failed with an error since v1.6.2. No schema or config change.
- **v1.7.0** (2026-10-14) — Report: **Total period** (default) and **Day wise** views; **upload and download** columns next to total in every table, per-user, per-device and unmatched, plus GB uploaded / downloaded tiles; day columns can show Total, Upload, Download or both; Excel follows the view (total-period workbook; day-wise workbook with Upload and Download sheets; *Download per device* follows the view too); CLI `report --view total|day` (default total). Page footer credit. No schema or config change.
- **v1.6.4** (2026-10-14) — Docs only: deployment notes and examples updated. No code, schema or config change.
- **v1.6.2** (2026-10-12) — **Download per device**: a separate one-sheet Excel of the per-device view, as a button on the Report page (per device per day), the 7-day/period page and Quick merge. The existing *Download Excel* is unchanged. No schema or config change.
- **v1.6.1** (2026-10-12) — 7-day/period page: new **Per device** tab (MAC, hostname, user, upload/download/total, GB, basis) and a matching *Usage per device* sheet first in the period Excel. No schema or config change.
- **v1.6.0** (2026-10-12) — New **Per device per day** view: a report tab (also on Quick merge with several daily files) and an Excel sheet *Usage per device per day* — one row per MAC with hostname, user(s) and usage per day, total, GB and days active. No schema or config change.
- **v1.5.3** (2026-10-10) — UI: page margins much smaller (full-width content, tighter cards); *Sign out* button aligned with the other nav links. No schema or config change.
- **v1.5.2** (2026-10-10) — Import page: *Usage files are* now defaults to **Daily**; the *Auto* option is removed, so a 7-day/period file must be chosen explicitly (dates still read from the file name or the start-date field). No schema or config change.
- **v1.5.1** (2026-10-10) — Import page: *Delete data* card moved to the bottom of the page, below Recent imports.
- **v1.5.0** (2026-10-10) — Import page: *Sites — last entry added* overview (last import per site, latest usage day and login); selecting a site loads that site's import status dynamically (new `/site_panel` fragment). No schema or config change.
- **v1.4.1** (2026-10-10) — Hostname (the usage export's *Name* column) now sits right after the MAC everywhere: the *Device name* column is renamed *Hostname* in the app and Excel, and the per-user MAC lists read `MAC (hostname)`. A usage file without a *Name*/*Hostname* column is accepted; the hostname is then left blank. No schema or config change.
- **v1.4.0** (2026-10-10) — **Multiple logins** (`ADMIN_USERS` in `.env`, signed-in name shown next to Sign out); **7-day/period usage upload** alongside daily files (new `usage_period` table, created automatically; period totals page + Excel); **Delete data** (daily usage, periods, logins or everything for a site, optional date range). `docker-compose.yml` changed: passes `ADMIN_USERS`. ONE-TIME STEP: optionally add `ADMIN_USERS=alice:pw1;bob:pw2` to `.env`; nothing else needed.
- **v1.3.0** (2026-10-10) — **Login** (user name/password from `.env`, CSRF on all forms, brute-force lockout, locked page while unconfigured); secrets moved out of `docker-compose.yml` into `.env` (`SECRET_KEY`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`) with `.env.example`; **app icon** (favicon + header + login page); gunicorn instead of Flask's dev server; compose healthcheck + log rotation; `.dockerignore`. Fixes: a bad date in `/report` no longer returns a 500; file type is decided by extension only (a CSV with "event" in its name was read as an event log); database connections are always closed. ONE-TIME STEP: create `.env` (see Configuration) before updating, or the app shows the "locked" page.
- **v1.2.1** (2026-10-10) — Renamed to *Peplink Usage Tracker* (DB file `data/peplink.db`, container `peplink-usage-tracker`); moved into its own git repo.
- **v1.2.0** (2026-10-09) — Ongoing tracking: SQLite store, Import page (dedup overlapping event logs, per-day usage upsert, multi-site), Report page with date range + Excel, CLI `import`/`report`; layout with nav; one-off merge moved to /quick.
- **v1.1.1** (2026-10-09) — Per-day mode: *Inferred from next login* fallback for days before the event log starts; count shown in summary.
- **v1.1.0** (2026-10-09) — Per-day mode: multiple daily usage CSVs → per-user-per-day pivot and device-day detail (web + CLI). Attribution by same-day login, with carry-over fallback.
- **v1.0.1** (2026-10-09) — Support *Top client usage* CSV export: `% Usage` column added to device sheet/preview; IP-only rows flagged in Unmatched; IP-only count on summary.
- **v1.0.0** (2026-10-09) — Initial release: web upload + preview + Excel download, CLI mode, Dockerfile/compose.
