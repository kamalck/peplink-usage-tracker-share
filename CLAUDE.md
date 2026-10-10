# Peplink Usage Tracker — project rules

## Git and releases
- Commit on `main` and push to `origin main`. Do not create feature branches or open pull requests unless asked.
- Commit each finished change as one commit with a clear subject line and a short body.
- If you change `docker-compose.yml`, the `Dockerfile` or `.dockerignore`, say so plainly in the commit message.
- If a change needs a one-time manual step on the host (moving a file, editing `.env`, creating a folder), spell it out in the commit body under a line **"ONE-TIME STEP:"**.
- Never rewrite history on `main` (no force-push, no amend of pushed commits).
- Do not edit files directly on the deployed host; deploy from git.

## Secrets and runtime data are NOT in git
- `.env` holds all secrets and site-specific values and is git-ignored. `docker-compose.yml` reads them as `${VAR:-default}`; a tracked `.env.example` documents every variable. **Never** put a password, key or token into any tracked file, and never commit `.env`, a key file, a `secrets/` folder or a `*.db` / `*.sqlite` file.
- `data/` (the SQLite database `peplink.db`, imports, exports) is git-ignored and bind-mounted at run time. Code changes must never assume it is empty, and schema changes must migrate an existing database in place.
- Current variables: `ADMIN_USERNAME`, `ADMIN_PASSWORD` (required — the app is locked while empty), `ADMIN_USERS` (optional extra accounts), `SECRET_KEY`. The sign-in is the only thing standing between the network and guest MAC/user data: never weaken the login gate in `app.py` (`gate()`), keep `/health` and `static` as the only open endpoints, and keep every POST form carrying the `_csrf` field.
- When you add a new environment variable: add it to `.env.example` with a comment, reference it in `docker-compose.yml` as `${NAME:-default}`, and mention in the commit body that it must be added to the host's `.env`.
- Never commit exported data files (CSV / XLSX) — they contain guest user IDs and MAC addresses.
- Never put host names, IP addresses, internal URLs, deploy paths or repository URLs into tracked files.

## This app
- Flask + SQLite, Dockerized. Container `peplink-usage-tracker`, port **5020**, bind-mount `./data:/app/data`, database `data/peplink.db`.
- Health: `GET /health`.
- On EVERY change: add a changelog entry to README.md in the existing style (version, date, one line) and keep the Features summary current — including small UI tweaks and bug fixes.
- When files change, state the docker rebuild/redeploy commands alongside the changed files.
- GB is always MB / 1024.
- Attribution rules (don't change silently): per device-day, the account with most logins that day; else last earlier login ("carried"); else next later login ("inferred"); only `[login]` events count.
- Line endings are LF. Do not introduce CRLF.
- Test locally before pushing: `python -m py_compile app/*.py`, and run the app with a scratch `DB_PATH` (never the real `data/` folder).
