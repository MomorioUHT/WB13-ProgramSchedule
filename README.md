# ⚠️ Disclaimer
This repository contains software, utilities, and web interfaces<br/>
that I developed during my internship as commissioned work for **ThaiPBS (Thai Public Broadcasting Service)**<br/>
The source code included here is published with permission from my supervisor and the organization<br/>
solely for portfolio and educational purposes.<br/>

### To protect the organization's infrastructure and sensitive information:
* All credentials, API keys, access tokens, secrets, and environment-specific configurations have been completely removed.
* Internal API endpoints, URLs, sensitive data or other confidential details have been sanitized, removed.

---

# WB13-ProgramSchedule (Fetcher only)
* Multiple channels program schedule fetcher & uploader (google sheets)
* `scheduler.py` keeps `program.py` running on a self-adjusting schedule: after every fetch it looks at
  the newly fetched schedule, finds the channel whose known programs run out soonest, and sleeps until
  10 minutes before that moment — then fetches again, so no channel's guide ever goes stale.

## Run the Application Locally (Python3 is needed)
This is the normal way to run it:
1. Clone this repository
2. `cd fetcher`
3. Install dependencies: `pip install -r requirements.txt`
4. Copy `fetcher/.env.example` to `fetcher/.env` and fill in `GSHEET_URL` (see the comment in that file:
   deploy `webdashboard/App.gs` as a Google Apps Script Web App on your Google Sheet, then use the
   deployment URL ending in `/exec`). All variables in `.env.example` are required, no fallback/default.
5. Run `python scheduler.py` — fetches immediately, then keeps looping on a self-adjusting schedule
   (see the note above on how it picks the next fetch time)

### Purge old records
Instead of `scheduler.py`, run `program.py` directly with `--purge DATE` (`DD-MM-YYYY`). This deletes
matching rows immediately — there is no dry-run/preview mode.
```
python program.py --purge 1-9-2026                    # purge every mapped sheet
python program.py --purge 1-9-2026 --sheet "Thai PBS"  # only the given sheet(s)
```
Running `python program.py` with no flags does a single one-off fetch (no loop, no purge).

## Run with Docker
`fetcher/docker-compose.yml` builds and runs the same `scheduler.py` loop inside a container.
It also needs `fetcher/.env` from step 4 above (loaded via `env_file`).
```
cd fetcher
docker compose up --build
```
For a one-off fetch or `--purge` instead of the continuous loop, override the command, e.g.
```
docker compose run --rm fetcher python program.py --purge 1-9-2026
```
See `fetcher/` for the service source (`program.py`, `scheduler.py`, `Dockerfile`, `configuration/`).