# ⚠️ Disclaimer
This repository contains software, utilities, and web interfaces<br/>
that I developed during my internship as commissioned work for **ThaiPBS (Thai Public Broadcasting Service)**<br/>
The source code included here is published with permission from my supervisor and the organization<br/>
solely for portfolio and educational purposes.<br/>

### To protect the organization's infrastructure and sensitive information:
* All credentials, API keys, access tokens, secrets, and environment-specific configurations have been completely removed.
* Internal API endpoints, URLs, sensitive data or other confidential details have been sanitized, removed.

---

# WB13-ProgramSchedule
* Multiple channels program schedule fetcher & uploader (google sheets)
* `scheduler.py` keeps `program.py` running on a self-adjusting schedule: after every fetch it looks at
  the newly fetched schedule, finds the channel whose known programs run out soonest, and sleeps until
  10 minutes before that moment — then fetches again, so no channel's guide ever goes stale.

## Run the Application Locally (Python3 is needed)
1. Clone this repository
2. `cd fetcher`
3. Run ```pip install -r requirements.txt```
4. Run ```python3 scheduler.py``` (continuous loop) or ```python3 program.py``` (single fetch; also
   supports `--purge`)

## Run with Docker
```
docker compose up --build fetcher
```
Runs `scheduler.py` continuously. For a one-off fetch or `--purge`, override the command, e.g.
```
docker compose run --rm fetcher python program.py --purge 1-9-2026
```
See `fetcher/` for the service source (`program.py`, `scheduler.py`, `Dockerfile`, `configuration/`).