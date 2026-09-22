from __future__ import annotations

import os
import sys
sys.dont_write_bytecode = True
import time
import logging
from datetime import datetime, timedelta

import program

# --------------------------------------------------------------------------- #
# Settings
# --------------------------------------------------------------------------- #

# required so datetime.now() reflects local time via the OS clock (container TZ env var)
os.environ["TZ"]
if hasattr(time, "tzset"):
    time.tzset()

LEAD_TIME = timedelta(minutes=10)
FALLBACK_WAIT = timedelta(minutes=30)
POLL_INTERVAL_SECONDS = 300

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)

log = logging.getLogger("scheduler")


# --------------------------------------------------------------------------- #
# Wait loop
# --------------------------------------------------------------------------- #

def sleep_until(target: datetime) -> None:
    while True:
        remaining = (target - datetime.now()).total_seconds()
        if remaining <= 0:
            return
        time.sleep(min(remaining, POLL_INTERVAL_SECONDS))


def main() -> int:
    # no state persists across restarts, so every start fetches immediately before waiting
    log.info("Startup: fetching immediately, ignoring any previous schedule")
    while True:
        min_latest = program.run_fetch_cycle()

        if min_latest is None:
            next_run = datetime.now() + FALLBACK_WAIT
            log.warning(
                "Could not determine minimum latest datetime, retrying at %s",
                next_run.isoformat(sep=" "),
            )
        else:
            next_run = min_latest - LEAD_TIME
            log.info(
                "Minimum latest datetime across sheets: %s -> next fetch at %s",
                min_latest.isoformat(sep=" "), next_run.isoformat(sep=" "),
            )

        sleep_until(next_run)


if __name__ == "__main__":
    sys.exit(main())
