from __future__ import annotations

import os
import sys
sys.dont_write_bytecode = True
import time
import dotenv

import argparse
import json
import logging
import requests

from collections import defaultdict
from datetime import datetime

# --------------------------------------------------------------------------- #
# Systems & Path Configuration
# --------------------------------------------------------------------------- #
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_CONFIG_DIR = os.path.join(_BASE_DIR, "configuration")

CHANNELS_FILE = os.path.join(_CONFIG_DIR, "channels.txt")
CHANNELS_FILE_LABEL = "configuration/channels.txt"

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("fetcher")

dotenv.load_dotenv(os.path.join(_BASE_DIR, ".env"))

# --------------------------------------------------------------------------- #
# Settings
# --------------------------------------------------------------------------- #

REQUEST_TIMEOUT = int(os.environ["REQUEST_TIMEOUT"])
APPS_SCRIPT_TIMEOUT = int(os.environ["APPS_SCRIPT_TIMEOUT"])
MAX_RETRIES = int(os.environ["MAX_RETRIES"])
RETRY_WAIT = int(os.environ["RETRY_WAIT"])

API_URL = os.environ["DTT_URL"]
API_PAYLOAD = {"channelType": "1"}

APPS_SCRIPT_URL = os.environ["GSHEET_URL"]

# --------------------------------------------------------------------------- #
# Fetch DTT's API
# --------------------------------------------------------------------------- #
def fetch_program_data() -> dict:
    last_err: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(API_URL, json=API_PAYLOAD, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            last_err = exc
            log.warning("Fetch attempt %d/%d failed: %s", attempt, MAX_RETRIES, exc)
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_WAIT)
            continue

        msg = data.get("responseMessage", {})
        if str(msg.get("code")) != "2000":
            raise RuntimeError(f"NBTC API returned an error: {msg}")
        return data

    raise RuntimeError(f"Fetch failed after {MAX_RETRIES} attempts") from last_err


# --------------------------------------------------------------------------- #
# Extract Data
# --------------------------------------------------------------------------- #
def extract_records(data: dict) -> list[dict]:
    result = data.get("results")
    if result is None:
        result = data.get("result")
    if result is None:
        return []
    if isinstance(result, list):
        return result
    if isinstance(result, dict):
        if "channelName" in result:
            return [result]
        for value in result.values():
            if isinstance(value, list):
                return value
    return []


def trim_time(value: str) -> str:
    # '09:00:00' -> '09:00'
    parts = (value or "").split(":")
    if len(parts) >= 2:
        return f"{parts[0]}:{parts[1]}"
    return value or ""


def date_sort_key(pg_date: str) -> tuple[int, int, int]:
    #'31-08-26' (DD-MM-YY) -> (2026, 8, 31) for sorting, also supports 2 and 4 year digits
    try:
        day, month, year = (int(p) for p in pg_date.split("-"))
        if year < 100:
            year += 2000
        return (year, month, day)
    except (ValueError, AttributeError):
        return (9999, 99, 99)


def load_sheet_name_map() -> dict[str, str]:
    # Channel name remapping with configuration/channels.txt and '>' splitter
    mapping: dict[str, str] = {}
    if not os.path.exists(CHANNELS_FILE):
        log.info("%s not found, using API channel names as-is", CHANNELS_FILE_LABEL)
        return mapping

    with open(CHANNELS_FILE, encoding="utf-8") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if ">" in line:
                src, dst = (part.strip() for part in line.split(">", 1))
            else:
                src = dst = line
            if not src:
                continue
            mapping[src] = dst or src

    log.debug("Loaded %d channel mappings from %s", len(mapping), CHANNELS_FILE_LABEL)
    return mapping


def build_sheets(
    records: list[dict], name_map: dict[str, str]
) -> dict[str, list[list[str]]]:
    # Sheet groupping
    grouped: dict[str, set[tuple[str, str, str]]] = defaultdict(set)
    unmapped: set[str] = set()

    for rec in records:
        channel = (rec.get("channelName") or "").strip()
        pg_date = (rec.get("pgDate") or "").strip()
        pg_time = trim_time((rec.get("pgBeginTime") or "").strip())
        pg_title = (rec.get("pgTitle") or "").strip()
        if not channel:
            continue
        sheet_name = name_map.get(channel, channel)
        if channel not in name_map:
            unmapped.add(channel)
        grouped[sheet_name].add((pg_date, pg_time, pg_title))

    if unmapped:
        log.warning(
            "%d channel(s) not in %s, using API name as-is: %s",
            len(unmapped), CHANNELS_FILE_LABEL, ", ".join(sorted(unmapped)),
        )

    seen_api = {(r.get("channelName") or "").strip() for r in records}
    stale_keys = sorted(k for k in name_map if k not in seen_api)
    if stale_keys:
        log.warning(
            "%d mapping(s) in %s not found in this API response: %s",
            len(stale_keys), CHANNELS_FILE_LABEL, ", ".join(stale_keys),
        )

    sheets: dict[str, list[list[str]]] = {}
    for sheet_name, rows in grouped.items():
        ordered = sorted(rows, key=lambda r: (date_sort_key(r[0]), r[1]))
        sheets[sheet_name] = [list(r) for r in ordered]
    return sheets


def parse_program_datetime(pg_date: str, pg_time: str) -> datetime | None:
    # 'DD-MM-YY(YY)' + 'HH:MM' -> naive local datetime (schedule times are always Thailand local time)
    year, month, day = date_sort_key(pg_date)
    if (year, month, day) == (9999, 99, 99):
        return None
    parts = trim_time(pg_time).split(":")
    if len(parts) < 2:
        return None
    try:
        hour, minute = int(parts[0]), int(parts[1])
        return datetime(year, month, day, hour, minute)
    except ValueError:
        return None


def compute_min_latest_datetime(sheets: dict[str, list[list[str]]]) -> datetime | None:
    # earliest "latest scheduled" datetime across sheets - the soonest deadline for the next fetch
    latest_per_sheet: list[datetime] = []
    for rows in sheets.values():
        if not rows:
            continue
        pg_date, pg_time, _ = rows[-1]
        dt = parse_program_datetime(pg_date, pg_time)
        if dt is not None:
            latest_per_sheet.append(dt)
    if not latest_per_sheet:
        return None
    return min(latest_per_sheet)


# --------------------------------------------------------------------------- #
# Write to Google Sheets
# --------------------------------------------------------------------------- #

def call_apps_script(action: str, **params) -> dict:
    payload = {"action": action, **params}
    last_err: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(
                APPS_SCRIPT_URL, json=payload, timeout=APPS_SCRIPT_TIMEOUT
            )
            resp.raise_for_status()
            try:
                body = resp.json()
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"Apps Script response not JSON (action={action}): {resp.text[:300]}"
                ) from exc
            if not body.get("ok"):
                raise RuntimeError(
                    f"Apps Script error (action={action}): {body.get('error')}"
                )
            return body.get("result", {})
        except (requests.RequestException, RuntimeError) as exc:
            last_err = exc
            log.warning("%s attempt %d/%d failed: %s", action, attempt, MAX_RETRIES, exc)
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_WAIT)

    raise RuntimeError(
        f"Apps Script action={action} failed after {MAX_RETRIES} attempts"
    ) from last_err


def push_to_sheets(sheets: dict[str, list[list[str]]]) -> None:
    # single batched request for every sheet - Apps Script loops internally (was 2 calls/sheet before)
    total = len(sheets)
    log.info("Syncing %d sheet(s) to Google Sheets in one batch request...", total)
    res = call_apps_script("sync_programs_batch", sheets=sheets)
    results = res.get("results", {})

    grand_appended = 0
    grand_skipped = 0
    for index, name in enumerate(sorted(results), start=1):
        r = results[name]
        if "error" in r:
            log.error("[%d/%d] %s failed: %s", index, total, name, r["error"])
            continue
        appended = r.get("appended", 0)
        skipped = r.get("skipped", 0)
        grand_appended += appended
        grand_skipped += skipped
        log.info(
            "[%d/%d] %s: +%d new, %d existing, %d skipped_old (new_sheet=%s, total=%s)",
            index, total, name, appended, skipped, r.get("skipped_old", 0),
            r.get("created", False), r.get("total_rows", "?"),
        )

    failed = res.get("failed", [])
    log.info(
        "Sync done: %d sheets, +%d new rows, %d already existed",
        total, grand_appended, grand_skipped,
    )
    if failed:
        log.warning("%d sheet(s) failed: %s", len(failed), ", ".join(failed))


# --------------------------------------------------------------------------- #
# --purge mode
# --------------------------------------------------------------------------- #

def normalize_purge_date(raw: str) -> str:
    #'1-9-2026' / '01-09-26' / '2026-09-01' -> 'YYYY-MM-DD'
    year, month, day = date_sort_key(raw)
    if (year, month, day) == (9999, 99, 99) or not (1 <= month <= 12 and 1 <= day <= 31):
        raise SystemExit(f"Invalid date: {raw!r} - use DD-MM-YYYY, e.g. 1-9-2026")
    return f"{year:04d}-{month:02d}-{day:02d}"


def resolve_target_sheets(sheet_args: list[str]) -> list[str]:
    # merge sheet names from --sheet (repeatable / comma-separated); default to every sheet in channels.txt
    names: list[str] = []
    for arg in sheet_args:
        names.extend(part.strip() for part in arg.split(",") if part.strip())
    if names:
        # de-dupe while preserving order
        seen: set[str] = set()
        return [n for n in names if not (n in seen or seen.add(n))]

    names = sorted(set(load_sheet_name_map().values()))
    if not names:
        raise SystemExit(
            "No target sheet found - specify --sheet \"NAME\" "
            f"or add channel mappings to {CHANNELS_FILE_LABEL}"
        )
    return names


def purge_old_records(cutoff_raw: str, sheet_names: list[str]) -> int:
    # single batched request for every sheet via Apps Script action=purge_before_batch
    cutoff = normalize_purge_date(cutoff_raw)
    total = len(sheet_names)
    log.info(
        "Purge: removing rows older than %s from %d sheet(s) in one batch request",
        cutoff, total,
    )

    res = call_apps_script("purge_before_batch", sheets=sheet_names, cutoff=cutoff)
    results = res.get("results", {})

    grand_removed = 0
    failed: list[str] = []
    for index, name in enumerate(sheet_names, start=1):
        r = results.get(name, {})
        if "error" in r:
            log.error("[%d/%d] %s failed: %s", index, total, name, r["error"])
            failed.append(name)
            continue

        removed = r.get("removed", 0)
        kept = r.get("kept", 0)
        grand_removed += removed
        log.info("[%d/%d] %s: %d removed, %d kept", index, total, name, removed, kept)
        for sample in r.get("removed_sample", [])[:3]:
            cols = list(sample) + ["", "", ""]
            log.debug("    sample removed: %s %s %s", cols[0], cols[1], cols[2])

    log.info(
        "Purge done: %d row(s) removed across %d sheet(s)",
        grand_removed, total - len(failed),
    )
    if failed:
        log.warning("%d sheet(s) failed: %s", len(failed), ", ".join(failed))
        return 1
    return 0


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #

def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="program.py",
        description=(
            "Fetch the TV schedule from NBTC and sync it to Google Sheets (no arguments); "
            "or use --purge to remove rows older than a given date from the sheets"
        ),
    )
    parser.add_argument(
        "--purge",
        metavar="DATE",
        help=(
            "Purge mode: remove rows older than DATE (format DD-MM-YYYY, e.g. 1-9-2026) "
            "from the target sheets and shift the remaining rows up"
        ),
    )
    parser.add_argument(
        "--sheet",
        action="append",
        default=[],
        metavar="NAME",
        help=(
            "Target sheet(s) for --purge (repeatable or comma-separated); "
            f"defaults to every sheet mapped in {CHANNELS_FILE_LABEL}"
        ),
    )
    return parser


def run_fetch_cycle() -> datetime | None:
    # fetch + sync one cycle; returns the next-fetch deadline, or None if no records were found
    log.info("Fetching program data from NBTC API...")
    data = fetch_program_data()
    records = extract_records(data)
    if not records:
        log.error("No records found in API response")
        return None

    name_map = load_sheet_name_map()
    sheets = build_sheets(records, name_map)
    log.info("Fetched %d records -> %d sheets", len(records), len(sheets))

    push_to_sheets(sheets)
    log.info("Done")
    return compute_min_latest_datetime(sheets)


def main(argv: list[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)

    if args.purge:
        sheet_names = resolve_target_sheets(args.sheet)
        return purge_old_records(args.purge, sheet_names)

    if args.sheet:
        log.warning("--sheet is only used with --purge, ignoring in scrape mode")

    return 0 if run_fetch_cycle() is not None else 1


if __name__ == "__main__":
    sys.exit(main())
