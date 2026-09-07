"""CombinedScraper

ดึงผังรายการทีวีจาก API ของ NBTC แล้วเขียนขึ้น Google Sheet
(แยกชีทตามสำนักข่าว) ผ่าน Google Apps Script Web App

ลำดับการทำงาน
    1. fetch API (POST channelType=1 ได้ครบทุกช่อง)
    2. extract pgDate (DD-MM-YY), pgBeginTime, pgTitle ; ใช้ channelName เป็นชื่อชีท
    3. group ตามช่อง, เรียงตามวัน -> เวลา
    4. ต่อชีท: create(sheet) แล้ว append_programs(sheet, data)
       -- sync แบบ append-only: Apps Script เทียบกับรายการเดิมในชีท
          แล้วเพิ่ม "เฉพาะรายการใหม่" ต่อท้าย (record เดิมไม่ถูกแตะ)
    5. เขียน configuration/channels.txt = ชื่อช่องที่ไม่ซ้ำ

ไฟล์ประกอบ (path อ้างอิงจากโฟลเดอร์ที่ไฟล์ program.py อยู่ ไม่ผูกกับ cwd):
    configuration/map.txt       -- input : map ชื่อช่องจาก API -> ชื่อชีท
    configuration/channels.txt  -- output : รายชื่อชีทที่ไม่ซ้ำ (โปรแกรมสร้างให้)
    .env                        -- APPS_SCRIPT_URL (อยู่ที่ root ของโปรเจกต์)
"""

from __future__ import annotations

import json
import os
import sys
import time
from collections import defaultdict

import requests

# คอนโซล Windows ดีฟอลต์เป็น cp1252 พิมพ์ภาษาไทยไม่ได้ -> บังคับเป็น UTF-8
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass


# path อ้างอิงจากตำแหน่งไฟล์นี้เสมอ (รันจากโฟลเดอร์ไหนก็หาไฟล์เจอ)
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_CONFIG_DIR = os.path.join(_BASE_DIR, "configuration")


def _load_dotenv(path: str = "") -> None:
    """โหลดค่าจากไฟล์ .env เข้า os.environ (parser เล็ก ๆ ไม่ต้องพึ่ง python-dotenv)

    - รูปแบบ KEY=VALUE ต่อบรรทัด ; ข้ามบรรทัดว่าง / ขึ้นต้นด้วย #
    - ตัด quote ครอบ value และ prefix "export " ออกให้
    - ไม่ทับค่าที่ตั้งไว้แล้วใน environment จริง (env จริงชนะ .env)
    """
    path = path or os.path.join(_BASE_DIR, ".env")
    try:
        with open(path, encoding="utf-8-sig") as fh:
            for raw in fh:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                if line.startswith("export "):
                    line = line[len("export "):]
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except FileNotFoundError:
        pass


_load_dotenv()

# --------------------------------------------------------------------------- #
# ตั้งค่า
# --------------------------------------------------------------------------- #

API_URL = "https://dttguide.nbtc.go.th/BcsEpgDataServices/BcsEpgDataController/getProgramDataWeb"
API_PAYLOAD = {"channelType": "1"}

# URL ของ Apps Script Web App หลัง deploy (กำหนดผ่านไฟล์ .env หรือ env var APPS_SCRIPT_URL)
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "NO URL")

# ไฟล์ mapping (input) และ channels (output) ย้ายไปอยู่ในโฟลเดอร์ configuration/
CHANNELS_FILE = os.path.join(_CONFIG_DIR, "channels.txt")
MAP_FILE = os.path.join(_CONFIG_DIR, "map.txt")
# ชื่อสั้นไว้แสดงใน log (ไม่ต้องโชว์ path เต็ม)
CHANNELS_FILE_LABEL = "configuration/channels.txt"
MAP_FILE_LABEL = "configuration/map.txt"
REQUEST_TIMEOUT = 60
APPS_SCRIPT_TIMEOUT = 120  # Apps Script บางครั้งตอบช้า
MAX_RETRIES = 3
RETRY_WAIT = 5  # วินาที
SHEET_PAUSE = 1.5  # วินาที เว้นระหว่างชีท


# --------------------------------------------------------------------------- #
# ขั้นตอนที่ 1 : fetch API
# --------------------------------------------------------------------------- #

def fetch_program_data() -> dict:
    """ยิง POST ไปที่ API ของ NBTC แล้วคืน JSON ทั้งก้อน"""
    last_err: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(f"[fetch] POST {API_URL} (ครั้งที่ {attempt}/{MAX_RETRIES})")
            resp = requests.post(API_URL, json=API_PAYLOAD, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            print(f"[fetch] ผิดพลาด: {exc}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_WAIT)
            continue

        msg = data.get("responseMessage", {})
        if str(msg.get("code")) != "2000":
            raise RuntimeError(f"API ตอบกลับไม่สำเร็จ: {msg}")
        return data

    raise RuntimeError(f"fetch API ไม่สำเร็จหลังจากลอง {MAX_RETRIES} ครั้ง") from last_err


# --------------------------------------------------------------------------- #
# ขั้นตอนที่ 2 : extract
# --------------------------------------------------------------------------- #

def extract_records(data: dict) -> list[dict]:
    """ดึง list ของ record ออกจาก response

    API ตอบกลับด้วยคีย์ "results" (list of program records)
    เผื่อไว้รองรับ "result" และรูปแบบ dict ด้วย
    """
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


def _trim_time(value: str) -> str:
    """'09:00:00' -> '09:00' (เอาแค่ระดับนาที)"""
    parts = (value or "").split(":")
    if len(parts) >= 2:
        return f"{parts[0]}:{parts[1]}"
    return value or ""


def _date_sort_key(pg_date: str) -> tuple[int, int, int]:
    """'31-08-26' (DD-MM-YY) -> (2026, 8, 31) สำหรับเรียงลำดับ

    รองรับปี 4 หลักเผื่อไว้ ('31-08-2026' -> (2026, 8, 31))
    """
    try:
        day, month, year = (int(p) for p in pg_date.split("-"))
        if year < 100:
            year += 2000
        return (year, month, day)
    except (ValueError, AttributeError):
        return (9999, 99, 99)


def load_sheet_name_map() -> dict[str, str]:
    """อ่าน map.txt -> { ชื่อช่องจาก API : ชื่อชีทที่ต้องการ }

    รูปแบบแต่ละบรรทัด:
        "ชื่อจาก API"                 -> ใช้ชื่อเดิม
        "ชื่อจาก API > ชื่อชีทใหม่"    -> เปลี่ยนชื่อ
    บรรทัดว่างหรือขึ้นต้นด้วย # จะถูกข้าม ; ถ้าคีย์ซ้ำ บรรทัดหลังชนะ
    """
    mapping: dict[str, str] = {}
    if not os.path.exists(MAP_FILE):
        print(f"[map] ไม่พบ {MAP_FILE_LABEL} - ใช้ชื่อช่องจาก API ตรงๆ")
        return mapping

    with open(MAP_FILE, encoding="utf-8") as fh:
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

    print(f"[map] โหลด {len(mapping)} รายการจาก {MAP_FILE_LABEL}")
    return mapping


def build_sheets(
    records: list[dict], name_map: dict[str, str]
) -> dict[str, list[list[str]]]:
    """group ตามชื่อชีท (หลัง map) -> [[วัน, เวลา, รายการ], ...] เรียงตามวัน -> เวลา"""
    grouped: dict[str, set[tuple[str, str, str]]] = defaultdict(set)
    unmapped: set[str] = set()

    for rec in records:
        channel = (rec.get("channelName") or "").strip()
        pg_date = (rec.get("pgDate") or "").strip()
        pg_time = _trim_time((rec.get("pgBeginTime") or "").strip())
        pg_title = (rec.get("pgTitle") or "").strip()
        if not channel:
            continue
        sheet_name = name_map.get(channel, channel)
        if channel not in name_map:
            unmapped.add(channel)
        grouped[sheet_name].add((pg_date, pg_time, pg_title))

    for channel in sorted(unmapped):
        print(f"[map] เตือน: ช่อง {channel!r} ไม่มีใน {MAP_FILE_LABEL} -> ใช้ชื่อเดิม")

    seen_api = {(r.get("channelName") or "").strip() for r in records}
    for key in sorted(k for k in name_map if k not in seen_api):
        print(f"[map] เตือน: {MAP_FILE_LABEL} มี {key!r} แต่ไม่พบช่องนี้ใน API")

    sheets: dict[str, list[list[str]]] = {}
    for sheet_name, rows in grouped.items():
        ordered = sorted(rows, key=lambda r: (_date_sort_key(r[0]), r[1]))
        sheets[sheet_name] = [list(r) for r in ordered]
    return sheets


# --------------------------------------------------------------------------- #
# ขั้นตอนที่ 3 : เขียนขึ้น Google Sheet ผ่าน Apps Script
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
                    f"Apps Script ตอบกลับไม่ใช่ JSON (action={action}): {resp.text[:300]}"
                ) from exc
            if not body.get("ok"):
                raise RuntimeError(
                    f"Apps Script error (action={action}): {body.get('error')}"
                )
            return body.get("result", {})
        except (requests.RequestException, RuntimeError) as exc:
            last_err = exc
            print(f"    !! {action} ผิดพลาด (ครั้งที่ {attempt}/{MAX_RETRIES}): {exc}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_WAIT)

    raise RuntimeError(
        f"Apps Script action={action} ไม่สำเร็จหลังจากลอง {MAX_RETRIES} ครั้ง"
    ) from last_err


def push_to_sheets(sheets: dict[str, list[list[str]]]) -> None:
    total = len(sheets)
    grand_appended = 0
    grand_skipped = 0
    for index, (channel, rows) in enumerate(sorted(sheets.items()), start=1):
        print(f"[sheet {index}/{total}] {channel} ({len(rows)} รายการจาก API)")
        created = call_apps_script("create", sheet=channel)
        if created.get("created"):
            print(f"    - สร้างชีทใหม่")

        # sync แบบ append-only: Apps Script เทียบกับของเดิมในชีท แล้วเพิ่มเฉพาะรายการใหม่ต่อท้าย
        res = call_apps_script("append_programs", sheet=channel, data=rows)
        appended = res.get("appended", 0)
        skipped = res.get("skipped", 0)
        grand_appended += appended
        grand_skipped += skipped
        print(
            f"    - เพิ่มใหม่ {appended} รายการ, มีอยู่แล้ว {skipped} รายการ"
            f" (รวมในชีท {res.get('total_rows', '?')})"
        )
        time.sleep(SHEET_PAUSE)  # เว้นจังหวะ ลดโอกาส Apps Script ตอบหน้า error ชั่วคราว

    print(f"[summary] เพิ่มใหม่รวม {grand_appended} รายการ, ข้ามที่มีอยู่แล้ว {grand_skipped} รายการ")


# --------------------------------------------------------------------------- #
# ขั้นตอนที่ 4 : channels.txt
# --------------------------------------------------------------------------- #

def write_channels_file(sheets: dict[str, list[list[str]]]) -> None:
    names = sorted(sheets.keys())
    os.makedirs(_CONFIG_DIR, exist_ok=True)
    with open(CHANNELS_FILE, "w", encoding="utf-8") as fh:
        fh.write("\n".join(names) + "\n")
    print(f"[channels] เขียน {CHANNELS_FILE_LABEL} : {len(names)} ช่อง")


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #

def main() -> int:
    if not APPS_SCRIPT_URL.lower().startswith("http"):
        print("!! ยังไม่ได้ตั้งค่า APPS_SCRIPT_URL")
        print("   - ใส่ในไฟล์ .env ที่โฟลเดอร์นี้:  APPS_SCRIPT_URL=https://script.google.com/macros/s/XXXX/exec")
        print('   - หรือ set env var:  $env:APPS_SCRIPT_URL = "https://script.google.com/macros/s/XXXX/exec"')
        return 1

    data = fetch_program_data()
    records = extract_records(data)
    print(f"[extract] ได้ {len(records)} record")
    if not records:
        print("!! ไม่พบข้อมูลใน result")
        return 1

    name_map = load_sheet_name_map()
    sheets = build_sheets(records, name_map)
    print(f"[extract] แยกได้ {len(sheets)} ชีท")

    push_to_sheets(sheets)
    write_channels_file(sheets)
    print("[done] เสร็จสิ้น")
    return 0


if __name__ == "__main__":
    sys.exit(main())
