#!/usr/bin/env python3
"""Export Armor public-API gateway calls for one customer and one UTC day to .xlsx.

Runs queries/armor_calls_export.dp through the Coralogix DataPrime API (Archive tier).
If the day returns 10,000 rows or a bytesScannedLimitWarning, it is split into 6-hour
windows and merged, deduped on armor_request_id.

The API key is read from $CORALOGIX_API_KEY, or from the JSON file named by
$CORALOGIX_KEY_FILE (field apiKey.keyValue). It is never printed.

    python armor_export.py --business-number 169699783343081 --key-prefix Ty6aI0lN --day 2026-10-11
"""
import argparse
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import Workbook

URL = "https://api.cx498.coralogix.com/api/v1/dataprime/query"
QUERY_FILE = Path(__file__).parent / "queries" / "armor_calls_export.dp"
ROW_LIMIT = 10000
IST = timedelta(hours=5, minutes=30)
COLUMNS = ["timestamp", "applicationName", "api", "armor_request_id", "client_ip",
           "full_uri", "method", "status", "time_ms", "trace_id"]


def load_key():
    key = os.environ.get("CORALOGIX_API_KEY")
    if not key and os.environ.get("CORALOGIX_KEY_FILE"):
        key = json.load(open(os.environ["CORALOGIX_KEY_FILE"]))["apiKey"]["keyValue"]
    if not key:
        sys.exit("Set CORALOGIX_API_KEY (or CORALOGIX_KEY_FILE pointing at the key JSON).")
    return key


def build_query(business_number, key_prefix):
    # Both values go inside DataPrime string literals, so keep them to a safe alphabet.
    if not re.fullmatch(r"\d{12,18}", business_number):
        sys.exit("--business-number must be 12-18 digits")
    if not re.fullmatch(r"[A-Za-z0-9]{4,16}", key_prefix):
        sys.exit("--key-prefix must be 4-16 alphanumeric characters")
    return (QUERY_FILE.read_text().strip()
            .replace("{{BUSINESS_NUMBER}}", business_number)
            .replace("{{KEY_PREFIX}}", key_prefix))


def iso(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def run_window(key, query, start, end):
    """One API call. Returns (rows, warning/error lines)."""
    r = requests.post(
        URL,
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        json={"query": query,
              "metadata": {"syntax": "QUERY_SYNTAX_DATAPRIME", "tier": "TIER_ARCHIVE",
                           "startDate": iso(start), "endDate": iso(end)}},
        timeout=300)
    if r.status_code != 200:
        sys.exit(f"HTTP {r.status_code} for {iso(start)}..{iso(end)}: {r.text[:500]}")
    rows, notes = [], []
    for line in r.text.splitlines():
        if not line.strip():
            continue
        o = json.loads(line)
        if "warning" in o or "error" in o:
            notes.append(o)
        for res in (o.get("result") or {}).get("results", []):
            rows.append(json.loads(res["userData"]))
    return rows, notes


def hit_limit(rows, notes):
    return len(rows) >= ROW_LIMIT or any("bytesScannedLimitWarning" in json.dumps(n) for n in notes)


def fetch_day(key, query, day_start):
    day_end = day_start + timedelta(days=1)
    rows, notes = run_window(key, query, day_start, day_end)
    print(f"full-day run: {len(rows)} rows, {len(notes)} warning/error lines")
    if not hit_limit(rows, notes):
        return rows, notes

    print("row limit or scan limit hit -> splitting into 6-hour windows")
    for n in notes:
        print("WARNING/ERROR LINE (full-day run, superseded by the split):", json.dumps(n)[:600])
    rows, notes = [], []
    for i in range(4):
        s = day_start + timedelta(hours=6 * i)
        e = s + timedelta(hours=6)
        w_rows, w_notes = run_window(key, query, s, e)
        print(f"  {iso(s)}..{iso(e)}: {len(w_rows)} rows, {len(w_notes)} warning/error lines")
        if hit_limit(w_rows, w_notes):
            print("  !! this window still hit a limit - results are PARTIAL; split further")
        rows += w_rows
        notes += w_notes
    seen, merged = set(), []
    for r in rows:
        rid = r.get("armor_request_id")
        if rid is not None:
            if rid in seen:
                continue
            seen.add(rid)
        merged.append(r)
    print(f"merged {len(rows)} -> {len(merged)} rows after dedupe on armor_request_id")
    return merged, notes


def write_xlsx(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "armor_logs"
    ws.append(COLUMNS)
    for r in rows:
        status = r.get("status")
        ws.append([r.get("timestamp"), r.get("applicationName"), r.get("api"),
                   r.get("armor_request_id"), r.get("client_ip"), r.get("full_uri"),
                   r.get("method"), int(status) if str(status).isdigit() else status,
                   r.get("time_ms"), r.get("trace_id")])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--business-number", required=True)
    ap.add_argument("--key-prefix", required=True, help="first 8 chars of the customer API key")
    ap.add_argument("--day", required=True, help="UTC day, YYYY-MM-DD")
    ap.add_argument("--out", help="default: out/<business-number>_<day>.xlsx")
    ap.add_argument("--expected-rows", type=int, help="flag the run if the total is >5%% off")
    args = ap.parse_args(argv)

    day_start = datetime.strptime(args.day, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    out = args.out or f"out/{args.business_number}_{args.day}.xlsx"
    query = build_query(args.business_number, args.key_prefix)
    key = load_key()

    s, e = day_start, day_start + timedelta(days=1)
    print(f"range UTC: {iso(s)} -> {iso(e)} | IST: {(s + IST):%d/%m/%Y %H:%M} -> {(e + IST):%d/%m/%Y %H:%M}")
    rows, notes = fetch_day(key, query, day_start)
    for n in notes:
        print("WARNING/ERROR LINE:", json.dumps(n)[:600])

    rows.sort(key=lambda r: r.get("timestamp") or "")
    write_xlsx(rows, out)

    print(f"\nTOTAL ROWS: {len(rows)}")
    if args.expected_rows and abs(len(rows) - args.expected_rows) > 0.05 * args.expected_rows:
        print(f"!! FLAG: more than 5% away from the expected {args.expected_rows}")
    print("\napi x status:")
    counts = Counter((r.get("api"), r.get("status")) for r in rows)
    for (api, st), n in sorted(counts.items(), key=lambda kv: (-kv[1], str(kv[0]))):
        print(f"{n:>7}  {api}  {st}")
    print(f"\nfile: {out}")


if __name__ == "__main__":
    main()
